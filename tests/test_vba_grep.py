import unittest
import pathlib
import struct
from zipfile import ZipFile
from unittest import mock
from core.vba_grep import (
    CompoundFile,
    VbaMatch,
    codepage_codec,
    decompress,
    extract_vba_modules,
    find_protect_calls,
    grep_vba_modules,
)


def directory_entry_offsets(data: bytes) -> list[int]:
    # File offset of each 128-byte compound file directory entry
    cfb = CompoundFile(data)
    offsets: list[int] = []
    sector = struct.unpack_from("<I", data, 0x30)[0]
    while sector < 0xFFFFFFFA:
        for entry in range(cfb.sector_size // 128):
            offsets.append((sector + 1) * cfb.sector_size + entry * 128)
        sector = cfb.fat[sector]
    return offsets


class TestVbaGrep(unittest.TestCase):
    def setUp(self):
        self.current_folder = pathlib.Path(__file__).parent.resolve()
        self.input_file = f"{self.current_folder}/data/book1.xlsm"

    def test_decompress_spec_examples(self):
        # [MS-OVBA] 3.2.1 No Compression Example
        self.assertEqual(
            decompress(
                bytes.fromhex(
                    "0119B00061626364656667680069"
                    "6A6B6C6D6E6F7000717273747576"
                    "2E"
                )
            ),
            b"abcdefghijklmnopqrstuv.",
        )
        # [MS-OVBA] 3.2.3 Example with Compression
        self.assertEqual(
            decompress(
                bytes.fromhex(
                    "012FB000236161616263646582660070"
                    "61676869"
                    "6A013808616B6C00306D6E6F70067102"
                    "70041072737475761077787"
                    "97A00"
                    "3C"
                )
            ),
            b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa",
        )

    def test_extract_vba_modules(self):
        with ZipFile(self.input_file, "r") as z:
            modules = extract_vba_modules(z.read("xl/vbaProject.bin"))

        self.assertEqual(
            [name for name, _ in modules], ["Module1", "ThisWorkbook"]
        )
        self.assertIn("Sub Macro1()", modules[0][1])

    def test_codepage_codec(self):
        # Code page 10000 (Mac Roman) has another name in Python
        self.assertEqual(codepage_codec(10000), "mac_roman")
        self.assertEqual(b"\x8e".decode(codepage_codec(10000)), "é")
        self.assertEqual(codepage_codec(1252), "cp1252")
        self.assertEqual(b"\xe9".decode(codepage_codec(28591)), "é")
        # No code page has this number, Windows included
        self.assertEqual(codepage_codec(12345), "latin_1")

    def test_extract_vba_modules_unknown_codepage(self):
        with ZipFile(self.input_file, "r") as z:
            vba_project = z.read("xl/vbaProject.bin")
        # PROJECTCODEPAGE record (id 3, size 2): code page 1252, then 12345
        cp1252 = b"\x03\x00\x02\x00\x00\x00\xe4\x04"
        cp12345 = b"\x03\x00\x02\x00\x00\x00\x39\x30"
        real_decompress = decompress
        dir_stream = real_decompress(CompoundFile(vba_project).stream("VBA/dir"))
        self.assertEqual(dir_stream.count(cp1252), 1)
        expected = extract_vba_modules(vba_project)

        def with_unknown_codepage(data: bytes) -> bytes:
            return real_decompress(data).replace(cp1252, cp12345)

        with mock.patch("core.vba_grep.decompress", side_effect=with_unknown_codepage):
            modules = extract_vba_modules(vba_project)

        self.assertEqual(modules, expected)

    def test_grep_protect_calls(self):
        modules = [
            (
                "Module1",
                "Sub Lock()\r\n"
                '    ActiveWorkbook.Protect Password:="secret"\r\n'
                "    Sheets(1).Unprotect\r\n"
                "    ' ActiveSheet.Protect\r\n"
                "    x = ActiveSheet.ProtectContents\r\n"
                "End Sub\r\n",
            ),
            ("Sheet1", 'Private Sub Worksheet_Activate()\r\n    Protect "pwd"\r\n'),
        ]

        self.assertEqual(
            grep_vba_modules(modules),
            [
                VbaMatch("Module1", 2, 'ActiveWorkbook.Protect Password:="secret"'),
                VbaMatch("Sheet1", 2, 'Protect "pwd"'),
            ],
        )

    def test_grep_skips_strings_and_comments(self):
        modules = [
            (
                "Module1",
                'MsgBox "Please Protect the sheet"\r\n'
                "x = 1 ' Protect later\r\n"
                "Rem\tProtect\r\n"
                "y = 2: Rem Protect\r\n"
                "' Comment going on _\r\n"
                "    over Protect\r\n"
                '    ActiveSheet.Protect Password:="it\'s"\r\n'
                's = "a""b": Sheets(1).Protect\r\n',
            ),
        ]

        self.assertEqual(
            grep_vba_modules(modules),
            [
                VbaMatch("Module1", 7, 'ActiveSheet.Protect Password:="it\'s"'),
                VbaMatch("Module1", 8, 's = "a""b": Sheets(1).Protect'),
            ],
        )

    def test_cyclic_fat_chain(self):
        with ZipFile(self.input_file, "r") as z:
            data = bytearray(z.read("xl/vbaProject.bin"))
        sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        dir_start = struct.unpack_from("<I", data, 0x30)[0]
        first_fat_sector = struct.unpack_from("<I", data, 0x4C)[0]
        # Link the first sector of the directory to itself
        struct.pack_into(
            "<I",
            data,
            (first_fat_sector + 1) * sector_size + 4 * dir_start,
            dir_start,
        )

        with self.assertRaises(ValueError):
            CompoundFile(bytes(data))

    def test_truncated_stream(self):
        with ZipFile(self.input_file, "r") as z:
            data = z.read("xl/vbaProject.bin")
        # Cut the file after the first byte of the last piece of PROJECT. The
        # header, FAT and directory come before the cut: the file still parses.
        offset, _ = CompoundFile(data).stream_spans("PROJECT")[-1]
        cfb = CompoundFile(data[: offset + 1])

        with self.assertRaises(ValueError):
            cfb.stream_spans("PROJECT")

    def test_ignores_size_high_bits_in_version_3(self):
        with ZipFile(self.input_file, "r") as z:
            data = bytearray(z.read("xl/vbaProject.bin"))
        self.assertEqual(struct.unpack_from("<H", data, 0x1A)[0], 3)
        original = bytes(data)
        # Fill the high 32 bits of every directory entry's stream size
        for offset in directory_entry_offsets(original):
            struct.pack_into("<I", data, offset + 124, 0xDEADBEEF)

        self.assertEqual(
            extract_vba_modules(bytes(data)), extract_vba_modules(original)
        )

    def test_same_stream_name_in_other_storage(self):
        with ZipFile(self.input_file, "r") as z:
            data = bytearray(z.read("xl/vbaProject.bin"))
        original = CompoundFile(bytes(data))
        # Rename VBA/_VBA_PROJECT to VBA/PROJECT, the name of a root stream
        old_name = "_VBA_PROJECT\0".encode("utf-16-le")
        new_name = "PROJECT\0".encode("utf-16-le")
        for offset in directory_entry_offsets(bytes(data)):
            if data[offset : offset + len(old_name)] == old_name:
                data[offset : offset + 64] = new_name.ljust(64, b"\0")
                struct.pack_into("<H", data, offset + 64, len(new_name))

        cfb = CompoundFile(bytes(data))
        self.assertEqual(cfb.stream("PROJECT"), original.stream("PROJECT"))
        self.assertEqual(
            cfb.stream("VBA/PROJECT"), original.stream("VBA/_VBA_PROJECT")
        )
        self.assertEqual(
            extract_vba_modules(bytes(data)), extract_vba_modules(original.data)
        )

    def test_find_protect_calls(self):
        self.assertEqual(find_protect_calls(self.input_file), [])


if __name__ == "__main__":
    unittest.main()
