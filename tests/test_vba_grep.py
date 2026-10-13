import unittest
import pathlib
import struct
from zipfile import ZipFile
from core.vba_grep import (
    CompoundFile,
    VbaMatch,
    decompress,
    extract_vba_modules,
    find_protect_calls,
    grep_vba_modules,
)


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

    def test_ignores_size_high_bits_in_version_3(self):
        with ZipFile(self.input_file, "r") as z:
            data = bytearray(z.read("xl/vbaProject.bin"))
        self.assertEqual(struct.unpack_from("<H", data, 0x1A)[0], 3)
        cfb = CompoundFile(bytes(data))
        sector = struct.unpack_from("<I", data, 0x30)[0]
        # Fill the high 32 bits of every directory entry's stream size
        while sector < 0xFFFFFFFA:
            for entry in range(cfb.sector_size // 128):
                offset = (sector + 1) * cfb.sector_size + entry * 128
                struct.pack_into("<I", data, offset + 124, 0xDEADBEEF)
            sector = cfb.fat[sector]

        self.assertEqual(
            extract_vba_modules(bytes(data)),
            extract_vba_modules(bytes(cfb.data)),
        )

    def test_find_protect_calls(self):
        self.assertEqual(find_protect_calls(self.input_file), [])


if __name__ == "__main__":
    unittest.main()
