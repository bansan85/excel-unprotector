import unittest
import pathlib
import random
import struct
from zipfile import ZipFile
from unittest import mock
from excel_unprotector.vba_grep import (
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


def move_last_fat_sector_to_difat(data: bytearray) -> int:
    # Move the last FAT sector index of the header into a DIFAT sector appended
    # at the end of the file, chained from header 0x44. Returns its index.
    sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
    last_slot = max(
        slot
        for slot in range(109)
        if struct.unpack_from("<I", data, 0x4C + 4 * slot)[0] < 0xFFFFFFFA
    )
    fat_sector = struct.unpack_from("<I", data, 0x4C + 4 * last_slot)[0]
    struct.pack_into("<I", data, 0x4C + 4 * last_slot, 0xFFFFFFFF)

    # Sector i starts at (i + 1) * sector_size, the header being sector -1
    new_index = len(data) // sector_size - 1
    values = [0xFFFFFFFF] * (sector_size // 4)
    values[0] = fat_sector
    values[-1] = 0xFFFFFFFE  # ENDOFCHAIN: this is the last DIFAT sector
    data += struct.pack(f"<{len(values)}I", *values)
    struct.pack_into("<I", data, 0x44, new_index)
    struct.pack_into("<I", data, 0x48, 1)
    return new_index


ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


def make_directory_entry(
    name: str,
    kind: int,
    black: bool,
    left: int,
    right: int,
    child: int,
    start: int,
    size: int,
) -> bytes:
    # [MS-CFB] 2.6.1: a 128-byte directory entry
    entry = bytearray(128)
    if name:
        encoded = (name + "\0").encode("utf-16-le")
        entry[: len(encoded)] = encoded
        struct.pack_into("<H", entry, 64, len(encoded))
    entry[66] = kind
    entry[67] = 1 if black else 0
    struct.pack_into("<3I", entry, 68, left, right, child)
    struct.pack_into("<IQ", entry, 116, start, size)
    return bytes(entry)


def build_compound_file(streams: dict[str, bytes], major_version: int = 3) -> bytes:
    # Minimal [MS-CFB] file: one FAT sector, one directory sector, no mini
    # stream and no mini FAT. Each stream is at least 4096 bytes, so it is in
    # regular sectors. The streams are the children of the root entry, in a
    # red-black tree: up to 3 streams fit, the middle name being the black
    # child and the other ones red.
    assert 1 <= len(streams) <= 3
    assert all(len(content) >= 4096 for content in streams.values())
    shift = 9 if major_version == 3 else 12
    sector_size = 1 << shift
    ints_per_sector = sector_size // 4

    # Sector 0 is the FAT, sector 1 the directory, then the streams in order
    fat = [FREESECT] * ints_per_sector
    fat[0] = FATSECT
    fat[1] = ENDOFCHAIN
    starts: dict[str, int] = {}
    sector = 2
    for name, content in streams.items():
        count = (len(content) + sector_size - 1) // sector_size
        starts[name] = sector
        fat[sector : sector + count] = range(sector + 1, sector + count + 1)
        fat[sector + count - 1] = ENDOFCHAIN
        sector += count
    assert sector <= ints_per_sector, "the streams must fit in one FAT sector"

    # Entry 0 is the root, entries 1 to n the streams sorted by name length
    # then name, as [MS-CFB] 2.6.4 orders the siblings
    names = sorted(streams, key=lambda name: (len(name), name.upper()))
    root_child = (len(names) + 1) // 2
    entries = [
        make_directory_entry(
            "Root Entry", 5, True, NOSTREAM, NOSTREAM, root_child, ENDOFCHAIN, 0
        )
    ]
    for index, name in enumerate(names, start=1):
        left = right = NOSTREAM
        black = index == root_child
        if black:
            left = index - 1 if index > 1 else NOSTREAM
            right = index + 1 if index < len(names) else NOSTREAM
        entries.append(
            make_directory_entry(
                name,
                2,
                black,
                left,
                right,
                NOSTREAM,
                starts[name],
                len(streams[name]),
            )
        )
    unused = make_directory_entry("", 0, True, NOSTREAM, NOSTREAM, NOSTREAM, 0, 0)
    entries += [unused] * (sector_size // 128 - len(entries))

    header = bytearray(sector_size)
    header[:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    # Minor version 0x3E, major version, byte order 0xFFFE, sector shift and
    # mini sector shift 6
    struct.pack_into("<5H", header, 0x18, 0x3E, major_version, 0xFFFE, shift, 6)
    struct.pack_into("<I", header, 0x28, 0 if major_version == 3 else 1)
    struct.pack_into("<I", header, 0x2C, 1)  # FAT sectors
    struct.pack_into("<I", header, 0x30, 1)  # first directory sector
    struct.pack_into("<I", header, 0x38, 4096)  # mini stream cutoff
    struct.pack_into("<I", header, 0x3C, ENDOFCHAIN)  # no mini FAT
    struct.pack_into("<I", header, 0x44, ENDOFCHAIN)  # no DIFAT sector
    difat = [0] + [FREESECT] * 108
    struct.pack_into("<109I", header, 0x4C, *difat)

    data = bytearray(header)
    data += struct.pack(f"<{ints_per_sector}I", *fat)
    data += b"".join(entries)
    for content in streams.values():
        data += content + b"\0" * (-len(content) % sector_size)
    return bytes(data)


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

    def test_decompress_rejects_invalid_signature(self):
        # [MS-OVBA] 2.4.1.1.1: the SignatureByte is always 0x01
        chunk = struct.pack("<H", 0x3FFF) + b"x" * 4096
        with self.assertRaises(ValueError):
            decompress(b"")
        with self.assertRaises(ValueError):
            decompress(b"\x00" + chunk)

    def test_decompress_uncompressed_chunk(self):
        # [MS-OVBA] 2.4.1.1.5: header 0x3FFF, compressed flag cleared, so the
        # 4096 bytes after it are copied as is
        data = bytes(range(256)) * 16
        self.assertEqual(
            decompress(b"\x01" + struct.pack("<H", 0x3FFF) + data), data
        )
        # The compressed chunk after it must start right after the 4096 bytes
        compressed = b"\x00" + b"abcdefgh"  # flags 0x00: 8 literals
        container = (
            b"\x01"
            + struct.pack("<H", 0x3FFF)
            + data
            + struct.pack("<H", 0xB000 | (len(compressed) - 1))
            + compressed
        )
        self.assertEqual(decompress(container), data + b"abcdefgh")

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

        with mock.patch(
            "excel_unprotector.vba_grep.decompress", side_effect=with_unknown_codepage
        ):
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

    def test_grep_line_numbers_match_vba_editor(self):
        # The editor hides the Attribute lines and does not break lines at form
        # feeds: the numbers must be those of the lines it shows
        source = "\r\n".join(
            [
                'Attribute VB_Name = "Sheet1"',
                'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"',
                "Attribute VB_GlobalNameSpace = False",
                "Sub Lock()",
                'Attribute Lock.VB_ProcData.VB_Invoke_Func = " \\n14"',
                "    ' form feed \x0c in a comment",
                "    ActiveSheet.Protect",
                "End Sub",
                "Sub Protect()",
                'Attribute Protect.VB_Description = "Locks"',
                "End Sub",
                "' comment going on _",
                "Attribute is a word of the comment here",
                "Protect",
            ]
        )

        self.assertEqual(
            grep_vba_modules([("Sheet1", source)]),
            [
                VbaMatch("Sheet1", 3, "ActiveSheet.Protect"),
                VbaMatch("Sheet1", 5, "Sub Protect()"),
                VbaMatch("Sheet1", 9, "Protect"),
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

    def test_difat_chain(self):
        # [MS-CFB] 2.5: the FAT sector indexes past the 109 in the header are in
        # the DIFAT sectors chained from header 0x44
        with ZipFile(self.input_file, "r") as z:
            original = z.read("xl/vbaProject.bin")
        data = bytearray(original)
        difat_index = move_last_fat_sector_to_difat(data)
        self.assertEqual(struct.unpack_from("<I", data, 0x44)[0], difat_index)

        cfb = CompoundFile(bytes(data))
        expected = CompoundFile(original)
        self.assertEqual(cfb.fat, expected.fat)
        self.assertEqual(cfb.streams, expected.streams)
        self.assertEqual(cfb.stream("PROJECT"), expected.stream("PROJECT"))
        self.assertEqual(
            extract_vba_modules(bytes(data)), extract_vba_modules(original)
        )

    def test_cyclic_difat_chain(self):
        with ZipFile(self.input_file, "r") as z:
            data = bytearray(z.read("xl/vbaProject.bin"))
        sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        difat_index = move_last_fat_sector_to_difat(data)
        # Link the DIFAT sector to itself through its last slot
        struct.pack_into(
            "<I", data, (difat_index + 1) * sector_size + sector_size - 4, difat_index
        )

        with self.assertRaisesRegex(ValueError, "Cyclic sector chain"):
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

    def test_keeps_size_high_bits_in_version_4(self):
        # [MS-CFB] 2.6.3: the high 32 bits of the size are only ignored in
        # version 3, where 4096-byte sectors are not allowed
        content = random.Random(0).randbytes(5000)
        data = bytearray(build_compound_file({"Big": content}, major_version=4))
        self.assertEqual(struct.unpack_from("<H", data, 0x1A)[0], 4)
        cfb = CompoundFile(bytes(data))
        self.assertEqual(cfb.sector_size, 4096)
        self.assertEqual(cfb.stream("Big"), content)

        # Fill the high 32 bits of every directory entry's stream size
        start, size = cfb.streams["big"]
        for offset in directory_entry_offsets(bytes(data)):
            struct.pack_into("<I", data, offset + 124, 0xDEADBEEF)

        cfb = CompoundFile(bytes(data))
        self.assertEqual(cfb.streams["big"], (start, (0xDEADBEEF << 32) | size))

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

    def test_first_entry_must_be_root_storage(self):
        with ZipFile(self.input_file, "r") as z:
            original = z.read("xl/vbaProject.bin")
        # The first directory entry is the root storage, of object type 5
        root_offset = directory_entry_offsets(original)[0]
        self.assertEqual(original[root_offset + 66], 5)

        for entry_type in (0, 1):
            with self.subTest(entry_type=entry_type):
                data = bytearray(original)
                data[root_offset + 66] = entry_type
                with self.assertRaisesRegex(ValueError, "Missing root entry"):
                    CompoundFile(bytes(data))

    def test_cyclic_directory_tree(self):
        with ZipFile(self.input_file, "r") as z:
            original = z.read("xl/vbaProject.bin")
        offsets = directory_entry_offsets(original)
        # Entries: 1 is the VBA storage, its child 3 is ThisWorkbook, whose left
        # sibling 2 is Module1; 7 is PROJECT, the child of the root entry
        cases = {
            # The child of the VBA storage is the storage itself
            "child to self": (1, 76, 1),
            # The right sibling of ThisWorkbook is PROJECT, already visited
            "sibling to root child": (3, 72, 7),
            # The right sibling of Module1 is its storage VBA
            "sibling to ancestor": (2, 72, 1),
        }
        for case, (index, field, target) in cases.items():
            with self.subTest(case):
                data = bytearray(original)
                struct.pack_into("<I", data, offsets[index] + field, target)
                with self.assertRaisesRegex(
                    ValueError, "Cyclic directory tree in compound file"
                ):
                    CompoundFile(bytes(data))

    def test_stream_in_regular_sectors(self):
        # [MS-CFB] 2.6.1: a stream of at least 4096 bytes is in regular sectors,
        # here through the FAT since the file has no mini FAT
        content = random.Random(0).randbytes(5000)
        cfb = CompoundFile(build_compound_file({"Big": content}))
        self.assertEqual(cfb.minifat, [])
        self.assertEqual(cfb.stream("Big"), content)
        # Sectors 2 to 11 hold the stream, the last one only 392 bytes
        self.assertEqual(
            cfb.stream_spans("Big"),
            [((i + 3) * 512, 512) for i in range(9)] + [(12 * 512, 392)],
        )

    def test_sector_chain_longer_than_size(self):
        # Big has 10 sectors, but its size field is lowered to 9 sectors: the
        # last sector of the chain is spare and must not be read
        content = random.Random(0).randbytes(5000)
        data = bytearray(build_compound_file({"Big": content}))
        offset = directory_entry_offsets(bytes(data))[1]
        self.assertEqual(struct.unpack_from("<I", data, offset + 120)[0], 5000)
        struct.pack_into("<I", data, offset + 120, 4608)

        cfb = CompoundFile(bytes(data))
        self.assertEqual(cfb.stream("Big"), content[:4608])
        # Sectors 2 to 10 hold the stream, sector 11 is not in any span
        self.assertEqual(
            cfb.stream_spans("Big"),
            [((i + 3) * 512, 512) for i in range(9)],
        )

    def test_find_protect_calls(self):
        self.assertEqual(find_protect_calls(self.input_file), [])


if __name__ == "__main__":
    unittest.main()
