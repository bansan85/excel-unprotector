import unittest
import pathlib
from zipfile import ZipFile
from core.vba_grep import (
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

    def test_find_protect_calls(self):
        self.assertEqual(find_protect_calls(self.input_file), [])


if __name__ == "__main__":
    unittest.main()
