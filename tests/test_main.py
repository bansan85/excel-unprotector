import unittest
import pathlib
import subprocess
import sys
import tempfile
import zipfile
from contextlib import redirect_stdout
from io import BytesIO, StringIO, TextIOWrapper
from unittest import mock

from excel_unprotector import cli
from excel_unprotector.vba_grep import VbaMatch


class TestMain(unittest.TestCase):
    def setUp(self):
        self.current_folder = pathlib.Path(__file__).parent.resolve()
        self.input_file = f"{self.current_folder}/data/book1.xlsm"

    def test_creates_missing_output_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_file = pathlib.Path(tmp) / "new.xlsm"
            argv = ["excel-unprotector", self.input_file, str(output_file)]
            with mock.patch("sys.argv", argv), redirect_stdout(StringIO()):
                cli.main()

            self.assertTrue(zipfile.is_zipfile(output_file))

    def test_prints_unencodable_vba_line(self):
        # cp1252, as a redirected stdout uses on Western Windows, has no Cyrillic
        stdout = TextIOWrapper(BytesIO(), encoding="cp1252")
        match = VbaMatch("Module1", 3, 'ActiveSheet.Protect Password:="пароль"')
        with tempfile.TemporaryDirectory() as tmp:
            output_file = pathlib.Path(tmp) / "new.xlsm"
            argv = ["excel-unprotector", self.input_file, str(output_file)]
            with mock.patch("sys.stdout", stdout), mock.patch("sys.argv", argv):
                with mock.patch(
                    "excel_unprotector.cli.unprotect_xlsx", return_value=[match]
                ):
                    cli.main()
            stdout.flush()

        output = stdout.buffer.getvalue()
        self.assertIn(b"Module1:3:", output)
        # The first letter of "пароль" is escaped, not dropped or replaced
        self.assertIn(b"\\u043f", output)

    def test_exits_when_input_is_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_file = pathlib.Path(tmp) / "missing.xlsm"
            output_file = pathlib.Path(tmp) / "new.xlsm"
            argv = ["excel-unprotector", str(input_file), str(output_file)]
            stdout = StringIO()
            with mock.patch("sys.argv", argv), redirect_stdout(stdout):
                with self.assertRaises(SystemExit) as cm:
                    cli.main()

        self.assertEqual(cm.exception.code, 1)
        self.assertEqual(stdout.getvalue(), f"Error: '{input_file}' is missing.\n")

    def test_exits_when_input_is_not_a_zip(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_file = pathlib.Path(tmp) / "not-a-workbook.xlsm"
            input_file.write_bytes(b"not a zip file")
            output_file = pathlib.Path(tmp) / "new.xlsm"
            argv = ["excel-unprotector", str(input_file), str(output_file)]
            stdout = StringIO()
            with mock.patch("sys.argv", argv), redirect_stdout(stdout):
                with self.assertRaises(SystemExit) as cm:
                    cli.main()

            self.assertFalse(output_file.exists())

        self.assertEqual(cm.exception.code, 1)
        self.assertTrue(stdout.getvalue().startswith("Unexpected exception: "))

    def test_runs_as_module(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_file = pathlib.Path(tmp) / "new.xlsm"
            subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "excel_unprotector",
                    self.input_file,
                    str(output_file),
                ],
                cwd=self.current_folder.parent,
                check=True,
                capture_output=True,
            )

            self.assertTrue(zipfile.is_zipfile(output_file))


if __name__ == "__main__":
    unittest.main()
