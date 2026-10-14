import unittest
import pathlib
import tempfile
import zipfile
from contextlib import redirect_stdout
from io import BytesIO, StringIO, TextIOWrapper
from unittest import mock

import main
from core.vba_grep import VbaMatch


class TestMain(unittest.TestCase):
    def setUp(self):
        self.current_folder = pathlib.Path(__file__).parent.resolve()
        self.input_file = f"{self.current_folder}/data/book1.xlsm"

    def test_creates_missing_output_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_file = pathlib.Path(tmp) / "new.xlsm"
            argv = ["excel-unprotector", self.input_file, str(output_file)]
            with mock.patch("sys.argv", argv), redirect_stdout(StringIO()):
                main.main()

            self.assertTrue(zipfile.is_zipfile(output_file))

    def test_prints_unencodable_vba_line(self):
        # cp1252, as a redirected stdout uses on Western Windows, has no Cyrillic
        stdout = TextIOWrapper(BytesIO(), encoding="cp1252")
        match = VbaMatch("Module1", 3, 'ActiveSheet.Protect Password:="пароль"')
        with tempfile.TemporaryDirectory() as tmp:
            output_file = pathlib.Path(tmp) / "new.xlsm"
            argv = ["excel-unprotector", self.input_file, str(output_file)]
            with mock.patch("sys.stdout", stdout), mock.patch("sys.argv", argv):
                with mock.patch("main.unprotect_xlsx", return_value=[match]):
                    main.main()
            stdout.flush()

        output = stdout.buffer.getvalue()
        self.assertIn(b"Module1:3:", output)
        # The first letter of "пароль" is escaped, not dropped or replaced
        self.assertIn(b"\\u043f", output)


if __name__ == "__main__":
    unittest.main()
