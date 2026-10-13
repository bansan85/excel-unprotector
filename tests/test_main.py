import unittest
import pathlib
import tempfile
import zipfile
from contextlib import redirect_stdout
from io import StringIO
from unittest import mock

import main


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


if __name__ == "__main__":
    unittest.main()
