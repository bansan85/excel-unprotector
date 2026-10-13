import unittest
import os
from core.excel_unprotector import unprotect_xlsx
from zipfile import ZipFile
import hashlib
import pathlib


class TestExcelUnprotect(unittest.TestCase):
    def setUp(self):
        self.current_folder = pathlib.Path(__file__).parent.resolve()
        self.input_file = f"{self.current_folder}/data/book1.xlsm"
        self.output_file = f"{self.current_folder}/data/book1_unprotected.xlsm"

    @staticmethod
    def get_zip_inventory(zip_path: str):
        inventory: dict[str, str] = {}
        with ZipFile(zip_path, "r") as z:
            for member in z.infolist():
                if member.is_dir():
                    continue

                content = z.read(member.filename)
                file_hash = hashlib.sha256(content).hexdigest()
                inventory[member.filename] = file_hash
        return inventory

    @staticmethod
    def compare_xlsx_content(input_file: str, output_file: str):
        inv1 = TestExcelUnprotect.get_zip_inventory(input_file)
        inv2 = TestExcelUnprotect.get_zip_inventory(output_file)

        files1 = set(inv1.keys())
        files2 = set(inv2.keys())

        if files1 != files2:
            return False

        differences: list[str] = []
        for filename in files1:
            if inv1[filename] != inv2[filename]:
                differences.append(filename)

        if differences:
            return False

        return True

    def test_integrity(self):
        self.assertEqual(unprotect_xlsx(self.input_file, self.output_file), [])

        self.assertFalse(self.compare_xlsx_content(self.input_file, self.output_file))
        self.assertTrue(
            self.compare_xlsx_content(
                self.output_file,
                f"{self.current_folder}/data/book1_unprotected_ref.xlsm",
            )
        )


if __name__ == "__main__":
    unittest.main()
