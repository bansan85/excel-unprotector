import unittest
import os
from core.excel_unprotector import unprotect_xlsx
from zipfile import ZipFile
import hashlib
import pathlib
import shutil
import tempfile


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

    def test_keeps_text_content(self):
        xml = (
            '<workbook><bookViews><workbookView windowProtection="1"/>'
            '</bookViews><sheets><sheet name="a" state="hidden"/></sheets>'
            '<t>x state="hidden" applyProtection="1" y</t></workbook>'
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("xl/workbook.xml", xml)

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("xl/workbook.xml").decode(),
                    "<workbook><bookViews><workbookView/></bookViews>"
                    '<sheets><sheet name="a"/></sheets>'
                    '<t>x state="hidden" applyProtection="1" y</t></workbook>',
                )

    def assert_refuses_same_file(self, input_file: str, output_file: str):
        original = pathlib.Path(input_file).read_bytes()
        with self.assertRaises(ValueError):
            unprotect_xlsx(input_file, output_file)
        self.assertEqual(pathlib.Path(input_file).read_bytes(), original)

    def test_refuses_same_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_file = shutil.copy(self.input_file, f"{tmp}/book.xlsm")
            self.assert_refuses_same_file(input_file, input_file)

    def test_refuses_same_path_with_other_case(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_file = shutil.copy(self.input_file, f"{tmp}/book.xlsm")
            output_file = f"{tmp}/BOOK.xlsm"
            if not os.path.exists(output_file):
                self.skipTest("case-sensitive file system")
            self.assert_refuses_same_file(input_file, output_file)

    def test_refuses_hard_link_to_input(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_file = shutil.copy(self.input_file, f"{tmp}/book.xlsm")
            output_file = f"{tmp}/link.xlsm"
            os.link(input_file, output_file)
            self.assert_refuses_same_file(input_file, output_file)


if __name__ == "__main__":
    unittest.main()
