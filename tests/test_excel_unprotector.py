import unittest
import os
from core.excel_unprotector import unprotect_xlsx
from core.vba_grep import CompoundFile
from zipfile import ZipFile
import codecs
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

    def test_shows_very_hidden_sheets(self):
        workbook = (
            "<workbook><sheets>"
            '<sheet name="a" sheetId="1" state="veryHidden"/>'
            '<sheet name="b" sheetId="2" state="hidden"/>'
            '<sheet name="c" sheetId="3" state="visible"/>'
            "</sheets></workbook>"
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("xl/workbook.xml", workbook)

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("xl/workbook.xml").decode(),
                    "<workbook><sheets>"
                    '<sheet name="a" sheetId="1"/>'
                    '<sheet name="b" sheetId="2"/>'
                    '<sheet name="c" sheetId="3" state="visible"/>'
                    "</sheets></workbook>",
                )

    def test_removes_prefixed_elements(self):
        workbook = (
            '<x:workbook xmlns:x="ns"><x:workbookProtection lockStructure="1"/>'
            "</x:workbook>"
        )
        sheet = (
            '<x:worksheet xmlns:x="ns"><x:sheetProtection sheet="1"/>'
            '<x:protectedRanges><x:protectedRange name="r" sqref="A1"/>'
            "</x:protectedRanges><x:protectedRanges/></x:worksheet>"
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("xl/workbook.xml", workbook)
                z.writestr("xl/worksheets/sheet1.xml", sheet)

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("xl/workbook.xml").decode(),
                    '<x:workbook xmlns:x="ns"></x:workbook>',
                )
                self.assertEqual(
                    z.read("xl/worksheets/sheet1.xml").decode(),
                    '<x:worksheet xmlns:x="ns"></x:worksheet>',
                )

    def test_removes_elements_with_end_tag(self):
        workbook = (
            '<workbook><workbookProtection lockStructure="1"></workbookProtection>'
            "<sheets/></workbook>"
        )
        sheet = (
            '<x:worksheet xmlns:x="ns"><x:sheetProtection sheet="1" password="CC3D">'
            '</x:sheetProtection ><x:sheetData/><sheetProtectionInfo a="1"/>'
            "</x:worksheet>"
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("xl/workbook.xml", workbook)
                z.writestr("xl/worksheets/sheet1.xml", sheet)

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("xl/workbook.xml").decode(),
                    "<workbook><sheets/></workbook>",
                )
                self.assertEqual(
                    z.read("xl/worksheets/sheet1.xml").decode(),
                    '<x:worksheet xmlns:x="ns"><x:sheetData/>'
                    '<sheetProtectionInfo a="1"/></x:worksheet>',
                )

    def test_removes_password_to_modify(self):
        workbook = (
            '<workbook><fileVersion appName="xl"/>'
            '<fileSharing readOnlyRecommended="1" userName="u"'
            ' reservationPassword="CC3D"/><workbookPr/></workbook>'
        )
        workbook2 = (
            '<x:workbook xmlns:x="ns"><x:fileSharing userName="u" '
            'algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" '
            'spinCount="100000"></x:fileSharing><x:workbookPr/></x:workbook>'
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("xl/workbook.xml", workbook)
                z.writestr("xl/workbook2.xml", workbook2)

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("xl/workbook.xml").decode(),
                    '<workbook><fileVersion appName="xl"/><workbookPr/></workbook>',
                )
                self.assertEqual(
                    z.read("xl/workbook2.xml").decode(),
                    '<x:workbook xmlns:x="ns"><x:workbookPr/></x:workbook>',
                )

    def test_keeps_xml_encoding(self):
        xml = (
            '<?xml version="1.0" encoding="{}"?>'
            '<worksheet><sheetProtection sheet="1"/><t>é</t></worksheet>'
        )
        expected = '<?xml version="1.0" encoding="{}"?><worksheet><t>é</t></worksheet>'
        parts = {
            "utf16le.xml": codecs.BOM_UTF16_LE
            + xml.format("UTF-16").encode("utf-16-le"),
            "utf16be.xml": xml.format("UTF-16").encode("utf-16-be"),
            "latin1.xml": xml.format("ISO-8859-1").encode("latin-1"),
        }
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                for name, content in parts.items():
                    z.writestr(name, content)

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("utf16le.xml"),
                    codecs.BOM_UTF16_LE
                    + expected.format("UTF-16").encode("utf-16-le"),
                )
                self.assertEqual(
                    z.read("utf16be.xml"),
                    expected.format("UTF-16").encode("utf-16-be"),
                )
                self.assertEqual(
                    z.read("latin1.xml"),
                    expected.format("ISO-8859-1").encode("latin-1"),
                )

    def test_keeps_undecodable_xml_unchanged(self):
        content = b'<worksheet><sheetProtection sheet="1"/><t>\xff</t></worksheet>'
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("xl/worksheets/sheet1.xml", content)

            with self.assertWarns(UserWarning):
                unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(z.read("xl/worksheets/sheet1.xml"), content)

    def test_patches_only_project_password(self):
        with ZipFile(self.input_file, "r") as z:
            parts = {name: z.read(name) for name in z.namelist()}
        vba_project = bytearray(parts["xl/vbaProject.bin"])
        # Write "DPB=" at the start of another stream
        offset, _ = CompoundFile(bytes(vba_project)).stream_spans("PROJECTwm")[0]
        vba_project[offset : offset + 4] = b"DPB="
        parts["xl/vbaProject.bin"] = bytes(vba_project)

        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsm"
            output_file = f"{tmp}/out.xlsm"
            with ZipFile(input_file, "w") as z:
                for name, content in parts.items():
                    z.writestr(name, content)

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                cfb = CompoundFile(z.read("xl/vbaProject.bin"))
        self.assertEqual(cfb.stream("PROJECTwm")[:4], b"DPB=")
        self.assertIn(b"\r\nDPX=", cfb.stream("PROJECT"))
        self.assertNotIn(b"DPB=", cfb.stream("PROJECT"))

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
