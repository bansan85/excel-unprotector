from __future__ import annotations

import unittest
import os
from excel_unprotector.unprotector import unprotect_xlsx
from excel_unprotector.vba_grep import CompoundFile
from zipfile import ZipFile, is_zipfile
import codecs
import hashlib
import pathlib
import shutil
import tempfile
import warnings
from unittest import mock

CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"


def write_package(
    path: str,
    parts: dict[str, str | bytes],
    content_types: dict[str, str] | None = None,
):
    # Each part gets an Override: a worksheet unless content_types says otherwise
    content_types = content_types or {}
    overrides = "".join(
        f'<Override PartName="/{name}" '
        f'ContentType="{content_types.get(name, WORKSHEET)}"/>'
        for name in parts
    )
    types = f'<Types xmlns="{CONTENT_TYPES_NS}">{overrides}</Types>'
    with ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", types)
        for name, content in parts.items():
            z.writestr(name, content)


class TestExcelUnprotect(unittest.TestCase):
    def setUp(self):
        self.current_folder = pathlib.Path(__file__).parent.resolve()
        self.input_file = f"{self.current_folder}/data/book1.xlsm"

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
        data_folder = f"{self.current_folder}/data"
        before = sorted(os.listdir(data_folder))
        with tempfile.TemporaryDirectory() as tmp:
            output_file = f"{tmp}/out.xlsm"
            self.assertEqual(unprotect_xlsx(self.input_file, output_file), [])

            self.assertFalse(self.compare_xlsx_content(self.input_file, output_file))
            self.assertTrue(
                self.compare_xlsx_content(
                    output_file,
                    f"{self.current_folder}/data/book1_unprotected_ref.xlsm",
                )
            )
        # Nothing else may be written into the sample data folder
        self.assertEqual(sorted(os.listdir(data_folder)), before)

    def test_keeps_text_content(self):
        xml = (
            '<workbook><bookViews><workbookView windowProtection="1"/>'
            '</bookViews><sheets><sheet name="a" state="hidden"/></sheets>'
            '<t>x state="hidden" applyProtection="1" y</t></workbook>'
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            write_package(input_file, {"xl/workbook.xml": xml})

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
            write_package(input_file, {"xl/workbook.xml": workbook})

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

    def test_removes_attributes_in_any_quoting(self):
        workbook = (
            "<workbook><sheets><sheet name='a' state='hidden'/>"
            '<sheet name="b" state = "veryHidden"/></sheets>'
            "<bookViews><workbookView windowProtection = '1'/></bookViews>"
            '<definedNames><definedName comment=\' state="hidden" \' name="n">'
            "A1</definedName></definedNames></workbook>"
        )
        expected = (
            "<workbook><sheets><sheet name='a'/>"
            '<sheet name="b"/></sheets>'
            "<bookViews><workbookView/></bookViews>"
            '<definedNames><definedName comment=\' state="hidden" \' name="n">'
            "A1</definedName></definedNames></workbook>"
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            write_package(input_file, {"xl/workbook.xml": workbook})

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(z.read("xl/workbook.xml").decode(), expected)

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
            write_package(
                input_file,
                {"xl/workbook.xml": workbook, "xl/worksheets/sheet1.xml": sheet},
            )

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
            write_package(
                input_file,
                {"xl/workbook.xml": workbook, "xl/worksheets/sheet1.xml": sheet},
            )

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
            write_package(
                input_file,
                {"xl/workbook.xml": workbook, "xl/workbook2.xml": workbook2},
            )

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
            write_package(input_file, parts)

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

    def test_keeps_xml_encoding_with_spaces(self):
        xml = (
            "<?xml version=\"1.0\" encoding = 'ISO-8859-1'?>"
            '<worksheet><sheetProtection sheet="1"/><t>é</t></worksheet>'
        )
        expected = (
            "<?xml version=\"1.0\" encoding = 'ISO-8859-1'?>"
            "<worksheet><t>é</t></worksheet>"
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            write_package(
                input_file, {"xl/worksheets/sheet1.xml": xml.encode("latin-1")}
            )

            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter("always")
                unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("xl/worksheets/sheet1.xml"), expected.encode("latin-1")
                )
        self.assertEqual(caught, [])

    def test_keeps_utf8_byte_order_mark(self):
        xml = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet><sheetProtection sheet="1"/><t>café</t></worksheet>'
        )
        expected = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            "<worksheet><t>café</t></worksheet>"
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            write_package(
                input_file,
                {"xl/worksheets/sheet1.xml": codecs.BOM_UTF8 + xml.encode("utf-8")},
            )

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                content = z.read("xl/worksheets/sheet1.xml")
        self.assertTrue(content.startswith(codecs.BOM_UTF8))
        self.assertEqual(content, codecs.BOM_UTF8 + expected.encode("utf-8"))

    def test_keeps_undecodable_xml_unchanged(self):
        content = b'<worksheet><sheetProtection sheet="1"/><t>\xff</t></worksheet>'
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            write_package(input_file, {"xl/worksheets/sheet1.xml": content})

            with self.assertWarns(UserWarning):
                unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(z.read("xl/worksheets/sheet1.xml"), content)

    def test_processes_only_spreadsheet_parts(self):
        content_types = (
            f'<Types xmlns="{CONTENT_TYPES_NS}">'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Default Extension="SHEET" ContentType="{WORKSHEET.upper()}"/>'
            '<Override PartName="/XL/Worksheets/Sheet1.xml" '
            f'ContentType="{WORKSHEET}"/>'
            '<Override PartName="/docProps/custom.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.custom-properties+xml"/>'
            "</Types>"
        )
        sheet = '<worksheet><sheetProtection sheet="1"/></worksheet>'
        custom = '<tasks><task name="x" state="hidden"/><sheetProtection/></tasks>'
        props = '<Properties><property state="hidden"/></Properties>'
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("[Content_Types].xml", content_types)
                # Matched by its Override, although the case differs
                z.writestr("xl/worksheets/sheet1.xml", sheet)
                # Matched by the Default of its extension, although the case
                # differs in the extension and in the content type
                z.writestr("xl/worksheets/sheet2.sheet", sheet)
                z.writestr("customXml/item1.xml", custom)
                z.writestr("docProps/custom.xml", props)

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("xl/worksheets/sheet1.xml").decode(),
                    "<worksheet></worksheet>",
                )
                self.assertEqual(
                    z.read("xl/worksheets/sheet2.sheet").decode(),
                    "<worksheet></worksheet>",
                )
                self.assertEqual(z.read("customXml/item1.xml").decode(), custom)
                self.assertEqual(z.read("docProps/custom.xml").decode(), props)

    def test_ignores_elements_other_than_default_and_override(self):
        content_types = (
            f'<Types xmlns="{CONTENT_TYPES_NS}">'
            '<Default Extension="xml" ContentType="application/xml"/>'
            # Neither a Default nor an Override: its attributes must not count
            # as either, or the XML parts below would become worksheets
            '<Unknown Extension="xml" PartName="/docProps/custom.xml" '
            f'ContentType="{WORKSHEET}"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            f'ContentType="{WORKSHEET}"/>'
            "</Types>"
        )
        sheet = '<worksheet><sheetProtection sheet="1"/></worksheet>'
        custom = '<tasks><task name="x" state="hidden"/><sheetProtection/></tasks>'
        props = '<Properties><property state="hidden"/></Properties>'
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("[Content_Types].xml", content_types)
                z.writestr("xl/worksheets/sheet1.xml", sheet)
                z.writestr("customXml/item1.xml", custom)
                z.writestr("docProps/custom.xml", props)

            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter("always")
                unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(
                    z.read("xl/worksheets/sheet1.xml").decode(),
                    "<worksheet></worksheet>",
                )
                self.assertEqual(z.read("customXml/item1.xml").decode(), custom)
                self.assertEqual(z.read("docProps/custom.xml").decode(), props)
        self.assertEqual(caught, [])

    def test_keeps_cdata_comments_and_processing_instructions(self):
        xml = (
            '<worksheet><!-- <sheetProtection/> <a state="hidden"/> -->'
            '<?pi <sheetProtection/>?><sheetProtection sheet="1"/>'
            '<t><![CDATA[<sheetProtection/><b state="hidden"/>]]></t></worksheet>'
        )
        expected = (
            '<worksheet><!-- <sheetProtection/> <a state="hidden"/> -->'
            "<?pi <sheetProtection/>?>"
            '<t><![CDATA[<sheetProtection/><b state="hidden"/>]]></t></worksheet>'
        )
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            write_package(input_file, {"xl/worksheets/sheet1.xml": xml})

            unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(z.read("xl/worksheets/sheet1.xml").decode(), expected)

    def test_processes_all_xml_without_content_types(self):
        sheet = '<worksheet><sheetProtection sheet="1"/></worksheet>'
        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsx"
            output_file = f"{tmp}/out.xlsx"
            with ZipFile(input_file, "w") as z:
                z.writestr("a.xml", sheet)

            with self.assertWarns(UserWarning):
                unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(z.read("a.xml").decode(), "<worksheet></worksheet>")

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

    def test_truncated_vba_project(self):
        with ZipFile(self.input_file, "r") as z:
            parts = {name: z.read(name) for name in z.namelist()}
        # Cut the file after the first byte of the last piece of PROJECT: the
        # header, FAT and directory are kept, but the piece runs past the end
        vba_project = parts["xl/vbaProject.bin"]
        offset, _ = CompoundFile(vba_project).stream_spans("PROJECT")[-1]
        truncated = vba_project[: offset + 1]
        self.assertIn(b"DPB=", truncated)
        parts["xl/vbaProject.bin"] = truncated

        with tempfile.TemporaryDirectory() as tmp:
            input_file = f"{tmp}/in.xlsm"
            output_file = f"{tmp}/out.xlsm"
            with ZipFile(input_file, "w") as z:
                for name, content in parts.items():
                    z.writestr(name, content)

            with self.assertWarns(UserWarning):
                unprotect_xlsx(input_file, output_file)

            with ZipFile(output_file, "r") as z:
                self.assertEqual(set(z.namelist()), set(parts))
                self.assertEqual(
                    z.read("xl/vbaProject.bin"), truncated.replace(b"DPB=", b"DPX=")
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

    def test_failure_keeps_existing_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_file = f"{tmp}/out.xlsm"
            pathlib.Path(output_file).write_bytes(b"previous")

            with mock.patch(
                "excel_unprotector.unprotector.disable_vba_password",
                side_effect=RuntimeError("boom"),
            ):
                with self.assertRaises(RuntimeError):
                    unprotect_xlsx(self.input_file, output_file)

            self.assertEqual(pathlib.Path(output_file).read_bytes(), b"previous")
            self.assertEqual(os.listdir(tmp), ["out.xlsm"])

    def test_success_leaves_no_temporary_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_file = f"{tmp}/out.xlsm"

            unprotect_xlsx(self.input_file, output_file)

            self.assertEqual(os.listdir(tmp), ["out.xlsm"])
            self.assertTrue(is_zipfile(output_file))

    def test_writes_through_symlinked_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            real_file = f"{tmp}/real.xlsm"
            link_file = f"{tmp}/link.xlsm"
            pathlib.Path(real_file).write_bytes(b"previous")
            try:
                os.symlink(real_file, link_file)
            except OSError:
                self.skipTest("symbolic links need a privilege on Windows")

            unprotect_xlsx(self.input_file, link_file)

            self.assertTrue(os.path.islink(link_file))
            self.assertTrue(is_zipfile(real_file))
            self.assertEqual(sorted(os.listdir(tmp)), ["link.xlsm", "real.xlsm"])


if __name__ == "__main__":
    unittest.main()
