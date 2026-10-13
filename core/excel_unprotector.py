import zipfile
import re


def unprotect_xlsx(input_path: str, output_path: str):
    patterns_to_remove = [
        r"<sheetProtection[^>]*/>",
        r"<workbookProtection[^>]*/>",
        r"<protectedRanges[^>]*/>",
        r"<protectedRanges[^>]*>.*?</protectedRanges>",
    ]

    attrs_to_remove = [
        r'\s+windowProtection="[^"]*"',
        r'\s+applyProtection="[^"]*"',
        r'\s+state="hidden"',
    ]

    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(
            output_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)

                if item.filename.endswith(".xml"):
                    text = content.decode("utf-8", errors="ignore")

                    for pattern in patterns_to_remove:
                        text = re.sub(pattern, "", text, flags=re.DOTALL)

                    for attr_pattern in attrs_to_remove:
                        text = re.sub(attr_pattern, "", text)

                    zout.writestr(item.filename, text.encode("utf-8"))

                elif item.filename == "xl/vbaProject.bin":
                    new_content = content.replace(b"DPB=", b"DPX=")
                    zout.writestr(item.filename, new_content)

                else:
                    zout.writestr(item.filename, content)
