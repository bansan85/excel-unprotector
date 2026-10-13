from __future__ import annotations

import os
import zipfile
import re
import warnings
from core.vba_grep import VbaMatch, find_protect_calls


def unprotect_xlsx(input_path: str, output_path: str) -> list[VbaMatch]:
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

    # Opening the output for writing would truncate the input while it is
    # read. samefile compares the files themselves, so a path differing only
    # by case on a case-insensitive file system, a symlink or a hard link to
    # the input is caught too.
    if os.path.exists(output_path) and os.path.samefile(input_path, output_path):
        raise ValueError("The output file must not be the input file")

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

    # The conversion is done: a macro that cannot be parsed must not fail it.
    try:
        return find_protect_calls(output_path)
    except Exception as e:
        warnings.warn(f"Unable to scan VBA macros for Protect calls: {e}")
        return []
