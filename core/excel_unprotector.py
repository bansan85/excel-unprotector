from __future__ import annotations

import os
import zipfile
import re
import warnings
from core.vba_grep import VbaMatch, find_protect_calls

# A start or empty-element tag. Text content cannot contain "<" unescaped, so
# this never matches inside a cell value or a shared string.
TAG_PATTERN = re.compile(r"<[A-Za-z_][^\"'>]*(?:(?:\"[^\"]*\"|'[^']*')[^\"'>]*)*>")


def unprotect_xlsx(input_path: str, output_path: str) -> list[VbaMatch]:
    # Some writers put the elements in a prefixed namespace, such as
    # <x:sheetProtection/>: the closing tag must then carry the same prefix.
    patterns_to_remove = [
        r"<(?:[\w.-]+:)?sheetProtection[^>]*/>",
        r"<(?:[\w.-]+:)?workbookProtection[^>]*/>",
        r"<(?:[\w.-]+:)?protectedRanges[^>]*/>",
        r"<((?:[\w.-]+:)?)protectedRanges[^>]*>.*?</\1protectedRanges>",
    ]

    attrs_to_remove = [
        r'\s+windowProtection="[^"]*"',
        r'\s+applyProtection="[^"]*"',
        r'\s+state="hidden"',
    ]

    def remove_attrs(tag: re.Match) -> str:
        text = tag.group(0)
        for attr_pattern in attrs_to_remove:
            text = re.sub(attr_pattern, "", text)
        return text

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

                    text = TAG_PATTERN.sub(remove_attrs, text)

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
