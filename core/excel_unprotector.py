from __future__ import annotations

import codecs
import os
import zipfile
import re
import warnings
from core.vba_grep import CompoundFile, VbaMatch, find_protect_calls

# A start or empty-element tag. Text content cannot contain "<" unescaped, so
# this never matches inside a cell value or a shared string.
TAG_PATTERN = re.compile(r"<[A-Za-z_][^\"'>]*(?:(?:\"[^\"]*\"|'[^']*')[^\"'>]*)*>")

ENCODING_PATTERN = re.compile(rb"<\?xml[^>]*encoding=[\"']([A-Za-z0-9._-]+)")


def xml_encoding(content: bytes) -> str:
    # [XML 1.0] appendix F: a byte order mark or the first characters tell
    # UTF-16 apart; otherwise the declaration names the encoding. A BOM is
    # decoded as a character and written back as is.
    if content.startswith((codecs.BOM_UTF16_LE, b"<\x00")):
        return "utf-16-le"
    if content.startswith((codecs.BOM_UTF16_BE, b"\x00<")):
        return "utf-16-be"
    if content.startswith(codecs.BOM_UTF8):
        return "utf-8"
    declaration = ENCODING_PATTERN.match(content)
    return declaration.group(1).decode("ascii") if declaration else "utf-8"


def disable_vba_password(vba_project: bytes) -> bytes:
    # The DPB line of the PROJECT stream holds the VBA project password.
    # Renaming its key to DPX makes Excel drop the password. Only that line
    # is touched: the same bytes may appear in any other stream.
    try:
        spans = CompoundFile(vba_project).stream_spans("PROJECT")
    except Exception as e:
        warnings.warn(f"Unable to read the VBA project, patching it blindly: {e}")
        return vba_project.replace(b"DPB=", b"DPX=")

    positions = [offset + i for offset, length in spans for i in range(length)]
    project = bytes(vba_project[position] for position in positions)
    out = bytearray(vba_project)
    for match in re.finditer(rb"(?m)^DPB=", project):
        out[positions[match.start() + 2]] = ord("X")
    return bytes(out)


def unprotect_xlsx(input_path: str, output_path: str) -> list[VbaMatch]:
    # Some writers put the elements in a prefixed namespace, such as
    # <x:sheetProtection/>: the closing tag must then carry the same prefix.
    # Both <name/> and <name></name> are valid XML and used by writers.
    patterns_to_remove: list[str] = []
    for name in ["sheetProtection", "workbookProtection", "protectedRanges"]:
        patterns_to_remove += [
            rf"<(?:[\w.-]+:)?{name}(?:\s[^>]*)?/>",
            rf"<((?:[\w.-]+:)?){name}(?:\s[^>]*)?(?<!/)>.*?</\1{name}\s*>",
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
                    encoding = xml_encoding(content)
                    try:
                        text = content.decode(encoding)
                    except (LookupError, UnicodeDecodeError) as e:
                        warnings.warn(f"{item.filename} left unchanged: {e}")
                        zout.writestr(item.filename, content)
                        continue

                    for pattern in patterns_to_remove:
                        text = re.sub(pattern, "", text, flags=re.DOTALL)

                    text = TAG_PATTERN.sub(remove_attrs, text)

                    zout.writestr(item.filename, text.encode(encoding))

                elif item.filename == "xl/vbaProject.bin":
                    zout.writestr(item.filename, disable_vba_password(content))

                else:
                    zout.writestr(item.filename, content)

    # The conversion is done: a macro that cannot be parsed must not fail it.
    try:
        return find_protect_calls(output_path)
    except Exception as e:
        warnings.warn(f"Unable to scan VBA macros for Protect calls: {e}")
        return []
