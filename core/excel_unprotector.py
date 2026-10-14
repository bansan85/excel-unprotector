from __future__ import annotations

import codecs
import os
import zipfile
import re
import warnings
from xml.etree import ElementTree
from core.vba_grep import CompoundFile, VbaMatch, find_protect_calls

# A start or empty-element tag. Text content cannot contain "<" unescaped, so
# this never matches inside a cell value or a shared string.
TAG_PATTERN = re.compile(r"<[A-Za-z_][^\"'>]*(?:(?:\"[^\"]*\"|'[^']*')[^\"'>]*)*>")

# CDATA sections, comments and processing instructions may hold any text, even
# something that looks like a tag: they are kept as they are. Each pattern
# tries this alternative first, so a match cannot start inside one of them.
UNPARSED = r"(?P<unparsed><!\[CDATA\[.*?\]\]>|<!--.*?-->|<\?.*?\?>)"

# One attribute, with the whitespace in front of it. XML allows both quote
# characters and whitespace around "=".
ATTRIBUTE_PATTERN = re.compile(r"""\s+([\w.:-]+)\s*=\s*("[^"]*"|'[^']*')""")

ENCODING_PATTERN = re.compile(rb"<\?xml[^>]*encoding\s*=\s*[\"']([A-Za-z0-9._-]+)")

# Content types of the SpreadsheetML parts (workbook, sheets, styles...): they
# end with "+xml". The second prefix covers macro-enabled workbooks, macro
# sheets and the parts added by later Excel versions (slicers, timelines).
SPREADSHEET_TYPES = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.",
    "application/vnd.ms-excel.",
)


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


def spreadsheet_parts(zin: zipfile.ZipFile) -> set[str]:
    # Custom XML and document properties are not spreadsheet markup, so they
    # are left alone. [OPC] A part takes the content type of its Override, else
    # the one of the Default for its extension.
    try:
        types = ElementTree.fromstring(zin.read("[Content_Types].xml"))
    except (KeyError, ElementTree.ParseError) as e:
        warnings.warn(
            f"Unable to read [Content_Types].xml, every XML part is processed: {e}"
        )
        return {name for name in zin.namelist() if name.lower().endswith(".xml")}

    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for element in types:
        kind = element.tag.rsplit("}", 1)[-1]
        if kind == "Default":
            key = element.get("Extension", "").lower()
            defaults[key] = element.get("ContentType", "")
        elif kind == "Override":
            key = element.get("PartName", "").lower()
            overrides[key] = element.get("ContentType", "")

    parts: set[str] = set()
    for name in zin.namelist():
        extension = os.path.splitext(name)[1][1:].lower()
        declared = overrides.get(f"/{name}".lower(), defaults.get(extension, ""))
        # Media types are case-insensitive [RFC 6838]
        content_type = declared.lower()
        if content_type.startswith(SPREADSHEET_TYPES) and content_type.endswith("+xml"):
            parts.add(name)
    return parts


def unprotect_xlsx(input_path: str, output_path: str) -> list[VbaMatch]:
    # Some writers put the elements in a prefixed namespace, such as
    # <x:sheetProtection/>: the closing tag must then carry the same prefix.
    # Both <name/> and <name></name> are valid XML and used by writers.
    patterns_to_remove: list[str] = []
    for name in [
        "sheetProtection",
        "workbookProtection",
        "protectedRanges",
        # Holds the password to modify and the read-only recommendation
        "fileSharing",
    ]:
        patterns_to_remove += [
            rf"<(?:[\w.-]+:)?{name}(?:\s[^>]*)?/>",
            rf"<(?P<prefix>(?:[\w.-]+:)?){name}(?:\s[^>]*)?(?<!/)>.*?"
            rf"</(?P=prefix){name}\s*>",
        ]

    removal_patterns = [
        re.compile(f"{UNPARSED}|{pattern}", re.DOTALL) for pattern in patterns_to_remove
    ]
    tag_pattern = re.compile(f"{UNPARSED}|{TAG_PATTERN.pattern}", re.DOTALL)

    def remove_element(match: re.Match) -> str:
        if match.group("unparsed") is not None:
            return match.group(0)
        return ""

    def remove_attr(attr: re.Match) -> str:
        name, value = attr.group(1), attr.group(2)[1:-1]
        if name in ("windowProtection", "applyProtection"):
            return ""
        # Shows hidden and very hidden sheets. Excel cannot unhide a very hidden
        # sheet from its interface, only a macro can.
        if name == "state" and value in ("hidden", "veryHidden"):
            return ""
        return attr.group(0)

    def remove_attrs(tag: re.Match) -> str:
        if tag.group("unparsed") is not None:
            return tag.group(0)
        # Each match covers a whole attribute, value included, so text inside
        # a value is never taken for an attribute.
        return ATTRIBUTE_PATTERN.sub(remove_attr, tag.group(0))

    # Opening the output for writing would truncate the input while it is
    # read. samefile compares the files themselves, so a path differing only
    # by case on a case-insensitive file system, a symlink or a hard link to
    # the input is caught too.
    if os.path.exists(output_path) and os.path.samefile(input_path, output_path):
        raise ValueError("The output file must not be the input file")

    with zipfile.ZipFile(input_path, "r") as zin:
        parts = spreadsheet_parts(zin)
        with zipfile.ZipFile(
            output_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)

                if item.filename in parts:
                    encoding = xml_encoding(content)
                    try:
                        text = content.decode(encoding)
                    except (LookupError, UnicodeDecodeError) as e:
                        warnings.warn(f"{item.filename} left unchanged: {e}")
                        zout.writestr(item.filename, content)
                        continue

                    for removal in removal_patterns:
                        text = removal.sub(remove_element, text)

                    text = tag_pattern.sub(remove_attrs, text)

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
