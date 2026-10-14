from __future__ import annotations

import re
import struct
import zipfile
from dataclasses import dataclass

END_OF_CHAIN = 0xFFFFFFFA

# Matches Protect calls (ActiveWorkbook.Protect, Sheets(1).Protect, bare
# Protect in a sheet module) but neither Unprotect nor ProtectContents.
PROTECT_PATTERN = re.compile(r"(?<!\w)Protect\b", re.IGNORECASE)

# A string literal (an escaped "" reads as two adjacent literals) or a
# comment: ' anywhere outside a string, Rem at the start of a statement.
STRING_OR_COMMENT = re.compile(
    r'"[^"]*"?|\'.*|(?:^|(?<=:))\s*Rem\b.*', re.IGNORECASE
)


@dataclass(frozen=True)
class VbaMatch:
    module: str
    line_number: int
    line: str


def decompress(data: bytes) -> bytes:
    # [MS-OVBA] 2.4.1 compressed container
    if not data or data[0] != 1:
        raise ValueError("Invalid VBA compressed container signature")
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        end = min(pos + (header & 0x0FFF) + 1, len(data))
        if not header & 0x8000:
            out += data[pos : pos + 4096]
            pos += 4096
            continue
        chunk_start = len(out)
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not flags & (1 << bit):
                    out.append(data[pos])
                    pos += 1
                else:
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    bit_count = max((len(out) - chunk_start - 1).bit_length(), 4)
                    offset = (token >> (16 - bit_count)) + 1
                    length = (token & (0xFFFF >> bit_count)) + 3
                    for _ in range(length):
                        out.append(out[-offset])
    return bytes(out)


def _follow(start: int, next_of) -> list[int]:
    # A corrupted or crafted file can link a chain back on itself
    indexes: list[int] = []
    seen: set[int] = set()
    while start < END_OF_CHAIN:
        if start in seen:
            raise ValueError("Cyclic sector chain in compound file")
        seen.add(start)
        indexes.append(start)
        start = next_of(start)
    return indexes


class CompoundFile:
    # Minimal [MS-CFB] reader: only what is needed to read streams by name
    def __init__(self, data: bytes):
        self.data = data
        major_version = struct.unpack_from("<H", data, 0x1A)[0]
        self.sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", data, 0x20)[0]
        dir_start = struct.unpack_from("<I", data, 0x30)[0]
        self.mini_cutoff = struct.unpack_from("<I", data, 0x38)[0]
        minifat_start = struct.unpack_from("<I", data, 0x3C)[0]
        difat_next = struct.unpack_from("<I", data, 0x44)[0]
        ints_per_sector = self.sector_size // 4

        difat = list(struct.unpack_from("<109I", data, 0x4C))

        def next_difat(index: int) -> int:
            values = struct.unpack(f"<{ints_per_sector}I", self._sector(index))
            difat.extend(values[:-1])
            return values[-1]

        _follow(difat_next, next_difat)

        self.fat: list[int] = []
        for sector in difat:
            if sector < END_OF_CHAIN:
                self.fat += struct.unpack(
                    f"<{ints_per_sector}I", self._sector(sector)
                )

        self.minifat: list[int] = []
        if minifat_start < END_OF_CHAIN:
            raw = self._chain(minifat_start)
            self.minifat = list(struct.unpack(f"<{len(raw) // 4}I", raw))

        entries = []
        raw = self._chain(dir_start)
        for i in range(0, len(raw) - 127, 128):
            entry = raw[i : i + 128]
            name_len = struct.unpack_from("<H", entry, 64)[0]
            name = entry[: max(name_len - 2, 0)].decode("utf-16-le")
            entry_type = entry[66]
            siblings_and_child = struct.unpack_from("<III", entry, 68)
            start, size = struct.unpack_from("<IQ", entry, 116)
            if major_version == 3:
                # [MS-CFB] 2.6.3: old writers left garbage in the high 32 bits
                size &= 0xFFFFFFFF
            entries.append((name, entry_type, siblings_and_child, start, size))
        if not entries or entries[0][1] != 5:
            raise ValueError("Missing root entry in compound file")

        # Streams are keyed by their path, such as "vba/dir": the same name
        # can be used in different storages. The entries of a storage form a
        # tree linked by their left and right sibling indexes.
        self.streams: dict[str, tuple[int, int]] = {}
        pending = [(entries[0][2][2], "")]
        seen: set[int] = set()
        while pending:
            index, prefix = pending.pop()
            if index >= len(entries):  # also NOSTREAM, the empty link
                continue
            if index in seen:
                raise ValueError("Cyclic directory tree in compound file")
            seen.add(index)
            name, entry_type, (left, right, child), start, size = entries[index]
            pending += [(left, prefix), (right, prefix)]
            path = prefix + name.lower()
            if entry_type == 1:
                pending.append((child, path + "/"))
            elif entry_type == 2:
                self.streams[path] = (start, size)

        # The small streams live in the mini stream, itself stored in the
        # sectors chained from the root entry
        self.mini_stream_sectors = _follow(entries[0][3], self.fat.__getitem__)

    def _sector(self, index: int) -> bytes:
        offset = (index + 1) * self.sector_size
        return self.data[offset : offset + self.sector_size]

    def _chain(self, start: int) -> bytes:
        return b"".join(
            self._sector(i) for i in _follow(start, self.fat.__getitem__)
        )

    def stream_spans(self, path: str) -> list[tuple[int, int]]:
        # The (file offset, length) pieces that make up a stream, in order.
        # path names the storages from the root, separated by "/".
        start, size = self.streams[path.lower()]
        if size >= self.mini_cutoff:
            pieces = [
                ((i + 1) * self.sector_size, self.sector_size)
                for i in _follow(start, self.fat.__getitem__)
            ]
        else:
            per_sector = self.sector_size // self.mini_sector_size
            pieces = []
            for i in _follow(start, self.minifat.__getitem__):
                sector = self.mini_stream_sectors[i // per_sector]
                offset = (sector + 1) * self.sector_size
                offset += (i % per_sector) * self.mini_sector_size
                pieces.append((offset, self.mini_sector_size))

        spans: list[tuple[int, int]] = []
        for offset, length in pieces:
            length = min(length, size)
            if length <= 0:
                break
            spans.append((offset, length))
            size -= length
        # A truncated file is missing the end of some streams: fail rather than
        # return them cut short
        for offset, length in spans:
            if offset + length > len(self.data):
                raise ValueError(
                    f"Stream {path} goes past the end of the compound file"
                )
        return spans

    def stream(self, path: str) -> bytes:
        return b"".join(
            self.data[offset : offset + length]
            for offset, length in self.stream_spans(path)
        )


def extract_vba_modules(vba_project: bytes) -> list[tuple[str, str]]:
    cfb = CompoundFile(vba_project)
    dir_stream = decompress(cfb.stream("VBA/dir"))
    pos = 0
    codepage = 1252
    stream_name = ""
    modules: list[tuple[str, str]] = []
    while pos + 6 <= len(dir_stream):
        record_id, size = struct.unpack_from("<HI", dir_stream, pos)
        pos += 6
        # PROJECTVERSION declares a size of 4 but holds 6 bytes
        if record_id == 0x0009:
            size = 6
        record = dir_stream[pos : pos + size]
        pos += size
        if record_id == 0x0003:
            codepage = struct.unpack_from("<H", record)[0]
        elif record_id == 0x001A:
            stream_name = record.decode(f"cp{codepage}", errors="replace")
        elif record_id == 0x0031:
            text_offset = struct.unpack_from("<I", record)[0]
            source = decompress(cfb.stream(f"VBA/{stream_name}")[text_offset:])
            modules.append(
                (stream_name, source.decode(f"cp{codepage}", errors="replace"))
            )
    return modules


def _split_comment(line: str) -> tuple[str, str | None]:
    # Returns the code with string literals emptied, and the comment if any
    comment = None

    def replace(match: re.Match) -> str:
        nonlocal comment
        if match.group(0).startswith('"'):
            return '""'
        comment = match.group(0)
        return ""

    return STRING_OR_COMMENT.sub(replace, line), comment


def grep_vba_modules(
    modules: list[tuple[str, str]], pattern: re.Pattern = PROTECT_PATTERN
) -> list[VbaMatch]:
    matches: list[VbaMatch] = []
    for module_name, source in modules:
        in_comment = False
        for line_number, line in enumerate(source.splitlines(), 1):
            code = line.strip()
            if in_comment:
                searched, comment = "", code
            else:
                searched, comment = _split_comment(code)
            # A comment ending with " _" goes on over the next line
            in_comment = comment is not None and (
                comment == "_" or comment.endswith(" _")
            )
            if pattern.search(searched):
                matches.append(VbaMatch(module_name, line_number, code))
    return matches


def find_protect_calls(xlsx_path: str) -> list[VbaMatch]:
    with zipfile.ZipFile(xlsx_path, "r") as z:
        if "xl/vbaProject.bin" not in z.namelist():
            return []
        vba_project = z.read("xl/vbaProject.bin")
    return grep_vba_modules(extract_vba_modules(vba_project))
