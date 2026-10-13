from __future__ import annotations

import re
import struct
import zipfile
from dataclasses import dataclass

END_OF_CHAIN = 0xFFFFFFFA

# Matches Protect calls (ActiveWorkbook.Protect, Sheets(1).Protect, bare
# Protect in a sheet module) but neither Unprotect nor ProtectContents.
PROTECT_PATTERN = re.compile(r"(?<!\w)Protect\b", re.IGNORECASE)


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

        self.streams: dict[str, tuple[int, int]] = {}
        root = (END_OF_CHAIN, 0)
        raw = self._chain(dir_start)
        for i in range(0, len(raw), 128):
            entry = raw[i : i + 128]
            name_len = struct.unpack_from("<H", entry, 64)[0]
            name = entry[: max(name_len - 2, 0)].decode("utf-16-le")
            entry_type = entry[66]
            start, size = struct.unpack_from("<IQ", entry, 116)
            if entry_type == 5:
                root = (start, size)
            elif entry_type == 2:
                self.streams.setdefault(name.lower(), (start, size))
        self.mini_stream = self._chain(root[0])[: root[1]]

    def _sector(self, index: int) -> bytes:
        offset = (index + 1) * self.sector_size
        return self.data[offset : offset + self.sector_size]

    def _chain(self, start: int) -> bytes:
        return b"".join(
            self._sector(i) for i in _follow(start, self.fat.__getitem__)
        )

    def stream(self, name: str) -> bytes:
        start, size = self.streams[name.lower()]
        if size >= self.mini_cutoff:
            return self._chain(start)[:size]
        out = bytearray()
        for i in _follow(start, self.minifat.__getitem__):
            offset = i * self.mini_sector_size
            out += self.mini_stream[offset : offset + self.mini_sector_size]
        return bytes(out[:size])


def extract_vba_modules(vba_project: bytes) -> list[tuple[str, str]]:
    cfb = CompoundFile(vba_project)
    dir_stream = decompress(cfb.stream("dir"))
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
            source = decompress(cfb.stream(stream_name)[text_offset:])
            modules.append(
                (stream_name, source.decode(f"cp{codepage}", errors="replace"))
            )
    return modules


def grep_vba_modules(
    modules: list[tuple[str, str]], pattern: re.Pattern = PROTECT_PATTERN
) -> list[VbaMatch]:
    matches: list[VbaMatch] = []
    for module_name, source in modules:
        for line_number, line in enumerate(source.splitlines(), 1):
            code = line.strip()
            if code.startswith("'") or code.upper().startswith("REM "):
                continue
            if pattern.search(code):
                matches.append(VbaMatch(module_name, line_number, code))
    return matches


def find_protect_calls(xlsx_path: str) -> list[VbaMatch]:
    with zipfile.ZipFile(xlsx_path, "r") as z:
        if "xl/vbaProject.bin" not in z.namelist():
            return []
        vba_project = z.read("xl/vbaProject.bin")
    return grep_vba_modules(extract_vba_modules(vba_project))
