#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
import re
import struct
import zipfile
from dataclasses import dataclass
from pathlib import Path


FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC


@dataclass
class DirEntry:
    name: str
    object_type: int
    start_sector: int
    size: int


class CfbFile:
    def __init__(self, data: bytes):
        if data[:8] != bytes.fromhex("d0cf11e0a1b11ae1"):
            raise ValueError("not an OLE compound file")
        self.data = data
        self.sector_size = 1 << struct.unpack_from("<H", data, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", data, 32)[0]
        self.first_dir_sector = struct.unpack_from("<I", data, 48)[0]
        self.mini_cutoff = struct.unpack_from("<I", data, 56)[0]
        self.first_mini_fat_sector = struct.unpack_from("<I", data, 60)[0]
        self.num_mini_fat_sectors = struct.unpack_from("<I", data, 64)[0]
        self.difat = [x for x in struct.unpack_from("<109I", data, 76) if x not in (FREESECT, DIFSECT)]
        self.fat = self._load_fat()
        self.entries = self._load_dir_entries()
        self.root = self.entries[0]
        self.mini_fat = self._load_mini_fat()
        self.mini_stream = self._read_regular_stream(self.root.start_sector, self.root.size)

    def _sector_offset(self, sector: int) -> int:
        return 512 + sector * self.sector_size

    def _sector(self, sector: int) -> bytes:
        off = self._sector_offset(sector)
        return self.data[off : off + self.sector_size]

    def _chain(self, start: int, fat: list[int] | None = None) -> list[int]:
        if start in (FREESECT, ENDOFCHAIN):
            return []
        table = self.fat if fat is None else fat
        out = []
        seen = set()
        cur = start
        while cur not in (FREESECT, ENDOFCHAIN) and cur < len(table) and cur not in seen:
            out.append(cur)
            seen.add(cur)
            cur = table[cur]
        return out

    def _load_fat(self) -> list[int]:
        fat: list[int] = []
        for sec in self.difat:
            block = self._sector(sec)
            fat.extend(struct.unpack("<" + "I" * (self.sector_size // 4), block))
        return fat

    def _read_regular_stream(self, start: int, size: int) -> bytes:
        buf = b"".join(self._sector(sec) for sec in self._chain(start))
        return buf[:size]

    def _load_dir_entries(self) -> list[DirEntry]:
        raw = self._read_regular_stream(self.first_dir_sector, 1 << 30)
        entries: list[DirEntry] = []
        for off in range(0, len(raw) - 127, 128):
            item = raw[off : off + 128]
            name_len = struct.unpack_from("<H", item, 64)[0]
            if name_len < 2:
                continue
            name = item[: name_len - 2].decode("utf-16le", errors="ignore")
            object_type = item[66]
            start_sector = struct.unpack_from("<I", item, 116)[0]
            size = struct.unpack_from("<Q", item, 120)[0]
            entries.append(DirEntry(name, object_type, start_sector, size))
        return entries

    def _load_mini_fat(self) -> list[int]:
        if self.first_mini_fat_sector in (FREESECT, ENDOFCHAIN):
            return []
        raw = b"".join(self._sector(sec) for sec in self._chain(self.first_mini_fat_sector))
        if self.num_mini_fat_sectors:
            raw = raw[: self.num_mini_fat_sectors * self.sector_size]
        return list(struct.unpack("<" + "I" * (len(raw) // 4), raw))

    def read_stream(self, entry: DirEntry) -> bytes:
        if entry.object_type != 2:
            return b""
        if entry.size < self.mini_cutoff and self.mini_fat and entry.start_sector < len(self.mini_fat):
            chunks = []
            for mini in self._chain(entry.start_sector, self.mini_fat):
                off = mini * self.mini_sector_size
                chunks.append(self.mini_stream[off : off + self.mini_sector_size])
            return b"".join(chunks)[: entry.size]
        return self._read_regular_stream(entry.start_sector, entry.size)


def clean_text(blob: bytes) -> str:
    parts = []
    for enc in ("utf-8", "gb18030", "gbk", "utf-16le", "latin1"):
        txt = blob.decode(enc, errors="ignore")
        txt = "".join(ch if ch.isprintable() else " " for ch in txt)
        txt = re.sub(r"\s+", " ", txt)
        if txt.strip():
            parts.append(txt)
    return "\n".join(parts)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ppt", required=True)
    parser.add_argument("--ole", action="append", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    rows = []
    with zipfile.ZipFile(args.ppt) as zf:
        for ole in args.ole:
            cfb = CfbFile(zf.read(ole))
            for entry in cfb.entries:
                if entry.object_type != 2:
                    continue
                blob = cfb.read_stream(entry)
                text = clean_text(blob)
                hits = []
                for kw in ("敲除", "细胞占比", "knockout", "function", "Prism", "Y1Title", "Dose", "Sample"):
                    if kw.lower() in text.lower():
                        hits.append(kw)
                rows.append(
                    {
                        "ole": ole,
                        "stream": entry.name,
                        "size": entry.size,
                        "hits": ";".join(hits),
                        "preview": text[:1500],
                    }
                )

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["ole", "stream", "size", "hits", "preview"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} stream rows to {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
