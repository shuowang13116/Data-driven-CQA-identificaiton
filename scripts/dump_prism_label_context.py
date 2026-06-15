#!/usr/bin/env python
from __future__ import annotations

import argparse
import struct
import zipfile
from pathlib import Path

from inspect_ole_cfb import CfbFile


def find_all(blob: bytes, needle: bytes) -> list[int]:
    out = []
    start = 0
    while True:
        idx = blob.find(needle, start)
        if idx < 0:
            return out
        out.append(idx)
        start = idx + 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ppt", required=True)
    parser.add_argument("--ole", required=True)
    parser.add_argument("--label", action="append", required=True)
    args = parser.parse_args()

    with zipfile.ZipFile(args.ppt) as zf:
        cfb = CfbFile(zf.read(args.ole))
    contents = b""
    for entry in cfb.entries:
        if entry.name == "CONTENTS":
            contents = cfb.read_stream(entry)
            break
    if not contents:
        raise SystemExit("No CONTENTS stream")

    for label in args.label:
        print(f"=== {args.ole} :: {label} ===")
        needles = [label.encode("latin1", errors="ignore")]
        for enc in ("utf-8", "gbk", "gb18030", "utf-16le"):
            try:
                val = label.encode(enc)
            except Exception:
                continue
            if val not in needles:
                needles.append(val)
        for needle in needles:
            if not needle:
                continue
            positions = find_all(contents, needle)
            if not positions:
                continue
            print("needle", needle[:32].hex(), "positions", positions[:8])
            for idx in positions[:3]:
                context = contents[max(0, idx - 32) : min(len(contents), idx + 96)]
                for enc in ("gb18030", "utf-8", "latin1"):
                    text = context.decode(enc, errors="ignore")
                    text = "".join(ch if ch.isprintable() else " " for ch in text)
                    if label[:2] in text:
                        print(enc, text)
                        break
                lo = max(0, idx - 64)
                hi = min(len(contents) - 4, idx + 256)
                vals = []
                for off in range(lo, min(len(contents) - 4, idx + 1200)):
                    val = struct.unpack("<f", contents[off : off + 4])[0]
                    if 0.5 <= val <= 100:
                        vals.append((off - idx, val))
                dedup = []
                for off, val in vals:
                    if not dedup or abs(dedup[-1][1] - val) > 1e-5 or off - dedup[-1][0] > 8:
                        dedup.append((off, val))
                print("offset floats", [(o, round(v, 6)) for o, v in dedup[:80]])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
