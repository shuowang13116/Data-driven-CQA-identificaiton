#!/usr/bin/env python
from __future__ import annotations

import argparse
import re
import struct
import zipfile

from inspect_ole_cfb import CfbFile


def first_contents(ppt: str, ole: str) -> bytes:
    with zipfile.ZipFile(ppt) as zf:
        cfb = CfbFile(zf.read(ole))
    for entry in cfb.entries:
        if entry.name == "CONTENTS":
            return cfb.read_stream(entry)
    return b""


def f32(blob: bytes, off: int) -> float:
    return struct.unpack("<f", blob[off : off + 4])[0]


def label_before(blob: bytes, off: int) -> str:
    raw = blob[max(0, off - 90) : off]
    text = raw.decode("gb18030", errors="ignore")
    text = "".join(ch if ch.isprintable() else " " for ch in text)
    tokens = re.findall(r"[A-Za-z0-9_()+%.-]*[\u4e00-\u9fffA-Za-z][A-Za-z0-9_()+%.\-\u4e00-\u9fff]*", text)
    junk = {"P", "B", "r", "m", "k", "d", "x", "q", "Y1Title", "Dose", "Sample"}
    tokens = [tok for tok in tokens if tok not in junk and len(tok) >= 2]
    return tokens[-1] if tokens else ""


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ppt", required=True)
    parser.add_argument("--ole", required=True)
    args = parser.parse_args()

    blob = first_contents(args.ppt, args.ole)
    seen = set()
    rows = []
    for off in range(0, len(blob) - 100):
        v1 = f32(blob, off)
        if not 0.5 <= v1 <= 100:
            continue
        v2_12 = f32(blob, off + 12)
        v2_48 = f32(blob, off + 48)
        v3_96 = f32(blob, off + 96)
        label = label_before(blob, off)
        if not label:
            continue
        if 0.5 <= v2_12 <= 100:
            key = (label, "rep", round(v1, 4), round(v2_12, 4))
            if key not in seen:
                rows.append((off, label, "two-replicate", [v1, v2_12]))
                seen.add(key)
        if 0.5 <= v2_48 <= 100 and 0.5 <= v3_96 <= 100:
            key = (label, "grad", round(v1, 4), round(v2_48, 4), round(v3_96, 4))
            if key not in seen:
                rows.append((off, label, "3-level gradient", [v1, v2_48, v3_96]))
                seen.add(key)
    for off, label, layout, vals in rows[:80]:
        print(off, label, layout, ",".join(f"{v:.6g}" for v in vals))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
