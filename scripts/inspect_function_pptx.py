#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
import re
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


TEXT_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
CONTROL_LABELS = {
    "wtih antibody no emig",
    "with antibody no emig",
    "1ug plasmid",
}


def natural_slide_key(name: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", name)
    return int(match.group(1)) if match else 0


def decode_text(blob: bytes) -> str:
    chunks: list[str] = []
    for enc in ("utf-8", "gb18030", "gbk", "utf-16le", "latin1"):
        try:
            text = blob.decode(enc, errors="ignore")
        except Exception:
            continue
        text = "".join(ch if ch.isprintable() or ch in "\r\n\t" else " " for ch in text)
        text = re.sub(r"\s+", " ", text)
        chunks.append(text)
    return "\n".join(chunks)


def slide_text(zf: zipfile.ZipFile, slide_xml: str) -> str:
    root = ET.fromstring(zf.read(slide_xml))
    texts = [node.text for node in root.findall(".//a:t", TEXT_NS) if node.text]
    return " | ".join(texts)


def slide_ole_targets(zf: zipfile.ZipFile, slide_no: int) -> list[str]:
    rel_name = f"ppt/slides/_rels/slide{slide_no}.xml.rels"
    if rel_name not in zf.namelist():
        return []
    xml = zf.read(rel_name).decode("utf-8", errors="ignore")
    return [
        "ppt/embeddings/" + target
        for target in re.findall(r'Target="\.\./embeddings/([^"]+\.bin)"', xml)
    ]


def keyword_hits(text: str) -> str:
    keywords = [
        "敲除",
        "knockout",
        "KO",
        "功能",
        "function",
        "梯度",
        "细胞占比",
        "viability",
        "gating",
        "population",
        "Prism",
    ]
    low = text.lower()
    hits = []
    for kw in keywords:
        if kw.lower() in low:
            hits.append(kw)
    return ";".join(hits)


def strings_near_keywords(text: str) -> str:
    terms = ["敲除", "knockout", "功能", "function", "梯度", "细胞占比", "viability", "Prism"]
    snippets = []
    low = text.lower()
    for term in terms:
        idx = low.find(term.lower())
        if idx >= 0:
            snippets.append(text[max(0, idx - 80) : idx + 160])
    return " || ".join(dict.fromkeys(snippets))[:1000]


def read_values_near_label(blob: bytes, label: str) -> list[float]:
    raw_label = label.encode("latin1", errors="ignore")
    if not raw_label:
        return []
    idx = blob.find(raw_label)
    if idx < 0:
        return []
    values = []
    for off in range(max(0, idx - 64), min(len(blob) - 4, idx + 512)):
        val = struct.unpack("<f", blob[off : off + 4])[0]
        if 0.0 <= val <= 100.0:
            if not values or abs(values[-1] - val) > 1e-5:
                values.append(val)
        if len(values) >= 12:
            break
    return values


def label_candidates(blob: bytes, labels: list[str]) -> str:
    parts = []
    for label in labels:
        if label in CONTROL_LABELS:
            continue
        vals = read_values_near_label(blob, label)
        if vals:
            parts.append(label + ":" + ",".join(f"{v:.4g}" for v in vals[:6]))
    return " | ".join(parts)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ppt", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--labels", default="F1,F2,F3,F4,F5,F6,F7,F8+9,F8-9,F8,F9,NC,PBS-速冻,TRE-慢冻")
    args = parser.parse_args()

    ppt = Path(args.ppt)
    header = ppt.read_bytes()[:8]
    if not header.startswith(b"PK") or not zipfile.is_zipfile(ppt):
        raise SystemExit("Input is not a readable/decrypted PPTX package; provide a decrypted copy.")

    labels = [x.strip() for x in args.labels.split(",") if x.strip()]
    rows: list[dict[str, str | int]] = []
    with zipfile.ZipFile(ppt) as zf:
        slides = sorted(
            [n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=natural_slide_key,
        )
        for slide in slides:
            slide_no = natural_slide_key(slide)
            text = slide_text(zf, slide)
            targets = slide_ole_targets(zf, slide_no)
            rows.append(
                {
                    "kind": "slide",
                    "slide": slide_no,
                    "ole": "",
                    "keywords": keyword_hits(text),
                    "summary": text[:1000],
                    "label_values": "",
                }
            )
            for target in targets:
                blob = zf.read(target)
                decoded = decode_text(blob)
                rows.append(
                    {
                        "kind": "ole",
                        "slide": slide_no,
                        "ole": target,
                        "keywords": keyword_hits(decoded),
                        "summary": strings_near_keywords(decoded),
                        "label_values": label_candidates(blob, labels),
                    }
                )

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["kind", "slide", "ole", "keywords", "summary", "label_values"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} rows to {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
