#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
import re
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from inspect_ole_cfb import CfbFile, clean_text


TEXT_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
CONTROL_LABELS = {
    "wtih antibody no emig",
    "with antibody no emig",
    "1ug plasmid",
}
LABELS = [
    "混合后eMig",
    "预组装beads",
    "eMig直接加Apo",
    "20um一级",
    "SD-3+0.8",
    "SD-1.2",
    "90%-400RPM",
    "70%-200RPM",
    "60%-200RPM",
    "50%-100RPM",
    "NC-10ug",
    "NC-6ug",
    "NC-2ug",
    "MK-2.5",
    "MK-1.2",
    "eC-eM",
    "SD-8",
    "SD-5",
    "SD-3",
    "MK-5",
    "TRE-effect-eMig",
    "PBS-速冻",
    "PBS-慢冻",
    "TRE-速冻",
    "TRE-慢冻",
    "PBS-fake",
    "未处理组",
    "(NC)90%-400RPM",
    "(NC)90%-600RPM",
    "(NC)80%-400RPM",
    "(NC)80%-600RPM",
    "F8+9",
    "F8-9",
    "F10",
    "F9",
    "F8",
    "F7",
    "F6",
    "F5",
    "F4",
    "F3",
    "F2",
    "F1",
    "NC",
]


def slide_no(name: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", name)
    return int(match.group(1)) if match else 0


def slide_text(zf: zipfile.ZipFile, slide_xml: str) -> str:
    root = ET.fromstring(zf.read(slide_xml))
    return " | ".join(node.text for node in root.findall(".//a:t", TEXT_NS) if node.text)


def extract_experiment_name(raw_slide_title: str) -> str:
    normalized = re.sub(r"\s*\|\s*", "-", raw_slide_title)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    match = re.search(r"(DWD-.+?)-王硕-\d{8}", normalized)
    if match:
        return match.group(1).strip("- ")
    match = re.search(r"(DWD-.+?)(?:-敲除率|-细胞占比|-保密信息|$)", normalized)
    if match:
        return match.group(1).strip("- ")
    return raw_slide_title


def slide_ole_targets(zf: zipfile.ZipFile, no: int) -> list[str]:
    rel_name = f"ppt/slides/_rels/slide{no}.xml.rels"
    if rel_name not in zf.namelist():
        return []
    xml = zf.read(rel_name).decode("utf-8", errors="ignore")
    return ["ppt/embeddings/" + target for target in re.findall(r'Target="\.\./embeddings/([^"]+\.bin)"', xml)]


def stream_map(blob: bytes) -> dict[str, bytes]:
    cfb = CfbFile(blob)
    out = {}
    for entry in cfb.entries:
        if entry.object_type == 2:
            out[entry.name] = cfb.read_stream(entry)
    return out


def is_function_ole(streams: dict[str, bytes], visible_slide_text: str = "") -> tuple[bool, str]:
    contents = clean_text(streams.get("CONTENTS", b""))
    preview_text = "\n".join(clean_text(data) for name, data in streams.items() if "OlePres" in name)
    has_function_axis = "敲除" in contents or "knockout" in contents.lower() or "function" in contents.lower()
    has_function_preview = "敲除" in preview_text or "knockout" in preview_text.lower() or "function" in preview_text.lower()
    has_cell_share = "细胞占比" in contents or "viability" in contents.lower() or "population" in contents.lower()
    if has_function_axis or has_function_preview:
        evidence = "y-axis/visible chart contains 敲除率"
        return True, evidence
    if has_cell_share:
        return False, "excluded: y-axis contains 细胞占比/cell-state"
    if "敲除" in visible_slide_text or "knockout" in visible_slide_text.lower() or "功能" in visible_slide_text:
        if "Prism" in clean_text(streams.get("\x01CompObj", b"")):
            return True, "visible slide title contains 敲除率; OLE is Prism and does not contain 细胞占比"
    return False, "excluded: no function/knockout y-axis evidence"


def label_needles(label: str) -> list[bytes]:
    needles = []
    encodings = ("latin1", "utf-8") if label.isascii() else ("utf-8", "gbk", "gb18030")
    for enc in encodings:
        try:
            raw = label.encode(enc, errors="strict")
        except Exception:
            continue
        if raw and raw not in needles:
            needles.append(raw)
    return needles


def find_positions(blob: bytes, needle: bytes) -> list[int]:
    positions = []
    start = 0
    while True:
        idx = blob.find(needle, start)
        if idx < 0:
            return positions
        positions.append(idx)
        start = idx + 1


def looks_embedded_in_longer_label(blob: bytes, idx: int, raw: bytes) -> bool:
    if not re.fullmatch(rb"[A-Za-z]+\d*", raw):
        return False
    nxt = idx + len(raw)
    if nxt >= len(blob):
        return False
    return blob[nxt : nxt + 1] in b"+-0123456789"


def candidate_floats(blob: bytes, idx: int) -> list[tuple[int, float]]:
    vals = []
    for off in range(idx, min(len(blob) - 4, idx + 180)):
        val = struct.unpack("<f", blob[off : off + 4])[0]
        if 0.5 <= val <= 100:
            rel = off - idx
            if not vals or abs(vals[-1][1] - val) > 1e-5 or rel - vals[-1][0] > 8:
                vals.append((rel, val))
    return vals


def extract_label_values(contents: bytes, label: str) -> tuple[str, list[float], int] | None:
    best: tuple[str, list[float], int] | None = None
    for raw in label_needles(label):
        for idx in find_positions(contents, raw):
            if looks_embedded_in_longer_label(contents, idx, raw):
                continue
            vals = candidate_floats(contents, idx)
            if not vals:
                continue
            rels = [rel for rel, _ in vals]
            nums = [val for _, val in vals]
            if len(vals) >= 3:
                for pos in range(0, len(vals) - 2):
                    if abs((rels[pos + 1] - rels[pos]) - 48) <= 4 and abs((rels[pos + 2] - rels[pos + 1]) - 48) <= 4:
                        return "3-level gradient", nums[pos : pos + 3], idx
            if len(vals) >= 2:
                for pos in range(0, len(vals) - 1):
                    if abs((rels[pos + 1] - rels[pos]) - 12) <= 3:
                        return "two-replicate comparison", nums[pos : pos + 2], idx
            if vals and 12 <= rels[0] <= 45:
                return "single-value comparison", nums[:1], idx
            if best is None:
                best = ("ambiguous", nums[: min(3, len(nums))], idx)
    return best


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ppt", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    ppt = Path(args.ppt)
    header = ppt.read_bytes()[:8]
    if not header.startswith(b"PK") or not zipfile.is_zipfile(ppt):
        raise SystemExit("Input is not a readable/decrypted PPTX package; provide a decrypted copy.")

    rows = []
    excluded = []
    with zipfile.ZipFile(ppt) as zf:
        slides = sorted(
            [name for name in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", name)],
            key=slide_no,
        )
        for slide in slides:
            no = slide_no(slide)
            title = slide_text(zf, slide)
            experiment_name = extract_experiment_name(title)
            for ole in slide_ole_targets(zf, no):
                streams = stream_map(zf.read(ole))
                is_target, evidence = is_function_ole(streams, title)
                if not is_target:
                    excluded.append({"slide": no, "experiment": experiment_name, "source_ole": ole, "reason": evidence})
                    continue
                contents = streams.get("CONTENTS", b"")
                for label in LABELS:
                    if label in CONTROL_LABELS:
                        continue
                    extracted = extract_label_values(contents, label)
                    if not extracted:
                        continue
                    layout, values, offset = extracted
                    row = {
                        "slide": no,
                        "experiment": experiment_name,
                        "group": label,
                        "assay_layout": layout,
                        "source": "Prism/OLE original data",
                        "source_ole": ole,
                        "selection_evidence": evidence,
                        "byte_offset_in_contents": offset,
                        "gradient_1_y_pct": "",
                        "gradient_2_y_pct": "",
                        "gradient_3_y_pct": "",
                        "replicate_1_y_pct": "",
                        "replicate_2_y_pct": "",
                        "mean_y_pct": "",
                        "single_y_pct": "",
                        "knockout_rate_pct_summary": "",
                    }
                    if layout == "3-level gradient" and len(values) >= 3:
                        row["gradient_1_y_pct"] = round(values[0], 6)
                        row["gradient_2_y_pct"] = round(values[1], 6)
                        row["gradient_3_y_pct"] = round(values[2], 6)
                        row["knockout_rate_pct_summary"] = round(sum(values[:3]) / 3.0, 6)
                    elif len(values) >= 2:
                        row["replicate_1_y_pct"] = round(values[0], 6)
                        row["replicate_2_y_pct"] = round(values[1], 6)
                        row["mean_y_pct"] = round(sum(values[:2]) / 2.0, 6)
                    elif len(values) == 1:
                        row["single_y_pct"] = round(values[0], 6)
                    rows.append(row)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "slide",
        "experiment",
        "group",
        "assay_layout",
        "gradient_1_y_pct",
        "gradient_2_y_pct",
        "gradient_3_y_pct",
        "knockout_rate_pct_summary",
        "replicate_1_y_pct",
        "replicate_2_y_pct",
        "mean_y_pct",
        "single_y_pct",
        "source",
        "source_ole",
        "selection_evidence",
        "byte_offset_in_contents",
    ]
    with out.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    excluded_out = out.with_name(out.stem + "-excluded_ole.csv")
    with excluded_out.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["slide", "experiment", "source_ole", "reason"])
        writer.writeheader()
        writer.writerows(excluded)
    print(f"Wrote {len(rows)} extracted rows to {out}")
    print(f"Wrote {len(excluded)} excluded OLE rows to {excluded_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
