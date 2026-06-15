#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
import html
import re
import zipfile
from pathlib import Path


REPORT_RE = re.compile(r"^\s*2[.\uFF0E、\-\s]*(实验汇报|实验汇总|汇报|汇总).*\.(ppt|pptx|pptm)$", re.I)


def extract_experiment_name(text: str) -> str | None:
    stem = Path(text).stem
    normalized = re.sub(r"\s*\|\s*", "-", stem)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    match = re.search(r"(DWD-.+?)-王硕(?:-\d{8})?", normalized)
    if match:
        return match.group(1).strip("- ")
    match = re.search(r"(DWD-.+?)(?:-敲除率|-细胞占比|-保密信息|$)", normalized)
    if match:
        return match.group(1).strip("- ")
    match = re.search(r"(DWD-.+)", normalized)
    if match:
        return match.group(1).strip("- ")
    return None


def collect_report_ppts(input_dir: Path, recursive: bool, sort_by: str = "name") -> list[Path]:
    files = input_dir.rglob("*") if recursive else input_dir.iterdir()
    reports = [
        p
        for p in files
        if p.is_file() and REPORT_RE.match(p.name) and extract_experiment_name(p.name)
    ]
    if sort_by == "experiment":
        return sorted(reports, key=lambda p: extract_experiment_name(p.name) or "")
    return sorted(reports, key=lambda p: p.name)


def xml_escape(text: str) -> str:
    return html.escape(text, quote=True)


def slide_xml(title: str, slide_no: int) -> str:
    title = xml_escape(title)
    subtitle = xml_escape(f"Slide {slide_no}")
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Experiment Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="640000" y="420000"/><a:ext cx="11400000" cy="850000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square"/><a:lstStyle/>
          <a:p><a:pPr algn="l"/><a:r><a:rPr lang="zh-CN" sz="3200" b="1"><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:rPr><a:t>{title}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Instruction"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="640000" y="1400000"/><a:ext cx="11400000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square"/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="zh-CN" sz="1500"><a:solidFill><a:srgbClr val="666666"/></a:solidFill></a:rPr><a:t>{subtitle}: paste the functional Prism plot page here.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''


def write_pptx(titles: list[str], out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    slide_overrides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, len(titles) + 1)
    )
    slide_rels = "\n".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, len(titles) + 1)
    )
    slide_ids = "\n".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i}"/>' for i in range(1, len(titles) + 1)
    )

    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  {slide_overrides}
</Types>''',
        )
        zf.writestr(
            "_rels/.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>''',
        )
        zf.writestr(
            "ppt/presentation.xml",
            f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>''',
        )
        zf.writestr(
            "ppt/_rels/presentation.xml.rels",
            f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {slide_rels}
</Relationships>''',
        )
        zf.writestr("ppt/presProps.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')
        zf.writestr("ppt/viewProps.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')
        zf.writestr("docProps/app.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application><Slides>{len(titles)}</Slides></Properties>')
        zf.writestr("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Functional assay summary template</dc:title></cp:coreProperties>')
        for i, title in enumerate(titles, start=1):
            zf.writestr(f"ppt/slides/slide{i}.xml", slide_xml(title, i))
            zf.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')


def write_manifest(rows: list[dict[str, str]], out: Path) -> None:
    manifest = out.with_suffix(".csv")
    with manifest.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["slide", "experiment", "source_file"])
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a title-only functional assay summary PPTX template from report PPT filenames.")
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--limit", type=int, default=None, help="Only include the first N matching report PPT files.")
    parser.add_argument("--sort-by", choices=["name", "experiment"], default="name", help="Sort matching report PPT files before applying --limit.")
    args = parser.parse_args()

    reports = collect_report_ppts(Path(args.input_dir), recursive=args.recursive, sort_by=args.sort_by)
    if args.limit is not None:
        reports = reports[: args.limit]
    if not reports:
        raise SystemExit("No report PPT/PPTX files found.")
    rows = [
        {"slide": str(i), "experiment": extract_experiment_name(path.name) or "", "source_file": str(path)}
        for i, path in enumerate(reports, start=1)
    ]
    out = Path(args.out)
    write_pptx([row["experiment"] for row in rows], out)
    write_manifest(rows, out)
    print(f"wrote {len(rows)} slides to {out}")
    print(f"wrote manifest to {out.with_suffix('.csv')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
