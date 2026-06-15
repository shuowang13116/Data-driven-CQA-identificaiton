# CQA Identification Project Operating Context

This project uses the D drive Codex folder as the canonical working area:

`D:\Migrasome Therapeutics\2.CMC\1. 帝王蝶\2.专项报告\Data-driven CQA identification\Codex`

## Working Rule

All project-specific Codex artifacts should be created or updated under this folder, including:

- Markdown notes and migration/context files: `docs\`
- Python scripts and helper programs: `scripts\`
- Generated outputs, pilot tables, CSV/XLSX files, and reports: `outputs\`
- Project skills: `skills\`

The local workspace at `C:\Users\王硕\Documents\CQA identification` can be used as a temporary editing/sandbox location when needed, but final project artifacts should be copied back to the D drive Codex folder.

## Current Project Skill

`skills\prism-ppt-knockout-extractor`

Purpose: extract functional assay Y values, especially knockout-rate data, from decrypted experiment-report PowerPoint files containing embedded GraphPad Prism/OLE charts.

Important rule: select Prism objects by Y-axis title first. Use knockout/function charts as potency Y; exclude cell share, viability, gating, and population-state charts unless explicitly requested.

## Current Coulter Script

`scripts\extract_coulter_summary_pilot.py`

Purpose: build a single wide summary table from paired Coulter PDF reports and particle-data exports. This is intentionally not the four-table long-format structure.

Output style: one row per sample/file pair, with PDF-reported summary values and XLS-recomputed particle distribution statistics in the same row.
