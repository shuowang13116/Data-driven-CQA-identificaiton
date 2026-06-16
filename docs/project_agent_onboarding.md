# Data-driven CQA Identification: Agent Onboarding Context

Last updated: 2026-06-16

This document is the shared onboarding note for Codex agents working on the Data-driven CQA identification project. Read this before editing scripts, extracting data, generating summaries, or writing new project notes.

## Project Purpose

The project supports data-driven CQA identification for the DWD / 帝王蝶 program. The main goal is to convert experimental reports and instrument exports into clean, reviewable CSV/XLSX tables for downstream CQA/QC data mining.

Current work focuses on two data streams:

- Functional assay data from experiment-report PowerPoint files containing embedded GraphPad Prism charts.
- Coulter particle-size features from paired PDF reports and exported particle-data files.

Raw experimental source files are intentionally not tracked in the code repository. Generated summaries, helper scripts, and operating notes should be stored under the project Codex folder.

## Canonical Project Locations

Use this D drive folder as the canonical project folder:

```text
D:\Migrasome Therapeutics\2.CMC\1. 帝王蝶\2.专项报告\Data-driven CQA identification\Codex
```

The local Codex workspace may be:

```text
C:\Users\王硕\Documents\CQA identification
```

Working rule:

- Use the C drive workspace as a temporary editing/sandbox location when Codex permissions require it.
- Final project artifacts should be copied back to the D drive `Codex` folder.
- Put Markdown notes and context files in `docs\`.
- Put Python scripts and helper programs in `scripts\`.
- Put generated CSV/XLSX/PPTX outputs in `outputs\`.
- Put reusable project skills in `skills\`.

Main raw-data folder:

```text
D:\Migrasome Therapeutics\2.CMC\1. 帝王蝶\2.专项报告\Data-driven CQA identification\Raw and processed data-CQA identification
```

Known raw-data subfolders as of 2026-06-16:

- `All 实验汇报`
- `帝王蝶1.实验备份-20260609`
- `瑞芯`

Top-level project files near the Codex folder include:

- `Data-driven CQA identification-王硕-20260609.pptx`
- `DWD-allData-王硕-20260609.xlsx`

## Current Repository Shape

Important files and folders:

- `README.md`: project overview and current workflows.
- `docs/project_operating_context.md`: previous operating context and artifact placement rules.
- `docs/project_agent_onboarding.md`: this onboarding document.
- `scripts/extract_function_pptx_summary.py`: main functional assay extraction script.
- `scripts/build_function_summary_template_pptx.py`: builds a title-only functional assay summary PPTX template from report PPT filenames.
- `scripts/extract_coulter_summary_pilot.py`: builds pilot Coulter particle-size summary tables.
- `scripts/inspect_function_pptx.py`: inspects slide text, OLE targets, and candidate Prism sources.
- `scripts/inspect_ole_cfb.py`: reads OLE compound-file streams embedded inside PPTX files.
- `scripts/dump_prism_label_context.py`: debugs Prism label-adjacent values in an OLE object.
- `scripts/scan_prism_rows.py`: scans Prism rows to discover labels and assay layouts.
- `skills/prism-ppt-knockout-extractor`: project skill for extracting functional assay / knockout-rate values from embedded Prism charts.
- `outputs/`: generated functional assay and Coulter output tables.

## Functional Assay Workflow

Purpose: extract functional assay outputs from decrypted experiment-report PPTX files, especially knockout rate / functional potency Y values, and keep cell-share values as separate quality-attribute columns.

Manual plus Codex workflow:

1. The user prepares backup copies of original experiment files.
2. Codex can gather report PPT/PPTX files into one working folder if asked.
3. Codex can build a functional-data summary PPTX template from report filenames.
4. The user manually copies the relevant functional assay Prism plot page from each original PPT into the summary PPTX.
5. The user manually standardizes visible Prism plot Y-axis titles when needed.
6. Codex extracts data from the summary PPTX using embedded Prism/OLE source data when available.
7. The user reviews extracted values and group names; control naming is standardized manually when needed, usually with `NC` as the main control label.

Core extraction rules:

- Process only decrypted, readable PPT/PPTX files. Standard PPTX files start with `PK`.
- Prefer embedded GraphPad Prism/OLE original data over screenshot digitization.
- Select the Prism object by Y-axis title, not by object order.
- Treat knockout/function charts as potency Y, usually output as `knockout_rate_pct`.
- Do not use cell-share, viability, gating, or population-state plots as potency Y.
- Keep cell-share charts, usually `cell_share_pct`, as separate quality-attribute columns when requested by the current script.
- Preserve experimental groups such as `F1-F7`, `F8+9`, `NC`, `PBS-速冻`, `PBS-慢冻`, `TRE-速冻`, and `TRE-慢冻`.
- Extract assay controls but mark them, instead of silently deleting them.
- Known assay controls include `no antibody no emig`, `with antibody no emig`, `wtih antibody no emig`, `1 ug plasmid`, and variants with `1ug`.
- In Excel outputs, assay controls should appear after main data rows and be highlighted yellow when the script supports it.
- Always keep source evidence columns such as source OLE object, endpoint-selection evidence, byte offset, and data-check warnings.

Functional assay layouts:

- Two-replicate comparison: preserve both replicate values and compute the arithmetic mean.
- Three-level gradient: preserve all three ordered Y values and compute one summary value for modeling.
- Single-value comparison: preserve the original value without inventing replicates.
- Ambiguous layouts: return raw extracted values and ask for confirmation before choosing a modeling summary.

Main functional extraction script:

```powershell
python scripts\extract_function_pptx_summary.py --ppt "path\to\summary.pptx" --out "outputs\DWD-function-data-extracted.xlsx"
```

The script writes:

- CSV output.
- XLSX output.
- `*-excluded_ole.csv` with skipped OLE objects and reasons.

Important output columns include:

- `slide`
- `experiment`
- `group`
- `is_assay_control`
- `source`
- `knockout_rate_*`
- `cell_share_*`
- `data_check_warning`

## Functional Summary PPT Template

Purpose: create a title-only PPTX template where each slide corresponds to one experiment report. The user can then paste the relevant functional assay Prism page into each slide.

Main script:

```powershell
python scripts\build_function_summary_template_pptx.py --input-dir "path\to\report_ppts" --out "outputs\DWD-function-summary-template.pptx" --recursive --sort-by experiment
```

The script also writes a CSV manifest next to the PPTX.

Filename matching should be permissive. Valid report files may be named with Chinese variants of "实验汇报", "实验汇总", or "汇报", and normally include `DWD`.

Do not include design files, raw-data PPTs, IF/WB-only PPTs, or non-report files unless the user explicitly asks.

## Coulter Particle-Size Workflow

Purpose: build one wide summary table from paired Coulter PDF reports and exported particle-data files.

Main script:

```powershell
python scripts\extract_coulter_summary_pilot.py --input-dir "path\to\coulter_folder" --output-dir "outputs" --recursive --prefix "coulter_summary_NAME_DATE"
```

Input expectation:

- Paired PDF report and particle data export with the same stem.
- Particle-data exports can be `.xls`, `.xlsx`, `.csv`, `.tsv`, or `.txt`.
- Tested `.xls` files may actually be tab-delimited text exports.

Output style:

- One row per sample/file pair.
- PDF-reported summary values and XLS/export-recomputed particle distribution statistics in the same row.
- CSV and XLSX outputs are both produced.

Current output fields include sample identity/path fields, PDF-extracted particle metrics, and recomputed distribution metrics such as count, mean, D50, D10, D90, D90/D10, span, SD, CV, skewness, kurtosis, min, and max.

## Existing Outputs Worth Knowing

Representative output files in `outputs\` include:

- `DWD-function-data-extracted_20260615-control-flagged.xlsx`
- `DWD-function-data-extracted_20260615-control-flagged.csv`
- `DWD-function-data-extracted_20260615-control-flagged-excluded_ole.csv`
- `DWD-function-wide-control-test.xlsx`
- `DWD-function-wide-control-test.csv`
- `DWD-function-and-cell-share-extracted.csv`
- `DWD-function-pptx-inspection.csv`
- `DWD-function-summary-template_regex-title-test.pptx`
- `DWD-function-summary-template_regex-title-test.csv`
- `coulter_summary_瑞芯_20260402.xlsx`
- `coulter_summary_瑞芯_20260402.csv`

When adding new outputs, use descriptive names with source, date, and purpose where possible.

## Agent Operating Requirements

General:

- Before changing behavior, read the relevant script and existing docs.
- Preserve existing user data and unrelated changes.
- Do not delete or overwrite raw experimental source files.
- Prefer structured parsing over ad hoc text manipulation.
- Keep outputs reviewable: include source path, evidence, warnings, and control flags.
- If a PPT/PPTX is encrypted or unreadable, stop and ask for a decrypted copy.
- If a result is inferred rather than extracted from source data, label the source clearly.

Data quality:

- Never silently drop groups unless the user explicitly asked for filtering.
- Never use cell-share or population-state data as potency Y by default.
- Do not choose a Prism object merely because it is first or because its values look plausible.
- Use visible chart titles/Y-axis titles and OLE contents as selection evidence.
- When values or layouts are ambiguous, preserve raw values and add a warning.
- Generated tables should be suitable for manual review before modeling.

Artifact placement:

- Draft/edit in the C drive workspace if needed.
- Copy finalized docs/scripts/outputs back to the D drive `Codex` folder.
- Keep raw-data references as paths; do not copy bulky raw files into the repo unless explicitly requested.

Communication:

- The user may refer to Fred as the project architecture role. Treat Fred's role as maintaining project-level consistency across multiple Codex conversations.
- When a new agent starts, it should state that it has read this context and summarize which task-specific files it will inspect next.
- Use Chinese when communicating with the user unless they ask otherwise.

