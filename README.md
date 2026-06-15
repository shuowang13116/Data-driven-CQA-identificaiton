# Data-driven CQA Identification

Utility scripts and operating notes for extracting CQA/QC features and functional assay outputs used in data-driven CQA identification.

## What This Code Does

- Extract functional assay outputs from experiment-report PPTX files, especially flow-cytometry knockout rate (`knockout_rate_pct`) and quality-attribute cell share (`cell_share_pct`).
- Read embedded GraphPad Prism/OLE source data when available, instead of estimating values from screenshots.
- Summarize assay layouts correctly:
  - two-replicate comparisons -> replicate values plus `mean_y_pct`
  - 3-level gradients -> ordered Y values plus an AUC-like summary
  - single-value comparisons -> preserve the original value without inventing a summary
- Extract Coulter particle-size features from paired PDF reports and exported particle-data files.
- Produce clean CSV/XLSX tables for downstream CQA/QC data mining.

## Main Scripts

- `scripts/extract_function_pptx_summary.py`: extracts knockout/function Y values and cell-share quality attributes from decrypted PowerPoint files containing embedded GraphPad Prism/OLE charts.
- `scripts/build_function_summary_template_pptx.py`: builds a title-only functional assay summary PPT template from multiple report PPT filenames, using `DWD-实验名` slide titles and no extra instruction text boxes.
- `scripts/extract_coulter_summary_pilot.py`: builds pilot Coulter particle-size summary tables from paired report/data exports.
- `scripts/inspect_function_pptx.py`: inspects slide text, OLE objects, and candidate Prism sources before extraction.
- `scripts/inspect_ole_cfb.py`: reads OLE compound-file streams embedded inside PPTX files.
- `scripts/dump_prism_label_context.py`: debugs label-adjacent Prism values in an OLE object.
- `scripts/scan_prism_rows.py`: scans Prism rows to discover labels and assay layouts when labels are not known in advance.
- `docs/project_operating_context.md`: local project operating context and artifact placement rules.

## Functional Assay Data Workflow

The manual/Codex workflow for building the functional assay summary is:

1. Manually prepare backup copies of the source experiment data.
2. Use Codex to copy all `2. 实验汇报` / `2. 实验汇总` PPTX files into one working folder.
3. Use Codex to build a functional-data summary PPT template from those PPT names. Each slide title should contain only the experiment name extracted from the file name, for example `DWD-DS260101-3rd密度梯度离心`.
4. Manually open each original PPT and copy the page containing the functional assay result Prism plot into the summary PPT.
5. Manually standardize Prism plot Y-axis titles, usually to `细胞占比%` and `敲除率（%）`, and remove irrelevant information.
6. Use Codex to extract functional assay data from the summary PPT:
   - include `敲除率` / knockout / function plots as potency Y
   - include `细胞占比%` plots as quality-attribute columns, not as potency Y
   - extract every Prism x-axis group that can be read, then mark assay controls instead of deleting them
   - output an Excel/CSV table with experiment names, group labels, assay layout, raw Y values, summaries, source OLE, selection evidence, and data-check warnings
   - use the Excel output for manual review; `no antibody no emig`, `wtih/with antibody no emig`, and `1 μg plasmid` assay controls are moved to the bottom and highlighted yellow
7. Manually review extracted values and group names. Standardize each experiment's control group name to `NC`.

## Key Extraction Rules

- Process only decrypted, readable PPT/PPTX files.
- Prefer Prism/OLE original data over image digitization.
- Select the Prism object by Y-axis title, not by object order.
- Do not use cell-share or cell-state plots as functional Y; keep `细胞占比%` as separate quality-attribute columns.
- Keep experimental groups such as `F1-F7`, `F8+9`, `NC`, `PBS-速冻`, and `TRE-慢冻`.
- Extract assay controls such as `no antibody no emig`, `with antibody no emig`, `wtih antibody no emig`, and `1 μg plasmid`, but mark them as `is_assay_control=TRUE`, move them after the main data rows, and highlight them yellow in Excel.
- Record the data source as `Prism/OLE original data`, `embedded image digitization estimate`, or `manual/visual estimate`.

Raw experimental files are intentionally not tracked in this repository.
