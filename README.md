# Annual Performance Evaluation Computation Tools

Python tools for generating and consolidating annual performance evaluation reports from WHONET and Excel data.

## Download Executable Tools

For users who do not want to run the Python scripts directly, download the Windows executable bundle from the repository's **Releases** page:

```text
annual_performance_evaluation_tools_v1.0.0_windows.zip
```

Extract the zip file, then run the numbered `.exe` tool you need.

## Executable Tool Guide

1. `01_convert_dbf_to_excel.exe` - Converts WHONET DBF-style files to Excel.
2. `02_add_organism_panel_group.exe` - Adds `Panel_Group_arsp` using organism mapping.
3. `03_compute_demographics_completeness.exe` - Computes monthly demographics completeness.
4. `04_consolidate_demographics_completeness.exe` - Consolidates demographics completeness reports.
5. `05_consolidate_diagnosis_completeness.exe` - Extracts and consolidates Diagnosis completeness.
6. `06_compute_qc_matrix.exe` - Computes QC matrix scores.
7. `07_consolidate_qc_matrix.exe` - Consolidates QC matrix outputs.
8. `08_compute_on_time_submission.exe` - Computes on-time submission scores.
9. `09_compute_panel_completeness.exe` - Computes antibiotic panel completeness.
10. `10_consolidate_panel_completeness.exe` - Consolidates panel completeness outputs.
11. `11_compute_no_growth_encoding.exe` - Computes no-growth/negative encoding scores.
12. `12_consolidate_no_growth_encoding.exe` - Consolidates no-growth/negative encoding outputs.
13. `13_compute_concordance.exe` - Computes concordance summary metrics.
14. `14_compute_referred_compliance.exe` - Computes referred compliance outputs.

## Source Folder Guide

- `1_convert_dbf_to_excel`: convert WHONET DBF-style files to Excel.
- `2_organism_grouping`: add `Panel_Group_arsp` organism groupings.
- `3_1_demogs_completeness`: compute demographics completeness.
- `3_2_conso_demogs_result`: consolidate demographics completeness results.
- `4_encoding_diagnosis`: consolidate diagnosis completeness.
- `5_1_encoding_qc`: compute QC matrix scores.
- `5_2_consolidate_qc_matrix`: consolidate QC matrix outputs.
- `6_on_time_submission`: compute on-time submission scores.
- `7_1_panel_computation`: compute antibiotic panel completeness.
- `7_2_consolidate_panel_comp`: consolidate panel completeness outputs.
- `8_1_compute_negatives`: compute no-growth/negative encoding scores.
- `8_2_consolidate_negatives`: consolidate no-growth/negative encoding outputs.
- `9_concordance_computation`: compute concordance summary metrics.
- `10_referred_computation`: compute referred compliance outputs.

Each folder contains the relevant Python script and, where available, a Word guide explaining the script context and computations.

## Source Repository Notes

The repository stores source scripts, reference files, specifications, and computation guides. Generated build folders, executables, report outputs, and release bundles are ignored by Git to keep the repository lightweight.

Some tools depend on reference workbooks included in their `references` folders. Keep those files with the scripts when running from source.
