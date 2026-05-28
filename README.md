# Annual Performance Evaluation Computation Tools

Python tools for generating and consolidating annual performance evaluation reports from WHONET and Excel data.


Download `annual_performance_evaluation_tools_v1.0.0_windows.zip`, extract it, then run the numbered executable tool you need.

## Tool Groups

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

Each folder contains the relevant Python script and, where available, a Word guide explaining the script context and computations.

## Notes

Generated build folders, executables, and report outputs are ignored by Git. Keep source scripts, reference files, and documentation in the repository; regenerate reports locally as needed.
