# Annual Performance Evaluation Computation Tools

Python tools for generating and consolidating annual performance evaluation reports from WHONET and Excel data.

## Tool Groups

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

## Notes

Generated build folders, executables, and report outputs are ignored by Git. Keep source scripts, reference files, and documentation in the repository; regenerate reports locally as needed.
