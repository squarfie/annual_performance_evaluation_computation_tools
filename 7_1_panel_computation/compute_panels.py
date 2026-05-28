import pandas as pd
import os
from tkinter import Tk, filedialog

# =========================================================
# MAIN LOOP
# =========================================================
while True:

    # =====================================================
    # SELECT FOLDER
    # =====================================================
    Tk().withdraw()

    folder_path = filedialog.askdirectory(
        title="Select folder containing WITH_PANEL Excel files"
    )

    if not folder_path:
        print("No folder selected.")
        break

    # =====================================================
    # SELECT PANEL REFERENCE FILE
    # =====================================================
    panel_file = filedialog.askopenfilename(
        title="Select Panel Reference Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not panel_file:
        print("No panel reference selected.")
        break

    # =====================================================
    # LOAD PANEL REFERENCE
    # =====================================================
    panel_df = pd.read_excel(panel_file)

    panel_df.columns = panel_df.columns.str.strip()

    required_cols = [
        "Arsp_org_grouping",
        "Antibiotic_Panels",
        "WHON5_CODE_DISK",
        "WHON5_CODE_MIC"
    ]

    for col in required_cols:
        if col not in panel_df.columns:
            raise ValueError(f"Missing required column: {col}")

    # =====================================================
    # OUTPUT FOLDER
    # =====================================================
    output_folder = os.path.join(
        folder_path,
        "panel_completeness_reports"
    )

    os.makedirs(output_folder, exist_ok=True)

    # =====================================================
    # GET FILES
    # =====================================================
    files = [
        f for f in os.listdir(folder_path)
        if f.endswith(".xlsx")
    ]

    if not files:
        print("No Excel files found.")
        continue

    # =====================================================
    # PROCESS FILES
    # =====================================================
    for file in files:

        file_path = os.path.join(folder_path, file)

        print(f"\nProcessing: {file}")

        try:

            df = pd.read_excel(file_path)

            df.columns = df.columns.str.strip()

            if "Panel_Group_arsp" not in df.columns:
                print("⚠ Panel_Group_arsp column not found.")
                continue

            # =============================================
            # GET SITE CODE
            # =============================================
            site_code = "UNKNOWN"

            if "LABORATORY" in df.columns:

                lab_vals = (
                    df["LABORATORY"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                )

                lab_vals = lab_vals[lab_vals != ""]

                if not lab_vals.empty:
                    site_code = lab_vals.iloc[0]

            if (
                site_code == "UNKNOWN"
                and
                "INSTITUT" in df.columns
            ):

                inst_vals = (
                    df["INSTITUT"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                )

                inst_vals = inst_vals[inst_vals != ""]

                if not inst_vals.empty:
                    site_code = inst_vals.iloc[0]

            site_code = str(site_code).upper()

            # =============================================
            # OUTPUT FILE
            # =============================================
            filename_only = os.path.splitext(file)[0]

            output_file = os.path.join(
                output_folder,
                f"panel_completeness_{filename_only}_{site_code}.xlsx"
            )

            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

                used_sheet_names = set()
                sheets_written = 0
                summary_rows = []

                unique_groups = (
                    df["Panel_Group_arsp"]
                    .dropna()
                    .astype(str)
                    .unique()
                )

                for org_group in unique_groups:

                    org_df = df[
                        df["Panel_Group_arsp"]
                        .astype(str)
                        == str(org_group)
                    ].copy()

                    if org_df.empty:
                        continue

                    # =============================================
                    # CLEAN TEXT FUNCTION
                    # =============================================
                    def clean_text(x):
                        return (
                            str(x)
                            .strip()
                            .lower()
                            .replace("\n", " ")
                            .replace("\r", " ")
                            .replace("\t", " ")
                        )

                    # =============================================
                    # PANEL REFERENCE FILTER
                    # =============================================
                    ref_df = panel_df[
                        panel_df["Arsp_org_grouping"]
                        .apply(clean_text)
                        ==
                        clean_text(org_group)
                    ].copy()

                    if ref_df.empty:
                        print(f"⚠ No panel reference for: {org_group}")
                        continue

                    results = []

                    for _, row in ref_df.iterrows():

                        abx_name = str(row["Antibiotic_Panels"]).strip()

                        disk_code = str(row["WHON5_CODE_DISK"]).strip()
                        mic_code = str(row["WHON5_CODE_MIC"]).strip()

                        possible_cols = []

                        if "|" in disk_code:
                            possible_cols.extend(
                                [x.strip() for x in disk_code.split("|")]
                            )
                        else:
                            possible_cols.append(disk_code)

                        if "|" in mic_code:
                            possible_cols.extend(
                                [x.strip() for x in mic_code.split("|")]
                            )
                        else:
                            possible_cols.append(mic_code)

                        possible_cols = [
                            x for x in possible_cols
                            if x and x.lower() != "nan"
                        ]

                        existing_cols = [
                            c for c in possible_cols
                            if c in org_df.columns
                        ]

                        total_isolates = len(org_df)

                        tested_count = 0

                        if total_isolates == 0:
                            percent = 0

                        elif not existing_cols:
                            percent = 0

                        else:

                            tested_mask = pd.Series(
                                False,
                                index=org_df.index
                            )

                            for col in existing_cols:

                                tested_mask |= (
                                    org_df[col]
                                    .notna()
                                    &
                                    (
                                        org_df[col]
                                        .astype(str)
                                        .str.strip()
                                        != ""
                                    )
                                )

                            tested_count = tested_mask.sum()

                            percent = round(
                                (tested_count / total_isolates) * 100,
                                2
                            )

                        results.append({
                            "organism": org_group,
                            "antibiotic": abx_name,
                            "number_total": total_isolates,
                            "number_tested": tested_count,
                            "percent_all": percent
                        })

                    result_df = pd.DataFrame(results)

                    if result_df.empty:
                        continue

                    summary_avg = round(
                        result_df["percent_all"].mean(),
                        2
                    )

                    summary_rows.append({
                        "Organism Group": org_group,
                        "Number of isolates": len(org_df),
                        "Average": summary_avg
                    })

                    # =========================================
                    # SAFE SHEET NAME
                    # =========================================
                    safe_sheet = str(org_group)

                    invalid_chars = [
                        "\\", "/", "*", "?", ":", "[", "]"
                    ]

                    for ch in invalid_chars:
                        safe_sheet = safe_sheet.replace(ch, "_")

                    safe_sheet = safe_sheet[:31]

                    original_name = safe_sheet
                    counter = 1

                    while safe_sheet in used_sheet_names:
                        suffix = f"_{counter}"
                        safe_sheet = (
                            original_name[:31-len(suffix)]
                            + suffix
                        )
                        counter += 1

                    used_sheet_names.add(safe_sheet)

                    result_df.to_excel(
                        writer,
                        sheet_name=safe_sheet,
                        index=False
                    )

                    sheets_written += 1

                # =============================================
                # SUMMARY SHEET
                # =============================================
                if summary_rows:

                    summary_df = pd.DataFrame(summary_rows)

                    summary_df.to_excel(
                        writer,
                        sheet_name="Summary",
                        index=False
                    )

                    sheets_written += 1

                # =============================================
                # FAILSAFE
                # =============================================
                if sheets_written == 0:

                    dummy_df = pd.DataFrame({
                        "Message": ["No valid sheets generated"]
                    })

                    dummy_df.to_excel(
                        writer,
                        sheet_name="EMPTY",
                        index=False
                    )

            print(f"✔ Saved: {output_file}")

        except Exception as e:
            print(f"❌ Error processing {file}: {e}")

    print("\n✅ All panel completeness reports generated!")

    # =====================================================
    # ASK USER
    # =====================================================
    again = input(
        "\nDo you want to compute another folder? (Y/N): "
    ).strip().lower()

    if again != "y":
        print("Done.")
        break