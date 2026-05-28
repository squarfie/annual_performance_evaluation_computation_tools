import pandas as pd
import os
from tkinter import Tk, filedialog

# =========================================================
# MONTH MAP
# =========================================================
month_lookup = {
    "01": "January",
    "02": "February",
    "03": "March",
    "04": "April",
    "05": "May",
    "06": "June",
    "07": "July",
    "08": "August",
    "09": "September",
    "10": "October",
    "11": "November",
    "12": "December"
}

ordered_months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
]

# =========================================================
# EXTRACT MONTH FROM FILE
# =========================================================
def extract_month(filename):

    filename = filename.upper()

    # FIND W0125PHL PATTERN
    import re

    match = re.search(r'W(\d{2})25PHL', filename)

    if not match:
        return None

    month_num = match.group(1)

    return month_lookup.get(month_num)

# =========================================================
# MAIN LOOP
# =========================================================
while True:

    # =====================================================
    # SELECT FOLDER
    # =====================================================
    Tk().withdraw()

    folder_path = filedialog.askdirectory(
        title="Select folder with panel completeness reports"
    )

    if not folder_path:
        print("No folder selected.")
        break

    # =====================================================
    # GET FILES
    # =====================================================
    files = [
        f for f in os.listdir(folder_path)
        if f.lower().endswith(".xlsx")
        and "panel_completeness" in f.lower()
    ]

    if not files:
        print("No panel completeness files found.")
        continue

    # =====================================================
    # STORAGE
    # =====================================================
    consolidated = {}

    # =====================================================
    # PROCESS FILES
    # =====================================================
    for file in files:

        file_path = os.path.join(folder_path, file)

        print(f"Processing: {file}")

        try:

            # =============================================
            # EXTRACT MONTH
            # =============================================
            month = extract_month(file)

            if month is None:
                print(f"⚠ Cannot determine month from {file}")
                continue

            # =============================================
            # EXTRACT SITE CODE
            # =============================================
            # panel_completeness_with_panel_W0125PHL_APM.xlsx

            filename_no_ext = os.path.splitext(file)[0]

            site_code = filename_no_ext.split("_")[-1].upper()

            # =============================================
            # READ SUMMARY SHEET
            # =============================================
            summary_df = pd.read_excel(
                file_path,
                sheet_name="Summary"
            )

            summary_df.columns = summary_df.columns.str.strip()

            required_cols = [
                "Average"
            ]

            for col in required_cols:

                if col not in summary_df.columns:
                    raise ValueError(
                        f"Missing column '{col}' in Summary sheet"
                    )

            # =============================================
            # GET OVERALL AVERAGE
            # =============================================
            overall_avg = round(
                summary_df["Average"].mean(),
                2
            )

            # =============================================
            # STORE
            # =============================================
            if month not in consolidated:
                consolidated[month] = {}

            consolidated[month][site_code] = overall_avg

        except Exception as e:
            print(f"❌ Error processing {file}: {e}")

    # =====================================================
    # CREATE FINAL TABLE
    # =====================================================
    final_df = pd.DataFrame(consolidated).T

    # order months
    final_df = final_df.reindex(ordered_months)

    # sort sites
    final_df = final_df.reindex(
        sorted(final_df.columns),
        axis=1
    )

    # =====================================================
    # MONTHLY AVERAGE COLUMN
    # =====================================================
    final_df["% Average"] = round(
        final_df.mean(axis=1),
        2
    )

    # =====================================================
    # OVERALL SITE AVERAGE ROW
    # =====================================================
    overall_row = round(
        final_df.mean(axis=0),
        2
    )

    final_df.loc["Monthly Average"] = overall_row

    # =====================================================
    # FORMAT %
    # =====================================================
    final_df = final_df.astype(object)

    for row in final_df.index:
        for col in final_df.columns:

            value = final_df.loc[row, col]

            if pd.notna(value):

                try:
                    final_df.loc[row, col] = f"{float(value):.2f}%"
                except:
                    pass

    # =====================================================
    # SAVE OUTPUT
    # =====================================================
    output_file = os.path.join(
        folder_path,
        "Consolidated_Panel_Completeness.xlsx"
    )

    with pd.ExcelWriter(output_file) as writer:

        final_df.to_excel(
            writer,
            sheet_name="Panel Completeness"
        )

    print(f"\n✔ Saved: {output_file}")

    # =====================================================
    # ASK USER
    # =====================================================
    again = input(
        "\nDo you want to process another folder? (Y/N): "
    ).strip().lower()

    if again != "y":
        print("Done.")
        break