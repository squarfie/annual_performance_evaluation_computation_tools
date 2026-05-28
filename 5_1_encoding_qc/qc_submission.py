import pandas as pd
import os
from tkinter import Tk, filedialog

# === MONTH MAP ===
month_lookup = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
    "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
    "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
}

ordered_months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def extract_month(filename):
    return month_lookup.get(filename[1:3], filename)

# TARGET ORGANISMS
target_orgs = ["aba", "pae", "sau", "eco", "kpn", "spn", "hin"]

# === MAIN PROCESS FUNCTION ===
def process_folder(folder_path):

    files = sorted([f for f in os.listdir(folder_path) if f.endswith(".xlsx")])

    if not files:
        print("No Excel files found.")
        return

    site_data = {}

    for file in files:
        file_path = os.path.join(folder_path, file)

        try:
            df = pd.read_excel(file_path)
            df.replace("", pd.NA, inplace=True)

            # === HELPER ===
            def get_first_valid(series):
                series = series.dropna().astype(str).str.strip()
                series = series[series != ""]
                return series.iloc[0] if not series.empty else None

            # === SITE CODE (LAB priority) ===
            site_code = "unknown"

            if "LABORATORY" in df.columns:
                val = get_first_valid(df["LABORATORY"])
                if val:
                    site_code = val

            elif "INSTITUT" in df.columns:
                val = get_first_valid(df["INSTITUT"])
                if val:
                    site_code = val

            site_code = site_code.upper()

            print(f"{file} → SITE = {site_code}")

            month = extract_month(file)

            # INIT
            if site_code not in site_data:
                site_data[site_code] = {org: {} for org in target_orgs}

            # === FILTER QC ===
            qc_df = pd.DataFrame()

            if "SPEC_TYPE" in df.columns:
                qc_df = df[df["SPEC_TYPE"].fillna("").str.lower() == "qc"]

            if not qc_df.empty and "ORGANISM" in qc_df.columns:
                org_series = qc_df["ORGANISM"].fillna("").str.lower()

                for org in target_orgs:
                    count = (org_series == org).sum()

                    # === SCORING RULE ===
                    if count >= 2:
                        score = 100
                    elif count == 1:
                        score = 50
                    else:
                        score = 0

                    site_data[site_code][org][month] = score

            else:
                # no QC data
                for org in target_orgs:
                    site_data[site_code][org][month] = 0

        except Exception as e:
            print(f"Error processing {file}: {e}")

    # === GENERATE OUTPUT ===
    for site, org_data in site_data.items():

        df_out = pd.DataFrame(org_data).T

        # Arrange months
        df_out = df_out.reindex(columns=ordered_months).fillna(0)

        # === ROW AVERAGE ===
        df_out["Average %"] = df_out.mean(axis=1).round(0)

        # === COLUMN AVERAGE ===
        avg_row = df_out[ordered_months].mean(axis=0).round(0)

        overall_avg = avg_row.mean().round(0)

        avg_row["Average %"] = overall_avg

        # ✅ LABEL ROW
        df_out.loc["Monthly Avg"] = avg_row

        # === CREATE TARGET TABLE (pae, sau, eco ONLY) ===
        target_rows = ["pae", "sau", "eco"]

        df_target = df_out.loc[target_rows].copy()

        # recompute row averages (only for these 3)
        df_target["Average %"] = df_target[ordered_months].mean(axis=1).round(0)

        # monthly avg (only these 3)
        target_avg_row = df_target[ordered_months].mean(axis=0).round(0)
        target_overall = target_avg_row.mean().round(0)
        target_avg_row["Average %"] = target_overall

        df_target.loc["Monthly Avg"] = target_avg_row

        # === SAVE FILE ===
        output_file = os.path.join(folder_path, f"qc_matrix_{site.lower()}.xlsx")

        with pd.ExcelWriter(output_file) as writer:
            df_out.to_excel(writer, sheet_name="QC Matrix")
            df_target.to_excel(writer, sheet_name="QC Target (PSE)")

        print(f"✔ Saved: {output_file}")


# === LOOP EXECUTION ===
while True:
    Tk().withdraw()
    folder_path = filedialog.askdirectory(title="Select folder with monthly Excel files")

    if not folder_path:
        print("No folder selected. Exiting...")
        break

    process_folder(folder_path)

    choice = input("\nConvert another folder? (Y/N): ").strip().lower()

    if choice != "y":
        print("Done.")
        break