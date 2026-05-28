import pandas as pd
import os
from tkinter import Tk, filedialog

# === SELECT FOLDER (where qc_matrix files are) ===
Tk().withdraw()
folder_path = filedialog.askdirectory(title="Select folder with qc_matrix files")

if not folder_path:
    print("No folder selected.")
    exit()

files = [f for f in os.listdir(folder_path) if f.startswith("qc_matrix_") and f.endswith(".xlsx")]

if not files:
    print("No qc_matrix files found.")
    exit()

ordered_months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

consolidated = []

for file in files:
    file_path = os.path.join(folder_path, file)

    try:
        site = file.replace("qc_matrix_", "").replace(".xlsx", "").upper()

        df = pd.read_excel(file_path, sheet_name="QC Target (PSE)", index_col=0)

        # Get Monthly Avg row
        if "Monthly Avg" not in df.index:
            print(f"⚠ Monthly Avg not found in {file}")
            continue

        row = df.loc["Monthly Avg"]

        record = {"Site": site}

        for m in ordered_months:
            val = row.get(m, 0)
            record[m] = f"{int(val)}%"  # format as %

        record["Average"] = f"{int(row.get('Average %', 0))}%"

        consolidated.append(record)

    except Exception as e:
        print(f"Error processing {file}: {e}")

# === CREATE FINAL TABLE ===
final_df = pd.DataFrame(consolidated)

# sort by site
final_df = final_df.sort_values("Site")

# === SAVE ===
output_path = os.path.join(folder_path, "qc_consolidated_report.xlsx")

with pd.ExcelWriter(output_path) as writer:
    final_df.to_excel(writer, index=False, sheet_name="QC Summary")

print("\n✅ Consolidated QC report generated!")
print(f"Saved at: {output_path}")