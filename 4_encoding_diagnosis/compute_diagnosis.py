import pandas as pd
import os
from tkinter import Tk, filedialog

# === SELECT FOLDER ===
Tk().withdraw()
folder_path = filedialog.askdirectory(title="Select folder with completeness reports")

if not folder_path:
    print("No folder selected.")
    exit()

# === GET FILES ===
files = [f for f in os.listdir(folder_path) if f.startswith("completeness_report_") and f.endswith(".xlsx")]

if not files:
    print("No completeness_report files found.")
    exit()

# === MONTH ORDER ===
months = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

# STORAGE
data = {}

for file in files:
    file_path = os.path.join(folder_path, file)

    try:
        site = file.replace("completeness_report_", "").replace(".xlsx", "").upper()

        df = pd.read_excel(file_path, sheet_name="Completeness", index_col=0)

        # normalize index (row names)
        df.index = df.index.str.strip().str.lower()

        if "diagnosis" not in df.index:
            print(f"⚠ DIAGNOSIS row not found in {file}")
            continue

        row = df.loc["diagnosis"]

        site_data = {}

        for m in months:
            if m in row.index:
                val = row[m]
                site_data[m] = val
            else:
                site_data[m] = None

        data[site] = site_data

    except Exception as e:
        print(f"Error processing {file}: {e}")

# === CREATE FINAL TABLE ===
final_df = pd.DataFrame(data)

# reorder months
final_df = final_df.reindex(months)

# format %
final_df = final_df.applymap(lambda x: f"{x:.2f}%" if pd.notna(x) else "")

# === SAVE ===
output_path = os.path.join(folder_path, "diagnosis_consolidated.xlsx")

with pd.ExcelWriter(output_path) as writer:
    final_df.to_excel(writer, sheet_name="Diagnosis Summary")

print("\n✅ Diagnosis consolidation complete!")
print(f"Saved at: {output_path}")