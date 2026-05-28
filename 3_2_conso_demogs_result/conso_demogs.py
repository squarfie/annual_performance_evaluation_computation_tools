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
month_order = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

# STORAGE
consolidated = {}

for file in files:
    file_path = os.path.join(folder_path, file)

    try:
        site = file.replace("completeness_report_", "").replace(".xlsx", "").upper()

        df = pd.read_excel(file_path, sheet_name="Monthly Average", index_col=0)

        # df index = Month
        # df column = Average

        site_data = {}

        for month in month_order:
            if month in df.index:
                val = df.loc[month].values[0]
                site_data[month] = round(val, 2)
            else:
                site_data[month] = None

        consolidated[site] = site_data

    except Exception as e:
        print(f"Error processing {file}: {e}")

# === CREATE FINAL TABLE ===
final_df = pd.DataFrame(consolidated)

# reorder months
final_df = final_df.reindex(month_order)


# convert for calculations
numeric_df = final_df.replace('%', '', regex=True).astype(float)

# monthly avg (row)
final_df["% Average"] = numeric_df.mean(axis=1).round(2).astype(str) + "%"

# overall avg (column)
overall_row = numeric_df.mean(axis=0).round(2)
overall_row = overall_row.astype(str) + "%"

# overall avg for % column
overall_row["% Average"] = numeric_df.mean().mean().round(2).astype(str) + "%"

# add bottom row
final_df.loc["Monthly Average"] = overall_row

# save
output_path = os.path.join(folder_path, "completeness_consolidated.xlsx")

with pd.ExcelWriter(output_path) as writer:
    final_df.to_excel(writer, sheet_name="Completeness Summary")