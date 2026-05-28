import os
import pandas as pd
from dbfread import DBF
from tkinter import Tk, filedialog

# === Select folder ===
Tk().withdraw()
folder_path = filedialog.askdirectory(title="Select folder with WHONET files")

if not folder_path:
    print("No folder selected.")
    exit()

# === Ask user for extension ===
ext = input("Enter file extension to process (e.g. .BRH, .BRT, .CMC): ").strip().lower()

if not ext.startswith("."):
    ext = "." + ext

print(f"Processing files with extension: {ext}")

# === Output folder ===
output_folder = os.path.join(folder_path, "converted_excel")
os.makedirs(output_folder, exist_ok=True)

# === Detect files ===
files = [f for f in os.listdir(folder_path) if f.lower().endswith(ext)]

if not files:
    print("No matching files found.")
    exit()

print("Files found:", files)

# === Process files ===
for file in files:
    file_path = os.path.join(folder_path, file)

    try:
        print(f"\nProcessing {file}...")

        table = DBF(
            file_path,
            load=True,
            encoding="latin1",
            ignore_missing_memofile=True
        )

        df = pd.DataFrame(iter(table))

        if df.empty:
            print("⚠️ Empty or unsupported file, skipped")
            continue

        # safer replacement (handles .brh, .BRT, etc.)
        output_name = os.path.splitext(file)[0] + ".xlsx"
        output_file = os.path.join(output_folder, output_name)

        df.to_excel(output_file, index=False)

        print(f"✔ Saved: {output_file}")

    except Exception as e:
        print(f"❌ Error processing {file}: {e}")

print("\n✅ Conversion finished!")