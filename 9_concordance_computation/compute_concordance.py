import pandas as pd
import os
from tkinter import Tk, filedialog

# === SELECT FILE ===
Tk().withdraw()
file_path = filedialog.askopenfilename(title="Select Concordance Excel file")

if not file_path:
    print("No file selected.")
    exit()

# === LOAD DATA ===
df = pd.read_excel(file_path)

# normalize columns
df.columns = df.columns.str.strip().str.lower()

# === CHECK SITE ===
if "site" not in df.columns:
    raise ValueError("Column 'Site' not found.")

# === SAFE DIVISION FUNCTION ===
def safe_div(num, denom):
    return num / denom.replace(0, pd.NA)

# === COMPUTE ALL REQUIRED METRICS ===
if {"con_genus","num_vandp"}.issubset(df.columns):
    df["calc_genus"] = safe_div(df["con_genus"], df["num_vandp"]) * 100

if {"con_spp","num_vandp"}.issubset(df.columns):
    df["calc_spp"] = safe_div(df["con_spp"], df["num_vandp"]) * 100

if {"crit_dev","total_ast"}.issubset(df.columns):
    df["calc_crit"] = safe_div(df["crit_dev"], df["total_ast"]) * 100

if {"total_d","total_ast"}.issubset(df.columns):
    df["calc_total_dev"] = safe_div(df["total_d"], df["total_ast"]) * 100

if {"num_vandp","num_isol"}.issubset(df.columns):
    df["calc_viable"] = safe_div(df["num_vandp"], df["num_isol"]) * 100

if {"num_nonv","num_isol"}.issubset(df.columns):
    df["calc_nonv"] = safe_div(df["num_nonv"], df["num_isol"]) * 100

if {"num_mix","num_isol"}.issubset(df.columns):
    df["calc_mix"] = safe_div(df["num_mix"], df["num_isol"]) * 100

# === COMPUTATIONS ===
computations = {
    "Sum_Num_Isol": ("num_isol", "sum"),
    "Avg_Genus": ("calc_genus", "mean"),
    "Avg_Spp": ("calc_spp", "mean"),
    "Avg_Crit": ("calc_crit", "mean"),
    "Avg_Total_Dev": ("calc_total_dev", "mean"),
    "Avg_Viable": ("calc_viable", "mean"),
    "Avg_Non_V": ("calc_nonv", "mean"),
    "Avg_Mix": ("calc_mix", "mean"),
    "Sum_Num_VandP": ("num_vandp", "sum"),
    "Sum_Total_AST": ("total_ast", "sum"),
}

# === OUTPUT ===
output_path = os.path.join(os.path.dirname(file_path), "concordance_summary.xlsx")

sheets_written = 0

with pd.ExcelWriter(output_path) as writer:

    for sheet_name, (col, agg_func) in computations.items():

        if col not in df.columns:
            print(f"⚠ Missing column: {col}")
            continue

        grouped = df.groupby("site")[col].agg(agg_func).reset_index()

        if agg_func == "mean":
            grouped[col] = grouped[col].round(2)

        grouped = grouped.rename(columns={
            col: sheet_name.replace("_", " ")
        })

        grouped.to_excel(writer, sheet_name=sheet_name, index=False)

        sheets_written += 1

    if sheets_written == 0:
        df.head().to_excel(writer, sheet_name="DEBUG")

print("\n✅ Concordance computation complete (formula-based)!")
print(f"Saved at: {output_path}")