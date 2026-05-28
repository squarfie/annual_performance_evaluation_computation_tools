import pandas as pd
import os
from tkinter import Tk, filedialog

# === STEP 1: Ask user for folder ===
Tk().withdraw()  # hide main window
folder_path = filedialog.askdirectory(title="Select folder containing monthly Excel files")
print("Folder selected:", folder_path)

if not folder_path:
    print("No folder selected. Exiting...")
    exit()

print(f"Selected folder: {folder_path}")

# === STEP 2: Detect Excel files ===
files = sorted([f for f in os.listdir(folder_path) if f.endswith(".xlsx")])

if not files:
    print("No Excel files found in the selected folder.")
    exit()

# Optional: Map filenames to month names automatically
month_lookup = {
    "01": "January", "02": "February", "03": "March", "04": "April",
    "05": "May", "06": "June", "07": "July", "08": "August",
    "09": "September", "10": "October", "11": "November", "12": "December"
}


def spec_type_cond(df):
    return ~df["SPEC_TYPE"].fillna("").str.lower().isin(
        ["un", "wa", "fo", "mi", "en", "qc", "fb"]
    )


def extract_month(filename):
    # Example: W0125PHL.xlsx → "01"
    try:
        month_num = filename[1:3]
        return month_lookup.get(month_num, filename)
    except:
        return filename
    
# def ward_condition(df):
#     return df["WARD_TYPE"].str.lower().isin(["in", "icu", "inx", "mix", "med", "brn", "pic", "ped", "obg", "sur", "neo"])


def ward_condition(df):
    exclude = ["out", "opd", "eme", "lab", "unk", "er", "oth", "env", "emr", "op", ""]

    # Choose column: WARD_TYPE → fallback to WARD
    if "WARD_TYPE" in df.columns:
        col = "WARD_TYPE"
    elif "WARD" in df.columns:
        col = "WARD"
    else:
        # If neither exists, return all False (no rows included)
        return pd.Series([False] * len(df), index=df.index)

    ward_series = df[col].fillna("").astype(str).str.strip().str.lower()

    return ~ward_series.isin(exclude)



# === STEP 3: Define indicators ===
indicators = {
    "Identification number": {"cols": ["PATIENT_ID"]},
    "Name": {"cols": ["LAST_NAME"]},
    "Sex": {"cols": ["SEX"]},
    "Age": {"cols": ["AGE"]},
    "Date of birth": {"cols": ["DATE_BIRTH"]},
    "Location/Ward": {"cols": ["WARD"]},
    "Department": {"cols": ["DEPARTMENT"]},
    "Ward type": {"cols": ["WARD_TYPE"]},
    "Specimen number": {"cols": ["SPEC_NUM"]},
    "Specimen date": {"cols": ["SPEC_DATE"]},
    "Specimen type": {"cols": ["SPEC_TYPE"]},
    "Organism": {"cols": ["ORGANISM"]},

    # SPECIAL CASE HERE
    "Date of admission (Total inpatient)": {
        "cols": ["DATE_ADMIS"],
        "condition": ward_condition
    },

    "Diagnosis": {"cols": ["DIAGNOSIS"]}
}

    


# === STEP 4: Completeness function ===
def compute_completeness(df, columns, condition=None):
    # Apply condition if provided
    if condition is not None:
        df = df[condition(df)]

    total_rows = len(df)

    if total_rows == 0:
        return 0

    filled = df[columns].notna().all(axis=1).sum()

    return (filled / total_rows) * 100

# === STEP 5: Process files ===
results = {}

for file in files:
    file_path = os.path.join(folder_path, file)

    try:
        df = pd.read_excel(file_path)
        df.replace("", pd.NA, inplace=True)

        # Apply global exclusion
        if "SPEC_TYPE" in df.columns:
            df = df[spec_type_cond(df)]

        month_name = extract_month(file)
        month_result = {}

        for indicator, config in indicators.items():
            cols = config["cols"]
            condition = config.get("condition")

            existing_cols = [c for c in cols if c in df.columns]

            if not existing_cols:
                month_result[indicator] = None
                continue

            completeness = compute_completeness(df, existing_cols, condition)
            month_result[indicator] = round(completeness, 2)

        results[month_name] = month_result

    except Exception as e:
        print(f"Error processing {file}: {e}")

# === STEP 6: Convert to DataFrame ===
report_df = pd.DataFrame(results)

# Sort columns by calendar order
ordered_months = list(month_lookup.values())
report_df = report_df.reindex(columns=[m for m in ordered_months if m in report_df.columns])

# === STEP 7: Monthly averages ===
monthly_avg = report_df.mean(axis=0).round(2)


# === Extract INSTITUT or fallback to LABORATORY ===
institut_code = "unknown"

try:
    # Try from first file (recommended approach)
    first_file = os.path.join(folder_path, files[0])
    df_temp = pd.read_excel(first_file)

    def get_first_valid(series):
        series = series.dropna().astype(str).str.strip()
        series = series[series != ""]
        return series.iloc[0] if not series.empty else None

    institut_val = None
    lab_val = None

    if "INSTITUT" in df_temp.columns:
        institut_val = get_first_valid(df_temp["INSTITUT"])

    if not institut_val and "LABORATORY" in df_temp.columns:
        lab_val = get_first_valid(df_temp["LABORATORY"])

    final_code = institut_val if institut_val else lab_val

    if final_code:
        institut_code = final_code.lower()

except Exception as e:
    print(f"Could not extract INSTITUT/LABORATORY: {e}")



# === STEP 8: Save output in same folder ===
output_filename = f"completeness_report_{institut_code}.xlsx"
output_path = os.path.join(folder_path, output_filename)

with pd.ExcelWriter(output_path) as writer:
    report_df.to_excel(writer, sheet_name="Completeness")
    monthly_avg.to_frame(name="Average").to_excel(writer, sheet_name="Monthly Average")

print("\nReport generated successfully!")
print(f"Saved at: {output_path}")