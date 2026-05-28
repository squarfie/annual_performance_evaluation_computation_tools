import pandas as pd
import os
from tkinter import Tk, filedialog

# === STEP 1: Select file ===
Tk().withdraw()
file_path = filedialog.askopenfilename(title="Select ONTIME submission Excel file")

if not file_path:
    print("No file selected.")
    exit()

# === STEP 2: Load file ===
df = pd.read_excel(file_path)

# First column = site codes
site_col = df.columns[0]
month_cols = df.columns[1:]  # Jan → Dec

# === STEP 3: Function to compute deadline ===
def get_deadline(year, month):
    # next month
    if month == 12:
        next_month = 1
        year += 1
    else:
        next_month = month + 1

    # last day of next month
    return pd.Timestamp(year, next_month, 1) + pd.offsets.MonthEnd(0)

# === STEP 4: Convert month names to numbers ===
month_map = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
    "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
    "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
}

# === Ask user for year ===
while True:
    try:
        YEAR = int(input("Enter reporting year (e.g. 2025): ").strip())
        if 1900 <= YEAR <= 2100:
            break
        else:
            print("Please enter a valid year (1900–2100).")
    except ValueError:
        print("Invalid input. Enter a numeric year.")

# === STEP 5: Compute ON-TIME matrix ===
result = []

for _, row in df.iterrows():
    site = row[site_col]
    row_result = {"Site": site}

    ontime_values = []
    valid_count = 0

    for col in month_cols:
        month_name = col[:3]
        month_num = month_map.get(month_name)

        value = row[col]

        # 🔴 Skip blanks completely
        if pd.isna(value) or str(value).strip() == "":
            row_result[col] = None  # or "" if you prefer
            continue

        try:
            submission_date = pd.to_datetime(value)
        except:
            row_result[col] = None
            continue

        deadline = get_deadline(YEAR, month_num)

        if submission_date <= deadline:
            row_result[col] = 1
            ontime_values.append(1)
        else:
            row_result[col] = 0
            ontime_values.append(0)

        valid_count += 1  # ✅ only count valid months

    # === Compute summary ===
    row_result["SUM"] = sum(ontime_values)

    if ontime_values:
        avg = sum(ontime_values) / len(ontime_values)
        row_result["AVG"] = round(avg, 2)
        row_result["%"] = round(avg * 100, 0)  # matches your Excel %
    else:
        row_result["AVG"] = 0
        row_result["%"] = 0

    result.append(row_result)

# === STEP 6: Convert to DataFrame ===
result_df = pd.DataFrame(result)

# === STEP 7: Save output ===
output_path = os.path.join(os.path.dirname(file_path), "on_time_submission_report.xlsx")

with pd.ExcelWriter(output_path) as writer:
    result_df.to_excel(writer, index=False, sheet_name="OnTime")

print("\n✅ Report generated!")
print(f"Saved at: {output_path}")