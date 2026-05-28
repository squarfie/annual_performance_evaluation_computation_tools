import pandas as pd
import os
from tkinter import Tk, filedialog

# =========================================
# MONTH ORDER
# =========================================
ordered_months = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
]

# =========================================
# SELECT FOLDER
# =========================================
Tk().withdraw()

folder_path = filedialog.askdirectory(
    title="Select folder containing no_growth_encoding files"
)

if not folder_path:
    print("No folder selected.")
    exit()

# =========================================
# FIND FILES
# =========================================
files = [
    f for f in os.listdir(folder_path)
    if f.startswith("no_growth_encoding_")
    and f.endswith(".xlsx")
]

if not files:
    print("No no_growth_encoding files found.")
    exit()

# =========================================
# CONSOLIDATED DATA
# =========================================
consolidated_data = []

# =========================================
# PROCESS EACH FILE
# =========================================
for file in sorted(files):

    file_path = os.path.join(folder_path, file)

    try:

        # =========================================
        # READ EXCEL
        # =========================================
        df = pd.read_excel(file_path)

        # =========================================
        # GET SITE NAME
        # =========================================
        #
        # filename example:
        # no_growth_encoding_zph.xlsx
        #
        site = (
            file
            .replace("no_growth_encoding_", "")
            .replace(".xlsx", "")
            .upper()
        )

        # =========================================
        # GET MONTH VALUES
        # =========================================
        row_data = {"Site": site}

        for month in ordered_months:

            if month in df.columns:

                value = df.loc[0, month]

                # convert to %
                row_data[month] = f"{int(value)}%"

            else:

                row_data[month] = "0%"

        # =========================================
        # GET AVERAGE
        # =========================================
        if "Average %" in df.columns:

            avg_value = df.loc[0, "Average %"]

            row_data["Average"] = f"{int(avg_value)}%"

        else:

            row_data["Average"] = "0%"

        consolidated_data.append(row_data)

        print(f"Processed: {site}")

    except Exception as e:

        print(f"Error processing {file}: {e}")

# =========================================
# CREATE CONSOLIDATED TABLE
# =========================================
df_consolidated = pd.DataFrame(consolidated_data)

# =========================================
# SORT BY AVERAGE DESCENDING
# =========================================
def extract_percent(value):

    return int(str(value).replace("%", ""))

df_consolidated["SortValue"] = (
    df_consolidated["Average"]
    .apply(extract_percent)
)

df_consolidated = (
    df_consolidated
    .sort_values(by="SortValue", ascending=False)
    .drop(columns=["SortValue"])
)

# =========================================
# CREATE RANKING TABLE
# =========================================
df_ranking = (
    df_consolidated[["Site", "Average"]]
    .copy()
)

# =========================================
# SAVE OUTPUT
# =========================================
output_file = os.path.join(
    folder_path,
    "negatives_consolidated_report.xlsx"
)

with pd.ExcelWriter(output_file) as writer:

    # Main summary
    df_consolidated.to_excel(
        writer,
        sheet_name="negatives Summary",
        index=False,
        startrow=0,
        startcol=0
    )

    # Ranking table
    df_ranking.to_excel(
        writer,
        sheet_name="negatives Ranking",
        index=False,
        startrow=0,
        startcol=0,
    )

print("\n===================================")
print("✔ CONSOLIDATED REPORT GENERATED")
print(output_file)
print("===================================")