import pandas as pd
import os
from tkinter import Tk, filedialog

# =========================================
# MONTH MAP
# =========================================
month_lookup = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
    "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
    "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"
}

ordered_months = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
]


def extract_month(filename):
    return month_lookup.get(filename[1:3], filename)


# =========================================
# MAIN PROCESS FUNCTION
# =========================================
def process_folder(folder_path):

    files = sorted([
        f for f in os.listdir(folder_path)
        if f.endswith(".xlsx")
    ])

    if not files:
        print("No Excel files found.")
        return

    site_data = {}

    for file in files:

        file_path = os.path.join(folder_path, file)

        try:

            df = pd.read_excel(file_path)

            # Replace blank cells
            df.replace("", pd.NA, inplace=True)

            # =========================================
            # HELPER FUNCTION
            # =========================================
            def get_first_valid(series):

                series = (
                    series
                    .dropna()
                    .astype(str)
                    .str.strip()
                )

                series = series[series != ""]

                return (
                    series.iloc[0]
                    if not series.empty
                    else None
                )

            # =========================================
            # GET SITE CODE
            # =========================================
            site_code = "UNKNOWN"

            if "LABORATORY" in df.columns:

                val = get_first_valid(df["LABORATORY"])

                if val:
                    site_code = val

            elif "INSTITUT" in df.columns:

                val = get_first_valid(df["INSTITUT"])

                if val:
                    site_code = val

            site_code = site_code.upper()

            print(f"\nProcessing: {file}")
            print(f"SITE: {site_code}")

            # =========================================
            # GET MONTH
            # =========================================
            month = extract_month(file)

            # =========================================
            # INITIALIZE SITE
            # =========================================
            if site_code not in site_data:
                site_data[site_code] = {}

            # =========================================
            # CHECK ORGANISM COLUMN
            # =========================================
            if "ORGANISM" not in df.columns:

                print("ORGANISM column not found.")

                score = 0

            else:

                # =========================================
                # CLEAN ORGANISM COLUMN
                # =========================================
                organism_series = (
                    df["ORGANISM"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.lower()
                )

                # =========================================
                # TOTAL RECORDS
                # =========================================
                total_records = len(organism_series)

                # =========================================
                # COUNT XXX
                # =========================================
                xxx_count = (
                    organism_series == "xxx"
                ).sum()

                # =========================================
                # COMPUTE XXX %
                # =========================================
                if total_records == 0:

                    xxx_percent = 0

                else:

                    xxx_percent = (
                        xxx_count / total_records
                    ) * 100

                # =========================================
                # SCORING LOGIC
                # =========================================
                #
                # 0% xxx      -> 0
                # <10% xxx    -> 50
                # >=10% xxx   -> 100
                #
                # =========================================
                if xxx_percent == 0:

                    score = 0

                elif xxx_percent < 10:

                    score = 50

                else:

                    score = 100

                # =========================================
                # DISPLAY RESULTS
                # =========================================
                print(f"Total Records : {total_records}")
                print(f"XXX Count     : {xxx_count}")
                print(f"XXX %         : {xxx_percent:.2f}%")
                print(f"Final Score   : {score}")

            # =========================================
            # SAVE SCORE
            # =========================================
            site_data[site_code][month] = score

        except Exception as e:

            print(f"Error processing {file}: {e}")

    # =========================================
    # GENERATE OUTPUT FILE
    # =========================================
    for site, month_scores in site_data.items():

        df_out = pd.DataFrame(
            [month_scores],
            index=["No Growth Encoding"]
        )

        # Arrange month order
        df_out = df_out.reindex(columns=ordered_months)

        # Fill empty months
        df_out = df_out.fillna(0)

        # Average
        df_out["Average %"] = (
            df_out.mean(axis=1)
            .round(0)
        )

        # =========================================
        # OUTPUT FILE
        # =========================================
        output_file = os.path.join(
            folder_path,
            f"no_growth_encoding_{site.lower()}.xlsx"
        )

        with pd.ExcelWriter(output_file) as writer:

            df_out.to_excel(
                writer,
                sheet_name="No Growth Encoding"
            )

        print(f"\n✔ Saved: {output_file}")


# =========================================
# LOOP EXECUTION
# =========================================
while True:

    Tk().withdraw()

    folder_path = filedialog.askdirectory(
        title="Select folder with monthly Excel files"
    )

    if not folder_path:

        print("No folder selected. Exiting...")
        break

    process_folder(folder_path)

    choice = input(
        "\nConvert another folder? (Y/N): "
    ).strip().lower()

    if choice != "y":

        print("Done.")
        break