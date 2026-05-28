import pandas as pd
import os
from tkinter import Tk, filedialog

# =========================================================
# MAIN PROCESS FUNCTION
# =========================================================
def process_files():

    # =====================================================
    # SELECT FOLDER WITH MONTHLY EXCEL FILES
    # =====================================================
    Tk().withdraw()

    monthly_folder = filedialog.askdirectory(
        title="Select folder containing monthly Excel files"
    )

    if not monthly_folder:
        print("No monthly folder selected.")
        return

    # =====================================================
    # SELECT MAPPING FILE
    # =====================================================
    mapping_file = filedialog.askopenfilename(
        title="Select org_groupings_whonet_completed.xlsx",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not mapping_file:
        print("No mapping file selected.")
        return

    # =====================================================
    # LOAD MAPPING FILE
    # =====================================================
    try:
        map_df = pd.read_excel(mapping_file)

        # normalize column names
        map_df.columns = map_df.columns.str.strip()

        # required columns
        required_cols = ["Whonet_Org_Code", "Panel_Group_arsp"]

        for col in required_cols:
            if col not in map_df.columns:
                raise ValueError(f"Missing column in mapping file: {col}")

        # normalize organism codes
        map_df["Whonet_Org_Code"] = (
            map_df["Whonet_Org_Code"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

        # create dictionary
        org_mapping = dict(
            zip(
                map_df["Whonet_Org_Code"],
                map_df["Panel_Group_arsp"]
            )
        )

    except Exception as e:
        print(f"Error reading mapping file: {e}")
        return

    # =====================================================
    # OUTPUT FOLDER
    # =====================================================
    output_folder = os.path.join(
        monthly_folder,
        "with_panel_group"
    )

    os.makedirs(output_folder, exist_ok=True)

    # =====================================================
    # PROCESS MONTHLY FILES
    # =====================================================
    files = [
        f for f in os.listdir(monthly_folder)
        if f.lower().endswith(".xlsx")
    ]

    if not files:
        print("No Excel files found.")
        return

    for file in files:

        # skip already generated files
        if file.startswith("with_panel_"):
            continue

        file_path = os.path.join(monthly_folder, file)

        try:
            print(f"\nProcessing: {file}")

            df = pd.read_excel(file_path)

            # normalize columns
            df.columns = df.columns.str.strip()

            # =================================================
            # FIND ORGANISM COLUMN
            # =================================================
            organism_col = None

            for col in df.columns:
                if col.strip().lower() == "organism":
                    organism_col = col
                    break

            if organism_col is None:
                print(f"⚠ ORGANISM column not found in {file}")
                continue

            # =================================================
            # NORMALIZE ORGANISM VALUES
            # =================================================
            df[organism_col] = (
                df[organism_col]
                .astype(str)
                .str.strip()
                .str.lower()
            )

            # =================================================
            # MAP PANEL GROUP
            # =================================================
            df["Panel_Group_arsp"] = df[organism_col].map(org_mapping)

            # replace NaN with blank
            df["Panel_Group_arsp"] = (
                df["Panel_Group_arsp"]
                .fillna("")
            )

            # =================================================
            # SAVE OUTPUT
            # =================================================
            output_file = os.path.join(
                output_folder,
                f"with_panel_{file}"
            )

            df.to_excel(output_file, index=False)

            print(f"✔ Saved: {output_file}")

        except Exception as e:
            print(f"❌ Error processing {file}: {e}")

    print("\n✅ All files processed successfully!")


# =========================================================
# LOOP EXECUTION
# =========================================================
while True:

    process_files()

    choice = input(
        "\nDo you want to process another folder? (Y/N): "
    ).strip().lower()

    if choice != "y":
        print("Program exited.")
        break