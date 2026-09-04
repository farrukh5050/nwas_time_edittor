import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def get_excel_file():
    print("Opening file selector...")
    root = Tk()
    root.lift()
    root.attributes("-topmost", True)
    root.withdraw()

    file_path = askopenfilename(
        title="Import NWAS File", filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    root.destroy()  # Close Tk instance properly

    if not file_path:
        print("No file selected.")
        return None

    print(f"File selected: {file_path}")

    NWAS_data = pd.read_excel(file_path, sheet_name=1, dtype={"Journey Time": str})
    ghost_data = pd.read_excel(file_path, sheet_name=0)

    columns_to_keep = [
        "Name",
        "Date/Time",
        "Pickup",
        "Destination",
        "Office Note",
        "Your Reference 1",
        "Vehicle Arrived at Time",
        "Completed at Time",
    ]
    ghost_data = ghost_data[columns_to_keep]

    def clean_and_split_ref(ref):
        ref = re.sub(r"[\t\s]+", "", str(ref))
        ref = re.sub(r"[^0-9+-]", "", ref)
        return ref

    def split_and_duplicate_rows(df):
        new_rows = []
        for _, row in df.iterrows():
            cleaned_ref = clean_and_split_ref(row["Your Reference 1"])
            refs = re.split(r"[+\-]", cleaned_ref)
            for ref in refs:
                if ref.strip():
                    new_row = row.copy()
                    new_row["Your Reference 1"] = int(ref.strip())
                    new_rows.append(new_row)
        return pd.DataFrame(new_rows)

    ghost_data = split_and_duplicate_rows(ghost_data)

    nwas_cols_to_keep = [
        "ID",
        "Passenger Name",
        "From Road",
        "To Road",
        "Journey Notes",
        "Journey Time",
        "Category Text",
    ]
    NWAS_data = NWAS_data[nwas_cols_to_keep]

    def fix_time_format(time_value):
        try:
            return pd.to_datetime(time_value, format="%H:%M:%S").strftime("%H:%M:%S")
        except ValueError:
            try:
                return pd.to_datetime(time_value, format="%H:%M").strftime("%H:%M:%S")
            except ValueError:
                return None

    NWAS_data["Journey Time"] = (
        NWAS_data["Journey Time"].astype(str).apply(fix_time_format)
    )

    NWAS_data = NWAS_data.merge(
        ghost_data[
            [
                "Your Reference 1",
                "Office Note",
                "Vehicle Arrived at Time",
                "Completed at Time",
            ]
        ],
        how="left",
        left_on="ID",
        right_on="Your Reference 1",
    )

    NWAS_data.drop(columns=["Your Reference 1"], inplace=True)

    NWAS_data.rename(
        columns={
            "ID": "Jrny ID",
            "Passenger Name": "Name",
            "From Road": "From Hospital Text",
            "To Road": "To Hospital Text",
            "Journey Notes": "Office Note",
            "Journey Time": "Time",
            "Category Text": "Cat",
        },
        inplace=True,
    )

    final_cols = [
        "Jrny ID",
        "Name",
        "From Hospital Text",
        "To Hospital Text",
        "Office Note",
        "Time",
        "Cat",
        "Vehicle Arrived at Time",
        "Completed at Time",
    ]
    NWAS_data = NWAS_data[final_cols]

    return NWAS_data, ghost_data


def calc_time_dif(NWAS_data, ghost_data):
    NWAS_data["Time"] = pd.to_datetime(
        NWAS_data["Time"], format="%H:%M:%S", errors="coerce"
    ).dt.time
    NWAS_data["Completed at Time"] = pd.to_datetime(
        NWAS_data["Completed at Time"], format="%Y-%m-%d %H:%M:%S", errors="coerce"
    ).dt.time

    def calculate_time_difference(time1, time2):
        if pd.isna(time1) or pd.isna(time2):
            return ""
        time1_seconds = time1.hour * 3600 + time1.minute * 60 + time1.second
        time2_seconds = time2.hour * 3600 + time2.minute * 60 + time2.second
        return (time2_seconds - time1_seconds) // 60

    NWAS_data["Time Difference (mins)"] = NWAS_data.apply(
        lambda row: calculate_time_difference(row["Time"], row["Completed at Time"]),
        axis=1,
    )

    def categorize_time_diff(row):
        diff = row["Time Difference (mins)"]
        category = row["Cat"]
        if diff == "":
            return "No Data"
        if diff < 0:
            if category in ["DIA", "Oncology"] and diff >= -45:
                return "Fine (Green)"
            elif category not in ["DIA", "Oncology"] and diff >= -60:
                return "Fine (Green)"
            else:
                return "Too Early (No Highlight)"
        if diff > 1:
            return "Late (Yellow)"
        return "On Time"

    NWAS_data["Status"] = NWAS_data.apply(categorize_time_diff, axis=1)

    output_filename = "Modified_NWAS_File.xlsx"
    with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
        NWAS_data.to_excel(writer, sheet_name="Modified_NWAS", index=False)
        ghost_data.to_excel(writer, sheet_name="Ghost_Data", index=False)

    wb = load_workbook(output_filename)
    ws = wb["Modified_NWAS"]

    green_fill = PatternFill(
        start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
    )
    yellow_fill = PatternFill(
        start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
    )
    red_fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")

    status_col = ws.max_column

    for row in range(2, ws.max_row + 1):
        status_value = ws.cell(row=row, column=status_col).value
        time_value = ws[f"G{row}"].value
        completed_value = ws[f"J{row}"].value

        if status_value == "Fine (Green)":
            ws.cell(row=row, column=status_col).fill = green_fill
        elif status_value == "Late (Yellow)":
            ws.cell(row=row, column=status_col).fill = yellow_fill

        if time_value and completed_value:
            time_value = pd.to_datetime(time_value).time()
            completed_value = pd.to_datetime(completed_value).time()
            if time_value > completed_value:
                ws[f"K{row}"].fill = red_fill

    wb.save(output_filename)
    print(f"File updated successfully Choomb: '{output_filename}'")


def export_cleaned_ghost_data():
    print("Opening file selector...")
    root = Tk()
    root.lift()
    root.attributes("-topmost", True)
    root.withdraw()

    file_path = askopenfilename(
        title="Select Modified_NWAS_File.xlsx",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    root.destroy()

    if not file_path:
        print("No file selected.")
        return

    print(f"File selected: {file_path}")

    # Load the Modified_NWAS sheet
    try:
        df = pd.read_excel(file_path, sheet_name="Modified_NWAS")
    except Exception as e:
        print(f"Failed to load 'Modified_NWAS' sheet: {e}")
        return

    try:
        cleaned = df[["Jrny ID", "Vehicle Arrived at Time", "Completed at Time"]].copy()
        cleaned.rename(columns={"Jrny ID": "Your Reference 1"}, inplace=True)

        def safe_parse_time(col):
            return (
                pd.to_datetime(col, format="%H:%M", errors="coerce")
                .fillna(pd.to_datetime(col, format="%H:%M:%S", errors="coerce"))
                .dt.strftime("%H:%M")
            )

        cleaned["Vehicle Arrived at Time"] = safe_parse_time(
            cleaned["Vehicle Arrived at Time"]
        )
        cleaned["Completed at Time"] = safe_parse_time(cleaned["Completed at Time"])

        cleaned.to_excel("cleaned_ghost_data.xlsx", index=False)
        print("\nCleaned ghost data saved as 'cleaned_ghost_data.xlsx'")

    except KeyError as e:
        print(f"Expected columns not found: {e}")


if __name__ == "__main__":
    print("What would you like to do?")
    print("1: Combine NWAS and Ghost data (full process)")
    print("2: Create cleaned Ghost data from Modified_NWAS_File.xlsx")
    choice = input("Enter option [1/2]: ").strip()

    if choice == "2":
        export_cleaned_ghost_data()
    elif choice == "1":
        result = get_excel_file()
        if result is not None:
            NWAS_data, ghost_data = result
            calc_time_dif(NWAS_data=NWAS_data, ghost_data=ghost_data)
    else:
        print("Invalid choice. Exiting.")
