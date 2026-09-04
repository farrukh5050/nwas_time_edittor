import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

def split_and_duplicate_rows(df):
    """Split rows where 'Your Reference 1' contains multiple references."""
    new_rows = []

    for _, row in df.iterrows():
        ref_raw = str(row['Your Reference 1'])
        cleaned_ref = re.sub(r'[^0-9+-]', '', re.sub(r'[\t\s]+', '', ref_raw))
        refs = re.split(r'[+\-]', cleaned_ref)
        for ref in refs:
            if ref.strip():
                new_row = row.copy()
                new_row['Your Reference 1'] = ref.strip()
                new_rows.append(new_row)

    return pd.DataFrame(new_rows)

def select_excel_file():
    """Open file dialog to select an Excel or CSV file."""
    Tk().withdraw()
    file_path = askopenfilename(
        title="Import Ghost File",
        filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv")]
    )

    if not file_path:
        print("No file selected.")
        return None

    # Determine file type
    if file_path.lower().endswith('.csv'):
        df = pd.read_csv(file_path, low_memory=False)
    else:
        df = pd.read_excel(file_path, sheet_name=0)

    # Keep required columns only
    columns_to_keep = ["Your Reference 1", "Vehicle Arrived at Time", "Completed at Time"]
    df = df[columns_to_keep]

    # Format time columns to HH:MM
    time_columns = ["Vehicle Arrived at Time", "Completed at Time"]
    for col in time_columns:
        df[col] = pd.to_datetime(df[col], format='%d/%m/%Y %H:%M', errors='coerce').dt.strftime('%H:%M')

    cleaned_df = split_and_duplicate_rows(df)
    return cleaned_df
