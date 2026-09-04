"""Build the Modified_NWAS workbook from the two raw CSV exports.

combine_nwas_ghost.get_excel_file() takes a single workbook holding both sheets.
This module takes the two separate exports instead, works out which file is
which from its column headers (filenames vary), and returns the same pair of
frames so combine_nwas_ghost.calc_time_dif() can write the workbook unchanged.
"""

import os
import re

import pandas as pd

import combine_nwas_ghost

# Columns each export must provide, and which are carried into the output.
NWAS_COLUMNS = [
    "ID",
    "Passenger Name",
    "From Road",
    "To Road",
    "Journey Notes",
    "Journey Time",
    "Category Text",
]
GHOST_COLUMNS = [
    "Name",
    "Date/Time",
    "Pickup",
    "Destination",
    "Office Note",
    "Your Reference 1",
    "Vehicle Arrived at Time",
    "Completed at Time",
]

# The ghost export writes times as "27/08/2026 13:38". They have to become real
# datetimes here, because calc_time_dif() expects what Excel used to hand it.
GHOST_TIME_COLUMNS = ["Date/Time", "Vehicle Arrived at Time", "Completed at Time"]
GHOST_TIME_FORMAT = "%d/%m/%Y %H:%M"

# Which column set each slot in the UI is validated against.
EXPECTED = {"nwas": NWAS_COLUMNS, "ghost": GHOST_COLUMNS}

FINAL_COLUMNS = [
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

RENAMES = {
    "ID": "Jrny ID",
    "Passenger Name": "Name",
    "From Road": "From Hospital Text",
    "To Road": "To Hospital Text",
    "Journey Notes": "Office Note",
    "Journey Time": "Time",
    "Category Text": "Cat",
}


class PipelineError(Exception):
    """Raised for problems worth showing the user verbatim."""


def _read(path, nrows=None, dtype=None):
    """Read a CSV or Excel export into a DataFrame."""
    if os.path.splitext(path)[1].lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, sheet_name=0, nrows=nrows, dtype=dtype)
    return pd.read_csv(path, low_memory=False, nrows=nrows, dtype=dtype)


def identify(path):
    """Return "nwas", "ghost", or None by looking at the file's columns."""
    try:
        columns = set(_read(path, nrows=0).columns)
    except Exception as exc:
        raise PipelineError(f"Could not read '{os.path.basename(path)}': {exc}")

    if set(NWAS_COLUMNS) <= columns:
        return "nwas"
    if set(GHOST_COLUMNS) <= columns:
        return "ghost"
    return None


def missing_for(path, kind):
    """Return the columns "kind" requires that this file does not have."""
    try:
        columns = set(_read(path, nrows=0).columns)
    except Exception as exc:
        raise PipelineError(f"Could not read '{os.path.basename(path)}': {exc}")
    return [c for c in EXPECTED[kind] if c not in columns]


def missing_columns(path):
    """Return the shortfall against whichever export the file most resembles."""
    nwas_short = missing_for(path, "nwas")
    ghost_short = missing_for(path, "ghost")
    return nwas_short if len(nwas_short) <= len(ghost_short) else ghost_short


def clean_and_split_ref(ref):
    ref = re.sub(r"[\t\s]+", "", str(ref))
    return re.sub(r"[^0-9+-]", "", ref)


def split_and_duplicate_rows(df):
    """One row per reference, for ghost rows covering several journeys."""
    new_rows = []
    for _, row in df.iterrows():
        for ref in re.split(r"[+\-]", clean_and_split_ref(row["Your Reference 1"])):
            if ref.strip():
                new_row = row.copy()
                new_row["Your Reference 1"] = int(ref.strip())
                new_rows.append(new_row)
    return pd.DataFrame(new_rows)


def fix_time_format(time_value):
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return pd.to_datetime(time_value, format=fmt).strftime("%H:%M:%S")
        except ValueError:
            continue
    return None


def build_frames(nwas_path, ghost_path, progress=None):
    """Return (NWAS_data, ghost_data) ready for calc_time_dif()."""
    def step(message):
        if progress:
            progress(message)

    step("Reading NWAS journeys...")
    nwas_data = _read(nwas_path, dtype={"Journey Time": str})[NWAS_COLUMNS]

    step("Reading ghost bookings...")
    ghost_data = _read(ghost_path)[GHOST_COLUMNS]

    step("Parsing booking times...")
    for column in GHOST_TIME_COLUMNS:
        ghost_data[column] = pd.to_datetime(
            ghost_data[column], format=GHOST_TIME_FORMAT, errors="coerce"
        )

    step(f"Splitting multi-journey references ({len(ghost_data)} rows)...")
    ghost_data = split_and_duplicate_rows(ghost_data)

    step("Normalising journey times...")
    nwas_data["Journey Time"] = (
        nwas_data["Journey Time"].astype(str).apply(fix_time_format)
    )

    step("Matching bookings to journeys...")
    nwas_data = nwas_data.merge(
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
    nwas_data.drop(columns=["Your Reference 1"], inplace=True)
    nwas_data.rename(columns=RENAMES, inplace=True)

    # Selecting "Office Note" keeps both same-named columns (the ghost note and
    # the NWAS journey note). calc_time_dif's hardcoded G/J/K cells depend on
    # that extra column being there, so leave it alone.
    nwas_data = nwas_data[FINAL_COLUMNS]
    return nwas_data, ghost_data


def build_workbook(nwas_path, ghost_path, output_path, progress=None):
    """Run the whole job and write the workbook. Returns (nwas_rows, ghost_rows)."""
    nwas_data, ghost_data = build_frames(nwas_path, ghost_path, progress=progress)
    if progress:
        progress("Writing workbook...")
    combine_nwas_ghost.calc_time_dif(
        nwas_data, ghost_data, output_filename=output_path
    )
    return len(nwas_data), len(ghost_data)
