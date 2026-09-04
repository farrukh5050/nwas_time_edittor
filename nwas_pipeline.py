"""Shared data handling for the NWAS time editor.

The column lists, reference splitting, time parsing and merge all live here,
so there is one definition of how the two exports become the output workbook.

This module deliberately imports nothing from the rest of the project, so
combine_nwas_ghost, portal and gui can all depend on it without a cycle.
"""

import os
import re

import pandas as pd

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

# Which column set each slot in the UI is validated against.
EXPECTED = {"nwas": NWAS_COLUMNS, "ghost": GHOST_COLUMNS}

# The CSV export writes times as "27/08/2026 13:38". Excel hands over real
# datetimes already, so _ensure_datetime() only parses when it has to.
GHOST_TIME_COLUMNS = ["Date/Time", "Vehicle Arrived at Time", "Completed at Time"]
GHOST_TIME_FORMAT = "%d/%m/%Y %H:%M"

# Output filenames the app reads and writes.
CLEANED_FILE = "cleaned_ghost_data.xlsx"
MODIFIED_FILE = "Modified_NWAS_File.xlsx"
MODIFIED_SHEET = "Modified_NWAS"

# The three columns the portal upload file needs.
UPLOAD_COLUMNS = ["Your Reference 1", "Vehicle Arrived at Time", "Completed at Time"]

MERGE_COLUMNS = [
    "Your Reference 1",
    "Office Note",
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


class PipelineError(Exception):
    """Raised for problems worth showing the user verbatim."""


# --------------------------------------------------------------------- reading


def read_export(path, nrows=None, dtype=None):
    """Read a CSV or Excel export into a DataFrame."""
    if os.path.splitext(path)[1].lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, sheet_name=0, nrows=nrows, dtype=dtype)
    return pd.read_csv(path, low_memory=False, nrows=nrows, dtype=dtype)


def inspect(path):
    """Read the header once and report what the file can be used as.

    Returns {"kind": "nwas" | "ghost" | None, "missing": {kind: [columns]}}.
    """
    try:
        columns = set(read_export(path, nrows=0).columns)
    except Exception as exc:
        raise PipelineError(f"Could not read '{os.path.basename(path)}': {exc}")

    missing = {
        kind: [c for c in expected if c not in columns]
        for kind, expected in EXPECTED.items()
    }
    kind = next((k for k, short in missing.items() if not short), None)
    return {"kind": kind, "missing": missing}


# -------------------------------------------------------------- transformation


def clean_and_split_ref(ref):
    """Strip everything from a reference field except digits and separators."""
    ref = re.sub(r"[\t\s]+", "", str(ref))
    return re.sub(r"[^0-9+-]", "", ref)


def split_and_duplicate_rows(df):
    """One row per reference, for ghost rows covering several journeys."""
    new_rows = []
    for _, row in df.iterrows():
        for ref in re.split(r"[+\-]", clean_and_split_ref(row["Your Reference 1"])):
            if ref.strip():
                # astype(object) so the int below can be written regardless of
                # the caller's column dtypes (an all-str frame rejects it).
                new_row = row.astype(object)
                new_row["Your Reference 1"] = int(ref.strip())
                new_rows.append(new_row)
    return pd.DataFrame(new_rows)


def fix_time_format(time_value):
    """Normalise a journey time to HH:MM:SS, or None if it will not parse."""
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return pd.to_datetime(time_value, format=fmt).strftime("%H:%M:%S")
        except ValueError:
            continue
    return None


def _ensure_datetime(series):
    """Parse the CSV date format, but leave Excel's real datetimes alone."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    return pd.to_datetime(series, format=GHOST_TIME_FORMAT, errors="coerce")


def combine_frames(nwas_raw, ghost_raw, progress=None):
    """Merge the two exports. Returns (NWAS_data, ghost_data)."""

    def step(message):
        if progress:
            progress(message)

    nwas_data = nwas_raw[NWAS_COLUMNS]
    ghost_data = ghost_raw[GHOST_COLUMNS]

    step("Parsing booking times...")
    for column in GHOST_TIME_COLUMNS:
        ghost_data[column] = _ensure_datetime(ghost_data[column])

    step(f"Splitting multi-journey references ({len(ghost_data)} rows)...")
    ghost_data = split_and_duplicate_rows(ghost_data)

    step("Normalising journey times...")
    nwas_data["Journey Time"] = (
        nwas_data["Journey Time"].astype(str).apply(fix_time_format)
    )

    step("Matching bookings to journeys...")
    nwas_data = nwas_data.merge(
        ghost_data[MERGE_COLUMNS],
        how="left",
        left_on="ID",
        right_on="Your Reference 1",
    )
    nwas_data.drop(columns=["Your Reference 1"], inplace=True)
    nwas_data.rename(columns=RENAMES, inplace=True)

    # Selecting "Office Note" keeps both same-named columns (the ghost note and
    # the NWAS journey note). calc_time_dif's hardcoded G/J/K cells depend on
    # that extra column being there, so leave it alone.
    return nwas_data[FINAL_COLUMNS], ghost_data


def build_frames(nwas_path, ghost_path, progress=None):
    """Read the two separate exports, then combine them."""
    if progress:
        progress("Reading NWAS journeys...")
    nwas_raw = read_export(nwas_path, dtype={"Journey Time": str})

    if progress:
        progress("Reading ghost bookings...")
    ghost_raw = read_export(ghost_path)

    return combine_frames(nwas_raw, ghost_raw, progress=progress)


def clean_ghost_export(path, progress=None):
    """Turn a raw ghost export into the portal upload frame."""
    if progress:
        progress("Reading ghost export...")
    df = read_export(path)

    missing = [c for c in UPLOAD_COLUMNS if c not in df.columns]
    if missing:
        raise PipelineError(
            "'%s' is missing these columns: %s"
            % (os.path.basename(path), ", ".join(missing))
        )

    df = df[UPLOAD_COLUMNS]
    for column in ("Vehicle Arrived at Time", "Completed at Time"):
        df[column] = pd.to_datetime(
            df[column], format=GHOST_TIME_FORMAT, errors="coerce"
        ).dt.strftime("%H:%M")

    if progress:
        progress(f"Splitting multi-journey references ({len(df)} rows)...")
    return split_and_duplicate_rows(df)


def cleaned_from_modified(path, progress=None):
    """Turn a built Modified_NWAS workbook back into the portal upload frame."""
    if progress:
        progress(f"Reading '{MODIFIED_SHEET}' sheet...")
    try:
        df = pd.read_excel(path, sheet_name=MODIFIED_SHEET)
    except Exception as exc:
        raise PipelineError(
            f"Could not read the '{MODIFIED_SHEET}' sheet of "
            f"'{os.path.basename(path)}': {exc}"
        )

    wanted = ["Jrny ID", "Vehicle Arrived at Time", "Completed at Time"]
    missing = [c for c in wanted if c not in df.columns]
    if missing:
        raise PipelineError(
            "'%s' is missing these columns: %s"
            % (os.path.basename(path), ", ".join(missing))
        )

    cleaned = df[wanted].copy()
    cleaned.rename(columns={"Jrny ID": "Your Reference 1"}, inplace=True)

    def safe_parse_time(col):
        return (
            pd.to_datetime(col, format="%H:%M", errors="coerce")
            .fillna(pd.to_datetime(col, format="%H:%M:%S", errors="coerce"))
            .dt.strftime("%H:%M")
        )

    cleaned["Vehicle Arrived at Time"] = safe_parse_time(cleaned["Vehicle Arrived at Time"])
    cleaned["Completed at Time"] = safe_parse_time(cleaned["Completed at Time"])
    return cleaned
