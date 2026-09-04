import os
import time
import calendar
from datetime import date
import pandas as pd
from io import StringIO
from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import requests
from dotenv import load_dotenv
import load_files
import combine_nwas_ghost

# Load credentials from .env file
load_dotenv()

nwas_username = os.getenv("NWAS_USERNAME")
nwas_password = os.getenv("NWAS_PASSWORD")
if not nwas_username or not nwas_password:
    raise RuntimeError(
        "Missing credentials: set NWAS_USERNAME and NWAS_PASSWORD in your .env file."
    )

# Re-bound as plain str so type checkers know they cannot be None past this point
NWAS_USERNAME: str = nwas_username
NWAS_PASSWORD: str = nwas_password

# Store the browser driver so it can be reused
driver: WebDriver | None = None
CLEANED_FILE = "cleaned_ghost_data.xlsx"


def load_and_clean_file():
    cleaned_df = load_files.select_excel_file()
    if cleaned_df is not None:
        cleaned_df.to_excel(CLEANED_FILE, index=False)
        print(f"Cleaned file saved as '{CLEANED_FILE}'.")


def combine_nwas_and_ghost():
    """Build Modified_NWAS_File.xlsx from a raw NWAS/Ghost workbook."""
    result = combine_nwas_ghost.get_excel_file()
    if result is None:
        return
    nwas_data, ghost_data = result
    combine_nwas_ghost.calc_time_dif(NWAS_data=nwas_data, ghost_data=ghost_data)


def export_cleaned_from_modified():
    """Write the portal upload file from an existing Modified_NWAS_File.xlsx."""
    combine_nwas_ghost.export_cleaned_ghost_data()


def open_chrome_and_login():
    global driver
    if driver:
        print("Chrome is already open.")
        return

    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("https://ptsed.nwas.nhs.uk/")
    driver.find_element(By.ID, "txtUsername").send_keys(NWAS_USERNAME)
    driver.find_element(By.ID, "txtPassword").send_keys(NWAS_PASSWORD)
    submitBtn = driver.find_element(By.ID, "cmdSubmit")
    submitBtn.click()
    WebDriverWait(driver, 30).until(EC.staleness_of(submitBtn))

    driver.get("https://ptsed.nwas.nhs.uk/frmQualityEntry.aspx")
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "txtPlanDate"))
    )
    input("Load dates on the page and press Enter to continue...")


def get_ids_from_table(driver: WebDriver, ghost_df):
    table_html = driver.find_element(By.XPATH, '//*[@id="tblResults"]').get_attribute("outerHTML")
    table_df = pd.read_html(StringIO(table_html))[0]
    table_df["ID"] = normalise_ids(table_df["ID"])
    merged = table_df.merge(ghost_df, how="left", left_on="ID", right_on="Your Reference 1")
    merged = merged.dropna(subset=["Vehicle Arrived at Time", "Completed at Time"])
    merged = merged.drop_duplicates(subset=["ID"])
    return merged[["ID", "Vehicle Arrived at Time", "Completed at Time"]]


def get_auth_cookie(driver: WebDriver):
    session = requests.Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie["name"], cookie["value"])
    return session


def update_ghost_times(session, matched_df, date_str):
    url = "https://ptsed.nwas.nhs.uk/webservice/wsrvJourney.asmx/ajaxQualityDataSave"
    headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
    }

    for _, row in matched_df.iterrows():
        payload = {
            "ID": row["ID"],
            "PickUp": f"{date_str} {row['Vehicle Arrived at Time']}",
            "DropOff": f"{date_str} {row['Completed at Time']}",
            "AbortID": "",
            "AbortText": "",
            "AbortNotes": "",
            "ExcepID": "",
            "ExcepText": "",
            "ExcepNotes": "",
            "FaultID": "",
            "FaultText": "",
            "AtLoc": "",
            "LeftLoc": "",
            "MobToDest": "",
            "AtDest": "",
            "Handover": "",
        }

        try:
            response = session.post(url, json={"Data": payload}, headers=headers, timeout=30)
            print(f"Updated ID {row['ID']} - Status: {response.status_code}")
            if response.status_code != 200:
                print("Error:", response.text)
        except Exception as e:
            print(f"Skipped ID {row['ID']} due to error: {e}")
            continue


def update_times(ghost_df=None):
    if driver is None:
        print("Please run option 2 first to open and log into Chrome.")
        return

    if ghost_df is None:
        try:
            ghost_df = load_ghost_file()
        except FileNotFoundError:
            print(f"'{CLEANED_FILE}' not found. Run option 1 first.")
            return

    try:
        selected_date = driver.find_element(By.ID, "txtPlanDate").get_attribute("value")
        matched_df = get_ids_from_table(driver, ghost_df)
        print(f"Date {selected_date}: {len(matched_df)} matched entries.")
        if matched_df.empty:
            return

        session = get_auth_cookie(driver)
        update_ghost_times(session, matched_df, date_str=selected_date)
        print("All times updated.")
    except Exception as e:
        print(f"Error during update: {e}")


def dismiss_message_box():
    """Click OK on the portal message box if it is showing. Returns True if dismissed."""
    if driver is None:
        return False
    try:
        msg_box = driver.find_element(By.XPATH, '//*[@id="gdivMsgBox"]/div')
        if msg_box.is_displayed():
            driver.find_element(By.ID, "gcmdMsgOK").click()
            print("Message box dismissed.")
            time.sleep(1)
            return True
    except Exception:
        pass
    return False


def wait_for_results(timeout=10):
    """Return True as soon as the results table is visible, or False if it never appears."""
    if driver is None:
        return False
    try:
        WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.ID, "tblResults"))
        )
        return True
    except TimeoutException:
        return False


def month_dates(year, month):
    """Return every day of the given month as DD/MM/YYYY strings."""
    num_days = calendar.monthrange(year, month)[1]
    return [date(year, month, day).strftime("%d/%m/%Y") for day in range(1, num_days + 1)]


def normalise_ids(series):
    """Coerce IDs to comparable digit strings: ' 12345 ' and 12345.0 both -> '12345'."""
    stripped = series.astype(str).str.replace(r"\s+", "", regex=True)
    return pd.to_numeric(stripped, errors="coerce").astype("Int64").astype(str)


def load_ghost_file(path=CLEANED_FILE):
    """Read the cleaned ghost file, with its reference column normalised for merging."""
    ghost_df = pd.read_excel(path)
    ghost_df["Your Reference 1"] = normalise_ids(ghost_df["Your Reference 1"])
    return ghost_df


def run_full_automation(year, month, call_signs=None, load_wait=10):
    """Loop over each call sign and every day of the month, search, then update times."""
    if driver is None:
        print("Please run option 2 first to open and log into Chrome.")
        return

    if call_signs is None:
        call_signs = [
            "STCDAM",
            "STCDEV",
            "STCDPM",
            "STCLAN",
            "STCPLAM",
            "STCPLAM2",
            "STCPLEV",
            "STCPLEV2",
            "STCPLPM",
            "STCPLPM2",
        ]

    dates = month_dates(year, month)

    try:
        ghost_df = load_ghost_file()
    except FileNotFoundError:
        print(f"'{CLEANED_FILE}' not found. Run option 1 first.")
        return

    for call_sign in call_signs:
        for date_str in dates:
            print(f"\n=== Call sign {call_sign} | Date {date_str} ===")
            try:
                # Enter call sign and select it from the autocomplete dropdown
                resource_field = driver.find_element(By.ID, "txtResource")
                resource_field.send_keys(Keys.CONTROL, "a")
                resource_field.send_keys(Keys.DELETE)
                resource_field.send_keys(call_sign)
                time.sleep(1)  # let the autocomplete dropdown appear
                resource_field.send_keys(Keys.ARROW_DOWN)
                resource_field.send_keys(Keys.ENTER)

                # Enter date as digits only (the field's mask inserts the "/")
                date_digits = date_str.replace("/", "")  # DDMMYYYY
                date_field = driver.find_element(By.ID, "txtPlanDate")
                date_field.send_keys(Keys.CONTROL, "a")
                date_field.send_keys(Keys.DELETE)
                date_field.send_keys(date_digits)

                # Click search
                driver.find_element(By.XPATH, "/html/body/form/div[5]/div[3]/button[1]").click()

                # Wait until the jobs table is visible (up to load_wait seconds)
                if wait_for_results(load_wait):
                    # Update times for the loaded jobs
                    update_times(ghost_df)
                else:
                    print(f"No jobs loaded for {call_sign} {date_str} - skipping.")

                # Dismiss any message box that popped up
                dismiss_message_box()

            except Exception as e:
                print(f"Skipped {call_sign} {date_str} due to error: {e}")
                dismiss_message_box()
                continue

    print("\nFull automation complete.")


def main_menu():
    while True:
        print("1. Load and clean a ghost file (Excel or CSV)")
        print("2. Open Chrome and login")
        print("3. Update times on the NWAS portal using cleaned file")
        print("4. Auto-run all call signs through a full month")
        print("5. Combine a raw NWAS/Ghost workbook into Modified_NWAS_File.xlsx")
        print("6. Create the cleaned ghost file from Modified_NWAS_File.xlsx")
        print("0. Exit")

        choice = input("Select an option (0-6): ").strip()

        if choice == "1":
            load_and_clean_file()
        elif choice == "2":
            open_chrome_and_login()
        elif choice == "3":
            update_times()
        elif choice == "4":
            try:
                month = int(input("Enter month number (1-12): ").strip())
                year = int(input("Enter year (e.g. 2026): ").strip())
                run_full_automation(year=year, month=month)
            except ValueError:
                print("Invalid month/year. Please enter numbers.")
        elif choice == "5":
            combine_nwas_and_ghost()
        elif choice == "6":
            export_cleaned_from_modified()
        elif choice == "0":
            break
        else:
            print("Invalid choice. Please try again.")


if __name__ == "__main__":
    try:
        main_menu()
    except KeyboardInterrupt:
        print("\nInterrupted.")
    finally:
        if driver:
            driver.quit()
        print("Exiting.")
