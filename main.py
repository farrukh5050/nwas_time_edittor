import os
import time
import calendar
from datetime import date
import pandas as pd
from io import StringIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import requests
from dotenv import load_dotenv
import load_files

# Load credentials from .env file
load_dotenv()

NWAS_USERNAME = os.getenv("NWAS_USERNAME")
NWAS_PASSWORD = os.getenv("NWAS_PASSWORD")
if not NWAS_USERNAME or not NWAS_PASSWORD:
    raise RuntimeError(
        "Missing credentials: set NWAS_USERNAME and NWAS_PASSWORD in your .env file."
    )

# Store the browser driver so it can be reused
driver = None


def load_and_clean_file():
    cleaned_df = load_files.select_excel_file()
    if cleaned_df is not None:
        cleaned_df.to_excel("cleaned_ghost_data.xlsx", index=False)
        print("Cleaned file saved as 'cleaned_ghost_data.xlsx'.")


def open_chrome_and_login():
    global driver
    if driver:
        print("Chrome is already open.")
        return

    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("https://ptsed.nwas.nhs.uk/")
    driver.find_element(By.ID, "txtUsername").send_keys(str(NWAS_USERNAME))
    driver.find_element(By.ID, "txtPassword").send_keys(str(NWAS_PASSWORD))
    driver.find_element(By.ID, "cmdSubmit").click()
    time.sleep(3)
    driver.get("https://ptsed.nwas.nhs.uk/frmQualityEntry.aspx")
    input("Load dates on the page and press Enter to continue...")


def get_selected_date(driver):
    return driver.find_element(By.ID, "txtPlanDate").get_attribute("value")


def get_ids_from_table(driver, excel_path="cleaned_ghost_data.xlsx"):
    ghost_df = pd.read_excel(excel_path)
    ghost_df["Your Reference 1"] = (
        ghost_df["Your Reference 1"].astype(str).str.replace(" ", "")
    )
    table_html = driver.find_element(
        By.XPATH, '//*[@id="tblResults"]'
    ).get_attribute("outerHTML")
    table_df = pd.read_html(StringIO(table_html))[0]
    table_df["ID"] = table_df["ID"].astype(str)
    merged = table_df.merge(
        ghost_df, how="left", left_on="ID", right_on="Your Reference 1"
    )
    merged = merged.dropna(subset=["Vehicle Arrived at Time", "Completed at Time"])
    return merged[["ID", "Vehicle Arrived at Time", "Completed at Time"]]


def get_auth_cookie(driver):
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
            response = session.post(url, json={"Data": payload}, headers=headers)
            print(f"Updated ID {row['ID']} - Status: {response.status_code}")
            if response.status_code != 200:
                print("Error:", response.text)
        except Exception as e:
            print(f"Skipped ID {row['ID']} due to error: {e}")
            continue


def update_times():
    global driver
    if not driver:
        print("❌ Please run option 2 first to open and log into Chrome.")
        return
    try:
        selected_date = get_selected_date(driver)
        print(f"Selected date from portal: {selected_date}")
        matched_df = get_ids_from_table(driver)
        print("Matched entries:")
        print(matched_df)

        session = get_auth_cookie(driver)
        update_ghost_times(session, matched_df, date_str=selected_date)
        print("🎉 All times updated.")
    except Exception as e:
        print(f"❌ Error during update: {e}")


def dismiss_message_box():
    """Click OK on the portal message box if it is showing. Returns True if dismissed."""
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


def run_full_automation(call_signs=None, year=2026, month=5, load_wait=10):
    """Loop over each call sign and every day of the month, search, then update times."""
    global driver
    if not driver:
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
                driver.find_element(
                    By.XPATH, "/html/body/form/div[5]/div[3]/button[1]"
                ).click()

                # Wait until the jobs table is visible (up to load_wait seconds)
                if wait_for_results(load_wait):
                    # Update times for the loaded jobs
                    update_times()
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
        print("0. Exit")

        choice = input("Select an option (0, 1, 2, 3, or 4): ").strip()

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
        elif choice == "0":
            print("Exiting.")
            break
        else:
            print("Invalid choice. Please try again.")


if __name__ == "__main__":
    main_menu()
