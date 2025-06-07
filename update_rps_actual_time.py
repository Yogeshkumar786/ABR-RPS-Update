import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright
from datetime import datetime
import time
import logging
import os

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Load Google Sheets credentials from external file
def get_sheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1NKaBL6fze5tKqu_TfKj3OxN1VkSbQCjAK9zJrxnf9Ac/edit?usp=sharing")
    return spreadsheet.worksheet("Sheet1")

# Scrape actual_time for a given RPS number
def fetch_rps_actual_time(rps_no: str) -> str:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        actual_time = None

        try:
            page.goto("http://smart.dsmsoft.com/FMSSmartApp/Safex_RPS_Reports/RPS_Reports.aspx?usergroup=NRM.101", wait_until="domcontentloaded", timeout=120000)

            # Perform required clicks and selections (fragile XPaths)
            page.locator('xpath=/html/body/form/div[5]/div/div/div/div/div/div/div[3]/div/div[4]/div[2]').click()
            page.wait_for_timeout(500)
            page.locator('xpath=/html/body/form/div[5]/div/div/div/div/div/div/div[3]/div/div[4]/div[3]/div[2]/ul/li[1]/input').click()
            page.wait_for_timeout(500)
            page.locator('xpath=/html/body/form/div[5]/div/div/div/div/div/div/div[3]/div/div[1]/div[2]/input').click()
            page.wait_for_timeout(500)
            page.locator('//div[contains(@class,"xdsoft_datepicker")]//button[contains(@class,"xdsoft_prev")]').nth(0).click()
            page.wait_for_timeout(1000)

            today = datetime.now()
            day_xpath = f'//td[@data-date="{today.day}" and contains(@class, "xdsoft_date") and not(contains(@class, "xdsoft_disabled"))]'
            page.locator(day_xpath).nth(0).click()
            page.wait_for_timeout(500)

            page.locator('xpath=/html/body/form/div[5]/div/div/div/div/div/div/div[3]/div/div[5]/div/button').click()
            page.wait_for_timeout(30000)

            rps_input_xpath = '/html/body/form/div[5]/div/div/div/div/div/div/div[4]/div/table/div/div[4]/div/div/div[3]/div[3]/div/div/div/div[1]/input'
            page.locator(f'xpath={rps_input_xpath}').click()
            page.wait_for_timeout(500)
            page.fill(f'xpath={rps_input_xpath}', rps_no)
            page.wait_for_timeout(2000)

            time_xpath = '/html/body/form/div[5]/div/div/div/div/div/div/div[4]/div/table/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[4]'
            actual_time = page.locator(f'xpath={time_xpath}').text_content(timeout=5000)
        except Exception as e:
            logging.error(f"Error fetching actual_time for RPS {rps_no}: {e}")
        finally:
            browser.close()

        return actual_time

# Update Google Sheet with fetched actual_times
def update_sheet_with_actual_times():
    sheet = get_sheet()
    records = sheet.get_all_records()

    for i, row in enumerate(records, start=2):
        if not row.get("Actual_Time"):
            rps_no = str(row.get("RPS No"))
            logging.info(f"Fetching Actual_time for {rps_no}")
            actual_time = fetch_rps_actual_time(rps_no)
            if actual_time:
                logging.info(f"Found: {actual_time}")
                sheet.update_cell(i, 2, actual_time)
                time.sleep(3)
            else:
                logging.warning(f"Skipping RPS {rps_no}: No data found")

if __name__ == "__main__":
    update_sheet_with_actual_times()
