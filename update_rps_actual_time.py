import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright
from datetime import datetime
import time

# Setup Google Sheets access
def get_sheet(sheet_name="vehicles"):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    return client.worksheet("Sheet4")

# Function to scrape actual_time for 1 RPS number
def fetch_rps_actual_time(rps_no: str) -> str:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        try:
            page.goto("http://smart.dsmsoft.com/FMSSmartApp/Safex_RPS_Reports/RPS_Reports.aspx?usergroup=NRM.101", wait_until="domcontentloaded", timeout=120000)
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
            page.wait_for_timeout(3000)
            rps_input_xpath = '/html/body/form/div[5]/div/div/div/div/div/div/div[4]/div/table/div/div[4]/div/div/div[3]/div[3]/div/div/div/div[1]/input'
            page.locator(f'xpath={rps_input_xpath}').click()
            page.wait_for_timeout(500)
            page.fill(f'xpath={rps_input_xpath}', rps_no)
            page.wait_for_timeout(2000)
            time_xpath = '/html/body/form/div[5]/div/div/div/div/div/div/div[4]/div/table/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[4]'
            actual_time = page.locator(f'xpath={time_xpath}').text_content(timeout=5000)
        except Exception as e:
            print(f"❌ Error for {rps_no}: {e}")
            actual_time = None
        browser.close()
        return actual_time

# Main function
def update_sheet_with_actual_times():
    sheet = get_sheet()
    records = sheet.get_all_records()
    for i, row in enumerate(records, start=2):
        if not row["Actual_Time"]:
            rps_no = str(row["RPS No"])
            print(f"🔍 Fetching Actual_time for {rps_no}")
            actual_time = fetch_rps_actual_time(rps_no)
            if actual_time:
                print(f"✅ Found: {actual_time}")
                sheet.update_cell(i, 2, actual_time)
                time.sleep(3)
            else:
                print(f"⚠️ Skipping RPS {rps_no}: No data found")

if __name__ == "__main__":
    update_sheet_with_actual_times()
