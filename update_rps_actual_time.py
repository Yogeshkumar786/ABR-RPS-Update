import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright
from datetime import datetime
import time
import logging
import os
from gspread.exceptions import APIError

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def retry_gspread_request(func, *args, retries=5, delay=2, **kwargs):
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            if "503" in str(e):
                logging.warning(f"[{attempt + 1}/{retries}] 503 APIError. Retrying in {delay} seconds...")
                time.sleep(delay)
                delay *= 2
            else:
                raise
    raise Exception("Max retries exceeded for gspread request.")

def get_sheets():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(creds)

    sheet1 = retry_gspread_request(
        lambda: retry_gspread_request(
            client.open_by_url("https://docs.google.com/spreadsheets/d/17JhXtQfzc6XVmOEXNyNEloMa-mTWgJN4AVAbbJSLoHI/edit?usp=sharing").worksheet,
            "Live_Tracking"
        )
    )
    sheet2 = retry_gspread_request(
        lambda: retry_gspread_request(
            client.open_by_url("https://docs.google.com/spreadsheets/d/1xUjnEup_k6jGleTsJZQwSYkapbalPUM-DMzNtQujDo0/edit?usp=sharing").worksheet,
            "Closed_RPS"
        )
    )
    return sheet1, sheet2

def get_column_index(headers, target_name):
    for idx, col in enumerate(headers):
        if col.strip().lower() == target_name.strip().lower():
            return idx + 1
    raise ValueError(f"Column '{target_name}' not found. Headers were: {headers}")

def fetch_times_with_playwright(rps_no: str):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        start_time, reaching_time, route = None, None, None

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
            reaching_xpath = '/html/body/form/div[5]/div/div/div/div/div/div/div[4]/div/table/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[4]'
            start_xpath = '/html/body/form/div[5]/div/div/div/div/div/div/div[4]/div/table/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[3]'
            route_xpath = '/html/body/form/div[5]/div/div/div/div/div/div/div[4]/div/table/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[6]'
            reaching_time = page.locator(f'xpath={reaching_xpath}').text_content(timeout=5000)
            start_time = page.locator(f'xpath={start_xpath}').text_content(timeout=5000)
            route = page.locator(f'xpath={route_xpath}').text_content(timeout=5000)
        except Exception as e:
            logging.error(f"Error fetching times for RPS {rps_no}: {e}")
        finally:
            browser.close()

        return start_time, reaching_time, route

def update_and_migrate_batch():
    sheet1, sheet2 = get_sheets()
    headers1 = sheet1.row_values(1)
    headers2 = sheet2.row_values(1)

    col_rps_1 = get_column_index(headers1, "RPS No")
    col_vehicle_1 = get_column_index(headers1, "Vehicle Number")

    col_start_2 = get_column_index(headers2, "Route_Start_Date_Time")
    col_vehicle_2 = get_column_index(headers2, "Vehicle_Number")
    col_rps_2 = get_column_index(headers2, "RPS No")
    col_reach_2 = get_column_index(headers2, "Route_Reaching_Date_Time")
    col_route_2 = get_column_index(headers2, "Route")

    records = sheet1.get_all_records()
    sheet2_records = sheet2.get_all_records()
    existing_rps_set = {str(row["RPS No"]).strip(): idx+2 for idx, row in enumerate(sheet2_records)}

    to_delete_indices = []

    for i, row in enumerate(records):
        rps_no = str(row.get("RPS No", "")).strip()
        col_8_value = row.get(headers1[7], "").strip().lower()

        if col_8_value != "closed" or not rps_no:
            continue

        vehicle_no = str(row.get("Vehicle Number", "")).strip()
        logging.info(f"Fetching times for RPS No: {rps_no}")
        start_time, reaching_time, route = fetch_times_with_playwright(rps_no)
        route = route.replace(" ", "") if route else ""

        if reaching_time and reaching_time.strip():
            if rps_no in existing_rps_set:
                row_idx = existing_rps_set[rps_no]
                sheet2.update_cell(row_idx, col_start_2, start_time)
                sheet2.update_cell(row_idx, col_vehicle_2, vehicle_no)
                sheet2.update_cell(row_idx, col_rps_2, rps_no)
                sheet2.update_cell(row_idx, col_reach_2, reaching_time)
                sheet2.update_cell(row_idx, col_route_2, route)
                logging.info(f"Updated RPS {rps_no} in RPS_Closed")
            else:
                new_row = [''] * max(col_start_2, col_vehicle_2, col_rps_2, col_reach_2, col_route_2)
                new_row[col_start_2 - 1] = start_time
                new_row[col_vehicle_2 - 1] = vehicle_no
                new_row[col_rps_2 - 1] = rps_no
                new_row[col_reach_2 - 1] = reaching_time
                new_row[col_route_2 - 1] = route
                sheet2.append_row(new_row)
                logging.info(f"Inserted new RPS {rps_no} in RPS_Closed")
            to_delete_indices.append(i + 2)
        else:
            logging.info(f"No reaching_time for RPS {rps_no}. Skipping.")

        time.sleep(3)

    for idx in sorted(to_delete_indices, reverse=True):
        try:
            retry_gspread_request(sheet1.delete_rows, idx)
            logging.info(f"Deleted row {idx} from Sheet3")
        except Exception as e:
            logging.error(f"Failed to delete row {idx} after retries: {e}")
        logging.info(f"Deleted row {idx} from Sheet3")

if __name__ == "__main__":
    update_and_migrate_batch()

