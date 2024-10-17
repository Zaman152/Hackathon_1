import requests
import csv
from io import StringIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import chromedriver_autoinstaller
import concurrent.futures
from google.oauth2 import service_account
from googleapiclient.discovery import build
import json


SHEET_ID = '1lFeMu62iHRMmYiyRLKLluveldwtsh16iBMpIpLTXU4w'
SHEET_NAME = 'Sheet1'
URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}'


WAIT_TIMEOUT = 30
FORM_URL = "https://tally.so/r/waDMG2"
MAX_WORKERS = 5
STATUS_COLUMN = 'C'

class SheetUpdater:
     def __init__(self):
        try:
            with open('credentials.json', 'r') as creds_file:
                creds_dict = json.load(creds_file)
            
            credentials = service_account.Credentials.from_service_account_info(
                creds_dict,
                scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            self.service = build('sheets', 'v4', credentials=credentials)
            print("Successfully connected to Google Sheets API")
        except FileNotFoundError:
            print("Error: credentials.json file not found in the current directory.")
            raise
        except json.JSONDecodeError:
            print("Error: Invalid JSON in credentials.json file.")
            raise
        except Exception as e:
            print(f"Error initializing SheetUpdater: {str(e)}")
            raise

     def update_status(self, row_num, status="Done"):
        """Update status in Google Sheets"""
        try:
            range_name = f'{SHEET_NAME}!{STATUS_COLUMN}{row_num}'
            body = {
                'values': [[status]]
            }
            result = self.service.spreadsheets().values().update(
                spreadsheetId=SHEET_ID,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            print(f"Successfully updated status for row {row_num}. Update result: {result}")
            return True
        except Exception as e:
            print(f"Error updating status for row {row_num}: {str(e)}")
            return False


sheet_updater = SheetUpdater()


def update_sheet_status(row_num):
    return sheet_updater.update_status(row_num)
def get_sheet_data(start_row):
    try:
        response = requests.get(URL, timeout=10)
        response.raise_for_status()
        data = list(csv.reader(StringIO(response.text)))
        return data[start_row-1:]
    except requests.exceptions.RequestException as e:
        print(f"Error fetching sheet data: {str(e)}")
        return []


def update_sheet_status(row_num, status="Done"):
       try:
           range_name = f'{SHEET_NAME}!C{row_num}'
           body = {
               'values': [[status]]
           }
           result = sheet_updater.service.spreadsheets().values().update(
               spreadsheetId=SHEET_ID,
               range=range_name,
               valueInputOption='RAW',
               body=body
           ).execute()
           print(f"Updated status for row {row_num}")
           return True
       except Exception as e:
           print(f"Error updating status for row {row_num}: {str(e)}")
           return False
   


def setup_driver():
    try:
        # Auto-install matching ChromeDriver
        chromedriver_autoinstaller.install()
        
        chrome_options = Options()
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        
        # Initialize Chrome driver without explicitly specifying the service
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        print(f"Error setting up Chrome driver: {str(e)}")
        raise

def click_button(driver, wait):
    button_locators = [
        (By.XPATH, "//button[contains(text(), 'done')]"),
        (By.XPATH, "//button[contains(text(), 'Done')]"),
        (By.XPATH, "//button[contains(text(), 'Submit')]"),
        (By.XPATH, "//button[@type='submit']"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "form button:last-child")
    ]

    for locator in button_locators:
        try:
            
            button = driver.find_element(*locator)
            
            
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            
            
            driver.execute_script("arguments[0].click();", button)
            print(f"Button clicked using JavaScript for locator: {locator}")
            
            
            time.sleep(1)
            
            
            if driver.current_url != FORM_URL or "success" in driver.page_source.lower():
                return True
            
            
            button.click()
            print(f"Button clicked using regular click for locator: {locator}")
            
            
            time.sleep(1)
            
            if driver.current_url != FORM_URL or "success" in driver.page_source.lower():
                return True
            
        except Exception as e:
            print(f"Failed to click button with locator {locator}: {str(e)}")
            continue

    print("Failed to click the button using all available locators")
    return False
def fill_form(name, email, row):
    driver = None
    try:
        driver = setup_driver()
        driver.set_page_load_timeout(WAIT_TIMEOUT)
        
        print(f"\nProcessing row {row}")
        print(f"Processing: {name}, {email}")
        
        driver.get(FORM_URL)
        wait = WebDriverWait(driver, WAIT_TIMEOUT)
        time.sleep(3)
        
        js_fill_script = f"""
        function fillForm() {{
            var nameField = document.querySelector('input[name="Name"]') || 
                            document.querySelector('input[placeholder*="Name" i]') ||
                            document.querySelector('input[aria-label*="Name" i]') ||
                            document.querySelector('label:has-text("Name")').nextElementSibling.querySelector('input');
            var emailField = document.querySelector('input[name="Email"]') || 
                             document.querySelector('input[type="email"]') ||
                             document.querySelector('input[placeholder*="Email" i]');
            
            if (nameField) {{
                nameField.value = "{name}";
                nameField.dispatchEvent(new Event('input', {{ bubbles: true }}));
                nameField.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
            
            if (emailField) {{
                emailField.value = "{email}";
                emailField.dispatchEvent(new Event('input', {{ bubbles: true }}));
                emailField.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
            
            return [nameField, emailField];
        }}
        return fillForm();
        """
        
        filled_fields = driver.execute_script(js_fill_script)
        
        if not all(filled_fields):
            print("Failed to fill form fields")
            return False
        
        print("Form fields filled successfully")
        
        if click_button(driver, wait):
            print(f"Form submitted for row {row}")
            update_success = update_sheet_status(row)
            if update_success:
                print(f"Successfully updated Google Sheet for row {row}")
            else:
                print(f"Failed to update Google Sheet for row {row}")
            return update_success
            
        return False
            
    except Exception as e:
        print(f"\nError processing row {row}")
        print(f"Error details: {str(e)}")
        return False
        
    finally:
        if driver:
            driver.quit()
def process_entry(entry):
    row, name, email = entry
    success = fill_form(name, email, row)
    return {
        'row': row,
        'name': name,
        'email': email,
        'success': success
    }


def process_batch(batch, start_row):
    entries = [(start_row + i, row[0].strip(), row[1].strip()) for i, row in enumerate(batch) if len(row) >= 2 and row[0].strip() and row[1].strip()]
    
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_entry = {executor.submit(process_entry, entry): entry for entry in entries}
        for future in concurrent.futures.as_completed(future_to_entry):
            entry = future_to_entry[future]
            try:
                result = future.result()
                results.append(result)
            except Exception as exc:
                print(f'Entry {entry} generated an exception: {exc}')
    
    print("\nProcessing Summary:")
    successful = sum(1 for r in results if r['success'])
    print(f"Total processed: {len(results)}")
    print(f"Successful: {successful}")
    print(f"Failed: {len(results) - successful}")

def main():
    try:
        start_row = 2  
        
        data = get_sheet_data(start_row)
        if not data:
            print("No data found or error fetching data.")
            return
            
        print(f"Found {len(data)} rows to process.")
        print("Starting form submission process...")
        
        process_batch(data, start_row)
            
        print("\nProcess completed!")
        
    except KeyboardInterrupt:
        print("\nProcess interrupted by user.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()