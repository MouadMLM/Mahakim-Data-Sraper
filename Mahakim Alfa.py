# Mahakim Beta Enhanced Progress.py
# Scraper for Mahakim.ma with robust table detection and parsing by Mouadev
import time
import random
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    WebDriverException,
    TimeoutException,
    NoSuchElementException
)
import csv
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

# ===== CONFIG =====
TARGET_URL = "https://www.mahakim.ma/#/suivi/rapport-police-judiciaire"
START_NUM = 1
END_NUM = 3000
YEAR = "2025"
#OUTPUT_CSV = "C:/Users/AlienM/Downloads/results.csv"
OUTPUT_XLSX = "C:/Users/AlienM/Downloads/results.xlsx"
PROGRESS_FILE = "C:/Users/AlienM/Downloads/progress.txt"
HEADLESS = False
MIN_DELAY = 0.8
MAX_DELAY = 2.2
RETRIES = 3
# ==================

# ===== INIT DRIVER =====
def init_driver():
    options = webdriver.ChromeOptions()
    if HEADLESS:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.page_load_strategy = 'eager'
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    driver.set_page_load_timeout(90)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

# ===== PROGRESS =====
def read_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            return int(open(PROGRESS_FILE).read().strip())
        except:
            return None
    return None

def write_progress(n):
    with open(PROGRESS_FILE, "w") as f:
        f.write(str(n))

# ===== PROPER CSV APPEND (Optimized) =====
def append_to_csv_properly(rows):
    if not rows:
        return

    columns = [
        "رقم المحضر بالمحكمة",
        "الإجراء", 
        "نوع المحضر",
        "موضوع المحضر",
        "رقم الملف الجنحي",
        "مزيد من المعلومات",
        "الرقم المستعلم",
        "السنة المستعلم بها"
    ]

    mapped_data = []
    for row in rows:
        mapped_row = {
            "رقم المحضر بالمحكمة": row.get("case_number", ""),
            "الإجراء": row.get("action", ""),
            "نوع المحضر": row.get("type", ""),
            "موضوع المحضر": row.get("subject", ""),
            "رقم الملف الجنحي": row.get("file_number", ""),
            "مزيد من المعلومات": row.get("more_info", ""),
            "الرقم المستعلم": row.get("queried_numero", ""),
            "السنة المستعلم بها": row.get("queried_annee", "")
        }
        mapped_data.append(mapped_row)

    df = pd.DataFrame(mapped_data, columns=columns)

    if os.path.exists(OUTPUT_XLSX):
        wb = load_workbook(OUTPUT_XLSX)
        ws = wb.active
        for r in dataframe_to_rows(df, index=False, header=False):
            ws.append(r)
        wb.save(OUTPUT_XLSX)
    else:
        with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Results")

    print(f"💾 Appended {len(rows)} new rows to Excel ({OUTPUT_XLSX})")

# ===== ROBUST TABLE DETECTION =====
def robust_table_detection(driver, case_number):
    """
    Use multiple methods to detect and parse table data
    """
    print(f"\n Checking for data in case {case_number}...")
    
    # Method 1: Check for no results message
    try:
        no_results = driver.find_elements(By.XPATH, "//p[contains(text(), 'لا توجد أية نتيجة للبحث')]")
        if no_results:
            for element in no_results:
                if element.is_displayed():
                    print(" No results found 'message detected' ")
                    return "no_results", []
    except:
        pass

    # Method 2: Look for the table by ID
    try:
        table = driver.find_element(By.ID, "pr_id_16-table")
        if table.is_displayed():
            print("  ✅ Table found by ID")
            data = parse_table_by_element(table, case_number)
            if data:
                return "has_data", data
    except:
        pass

    # Method 3: Look for any table with data patterns
    try:
        all_tables = driver.find_elements(By.TAG_NAME, "table")
        for table in all_tables:
            if table.is_displayed():
                data = parse_table_by_element(table, case_number)
                if data:
                    print(f"  ✅ Data found in alternative table")
                    return "has_data", data
    except:
        pass

    # Method 4: Look for data patterns in the entire page
    try:
        # Look for case number patterns like "5248/3205/2025"
        case_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '/') and contains(text(), '2025')]")
        if case_elements:
            print(f"  🔍 Found {len(case_elements)} elements with case number pattern")
            # If we found case numbers but no table, there might be a display issue
            return "possible_data", []
    except:
        pass

    # Method 5: Check for loading indicators or errors
    try:
        loading = driver.find_elements(By.XPATH, "//*[contains(text(), 'جاري') or contains(text(), 'تحميل') or contains(text(), 'loading')]")
        if loading:
            print(" Loading indicator found")
            return "loading", []
    except:
        pass

    print("  ❓ No clear data status detected")
    return "unknown", []

# ===== PARSE TABLE BY ELEMENT =====
def parse_table_by_element(table_element, case_number):
    """
    Parse table data from a table element
    """
    try:
        data_rows = []
        
        # Get all rows
        rows = table_element.find_elements(By.TAG_NAME, "tr")
        print(f"    Found {len(rows)} rows in table")
        
        for i, row in enumerate(rows):
            try:
                if not row.is_displayed():
                    continue
                    
                cells = row.find_elements(By.TAG_NAME, "td")
                
                # We need exactly 6 cells for a data row
                if len(cells) == 6:
                    # Skip if any cell has colspan (no-data row)
                    if any(cell.get_attribute("colspan") for cell in cells):
                        continue
                        
                    # Extract data
                    row_data = {
                        "case_number": cells[0].text.strip(),
                        "action": cells[1].text.strip(),
                        "type": cells[2].text.strip(),
                        "subject": cells[3].text.strip(),
                        "file_number": cells[4].text.strip(),
                        "more_info": cells[5].text.strip(),
                        "queried_numero": case_number,
                        "queried_annee": YEAR
                    }
                    
                    # Validate it's real data (has case number)
                    if row_data["case_number"] and '/' in row_data["case_number"]:
                        data_rows.append(row_data)
                        print(f"    ✅ Row {i}: {row_data['case_number']}")
                        
            except Exception as e:
                continue
                
        return data_rows
        
    except Exception as e:
        print(f"    ❌ Error parsing table: {e}")
        return []

# ===== WAIT FOR RESULTS =====
def wait_for_results(driver, timeout=10):
    """
    Wait for search results to load
    """
    try:
        wait = WebDriverWait(driver, timeout)
        
        # Wait for either results or no results to appear
        wait.until(lambda d: 
            d.find_elements(By.XPATH, "//p[contains(text(), 'لا توجد أية نتيجة للبحث')]") or
            d.find_elements(By.ID, "pr_id_16-table") or
            d.find_elements(By.XPATH, "//*[contains(text(), '/') and contains(text(), '2025')]")
        )
        return True
    except TimeoutException:
        return False

# ===== SAFE FIND =====
def safe_find(driver, by, value, timeout=10):
    wait = WebDriverWait(driver, timeout)
    return wait.until(EC.element_to_be_clickable((by, value)))

# ===== JS CLICK =====
def js_click(driver, element):
    driver.execute_script("arguments[0].click();", element)

# ===== PRECISE DROPDOWN SELECTION =====
def select_dropdown_by_placeholder(driver, placeholder_text, option_to_select, step_number):
    print(f"\n=== STEP {step_number}: Selecting {placeholder_text} ===")
    
    wait = WebDriverWait(driver, 15)
    dropdown_xpath = f"//span[contains(@class, 'p-dropdown-label') and contains(@class, 'p-placeholder') and contains(text(), '{placeholder_text}')]/ancestor::div[contains(@class, 'p-dropdown')]"
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
    
    print(f"Found dropdown: {placeholder_text}")
    
    js_click(driver, dropdown)
    time.sleep(1)
    
    wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-dropdown-panel')]")))
    time.sleep(0.5)
    
    option_xpath = f"//li[contains(@class, 'p-dropdown-item')]//span[contains(text(), '{option_to_select}')]"
    
    try:
        option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        js_click(driver, option)
        print(f"✓ Selected: {option_to_select}")
        time.sleep(2)
        return True
    except TimeoutException:
        options = driver.find_elements(By.XPATH, "//li[contains(@class, 'p-dropdown-item')]")
        available_options = [opt.text for opt in options if opt.text.strip()]
        print(f"✗ Option '{option_to_select}' not found. Available: {available_options}")
        dropdown.click()
        time.sleep(0.5)
        return False

# ===== CLICK CHECKBOX =====
def click_checkbox(driver):
    print("\n=== STEP 2: Clicking Checkbox ===")
    wait = WebDriverWait(driver, 10)
    checkbox_xpath = "//div[contains(@class, 'p-checkbox-box')]"
    checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
    js_click(driver, checkbox)
    print("✓ Checkbox clicked")
    time.sleep(2)
    return True

# ===== FILL CASE NUMBER AND YEAR =====
def fill_case_details(driver, case_number, year):
    print(f"🔎 Searching case {case_number}/{year}...")
    
    wait = WebDriverWait(driver, 10)
    container_xpath = "//div[contains(@class, 'three-inputs')]"
    container = wait.until(EC.presence_of_element_located((By.XPATH, container_xpath)))
    
    numero_input = container.find_element(By.XPATH, ".//input[@formcontrolname='numero' and contains(@class, 'right')]")
    annee_input = container.find_element(By.XPATH, ".//input[@formcontrolname='annee' and contains(@class, 'left')]")
    
    numero_input.clear()
    numero_input.send_keys(str(case_number))
    
    annee_input.clear()
    annee_input.send_keys(str(year))
    
    annee_input.send_keys(Keys.ENTER)
    
    # Wait for results with better detection
    print("⏳...")
    if wait_for_results(driver, 10):
        print("  ✅ Results loaded")
    else:
        print("  ⚠️  Results timeout, continuing...")
    
    time.sleep(2)  # Additional wait
    return True

# ===== RUN SCRAPER =====
def run_scraper():
    start_resume = read_progress()
    if start_resume:
        start_n = start_resume + 1
        print(f"🔄 Resuming from number {start_n} (last progress: {start_resume})")
    else:
        start_n = START_NUM
        print(f" Starting from number {start_n}")

    driver = init_driver()

    try:
        # Open page
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                print(f"Loading page... (attempt {attempt + 1})")
                driver.get(TARGET_URL)
                time.sleep(5)
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                print("Page loaded successfully")
                break
            except TimeoutException:
                print(f"Timeout loading page, retry {attempt+1}/{max_attempts}")
                if attempt == max_attempts - 1:
                    raise

        # ===== SETUP DROPDOWNS =====
        select_dropdown_by_placeholder(driver, "اختيار محكمة الاستئناف", "محكمة الاستئناف بمراكش", 1)
        click_checkbox(driver)
        select_dropdown_by_placeholder(driver, "اختيار المحكمة الإبتدائية", "المحكمة الابتدائية بمراكش", 3)

        # ===== SELECT POLICE UNIT AND STATION =====
        print("\n=== STEP 4: Selecting Police Unit ===")
        wait = WebDriverWait(driver, 10)
        blank_dropdowns = driver.find_elements(By.XPATH, "//span[contains(@class, 'p-dropdown-label') and contains(@class, 'p-placeholder') and contains(text(), '---')]/ancestor::div[contains(@class, 'p-dropdown')]")
        
        if len(blank_dropdowns) >= 1:
            first_blank_dropdown = blank_dropdowns[0]
            js_click(driver, first_blank_dropdown)
            time.sleep(1)
            
            police_options = ["الدرك الملكي", "الشرطة القضائية", "الامن الوطني"]
            options = driver.find_elements(By.XPATH, "//li[contains(@class, 'p-dropdown-item')]")
            available_options = [opt.text for opt in options if opt.text.strip()]
            print(f"Available police units: {available_options}")
            
            selected = False
            for police_opt in police_options:
                for option in options:
                    if police_opt in option.text:
                        js_click(driver, option)
                        print(f"✓ Selected police unit: {police_opt}")
                        selected = True
                        time.sleep(2)
                        break
                if selected:
                    break

        # ===== STEP 5: Select Police Station =====
        print("\n=== STEP 5: Selecting Police Station ===")
        
        if len(blank_dropdowns) >= 2:
            second_blank_dropdown = blank_dropdowns[1]
            js_click(driver, second_blank_dropdown)
            time.sleep(1)
            
            station_options = ["قائد مركز الدرك الملكي بايت اورير", "مركز الدرك الملكي بايت اورير"]
            options = driver.find_elements(By.XPATH, "//li[contains(@class, 'p-dropdown-item')]")
            available_options = [opt.text for opt in options if opt.text.strip()]
            print(f"Available police stations: {available_options}")
            
            selected = False
            for station_opt in station_options:
                for option in options:
                    if station_opt in option.text:
                        js_click(driver, option)
                        print(f"✓ Selected police station: {station_opt}")
                        selected = True
                        time.sleep(2)
                        break
                if selected:
                    break

        # ===== START SCRAPING =====
        print("\n" + "="*50)
        print("STARTING SCRAPING")
        print("="*50)
        
        for n in range(start_n, END_NUM + 1):
            write_progress(n)
            attempt = 0
            success = False
            
            while not success and attempt < RETRIES:
                attempt += 1
                try:
                    fill_case_details(driver, n, YEAR)
                    
                    # Use robust detection
                    status, rows_data = robust_table_detection(driver, n)
                    
                    if status == "no_results":
                        print(f"🚫 [NO RESULTS] {n} - Skipping")
                        success = True
                        
                    elif status == "has_data" and rows_data:
                        append_to_csv_properly(rows_data)
                        print(f"✅ [FOUND] {n} -> {len(rows_data)} rows")
                        success = True
                        
                    elif status == "possible_data":
                        print(f"🔍 [POSSIBLE DATA] {n} - Data patterns found but couldn't parse")
                        # Try one more time with different approach
                        if attempt == RETRIES:
                            success = True
                            
                    elif status == "loading":
                        print(f"⏳ [STILL LOADING] {n} - Retrying...")
                        time.sleep(3)  # Wait longer and retry
                        
                    else:
                        print(f"❓ [UNKNOWN: {status}] {n} - Retrying...")
                        if attempt == RETRIES:
                            success = True

                    delay = MIN_DELAY + random.random() * (MAX_DELAY - MIN_DELAY)
                    time.sleep(delay)
                    
                except Exception as e:
                    print(f"Error for {n}: {e}, retrying...")
                    time.sleep(2)
                    
                    if attempt == RETRIES - 1:
                        print("Refreshing page...")
                        driver.refresh()
                        time.sleep(5)
                        # Re-setup
                        select_dropdown_by_placeholder(driver, "اختيار محكمة الاستئناف", "محكمة الاستئناف بمراكش", "1-re")
                        click_checkbox(driver)
                        select_dropdown_by_placeholder(driver, "اختيار المحكمة الإبتدائية", "المحكمة الابتدائية بمراكش", "3-re")

        print(f"\n✅ Scraping completed successfully!")

    except Exception as e:
        print(f"❌ Fatal error: {e}")
        print("💾 Data saved so far is preserved in CSV file")
        raise

    finally:
        driver.quit()
        print(f"📊 Results saved to {OUTPUT_CSV}")
        print(f"📝 Progress saved to {PROGRESS_FILE}")

# ===== MAIN =====
if __name__ == "__main__":

    run_scraper()
