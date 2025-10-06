# Mahakim Beta Enhanced Progress.py
# Scraper for Mahakim.ma with robust table detection and parsing by Mouadev
#Preconfigured for Marrakech > Ait Ourir > GR Reports 
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

TARGET_URL = "https://www.mahakim.ma/#/suivi/rapport-police-judiciaire"
START_NUM = 1
END_NUM = 3000
YEAR = "2025"
OUTPUT_XLSX = "C:/Users/AlienM/Downloads/results.xlsx"
PROGRESS_FILE = "C:/Users/AlienM/Downloads/progress.txt"
HEADLESS = False
MIN_DELAY = 0.8
MAX_DELAY = 2.2
RETRIES = 3

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

def read_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                lines = f.readlines()
                for line in reversed(lines):
                    if line.startswith("PROGRESS:"):
                        return int(line.replace("PROGRESS:", "").strip())
        except:
            return None
    return None

def write_progress(n, info=""):
    with open(PROGRESS_FILE, "a", encoding="utf-8") as f:
        if info:
            f.write(info + "\n")
        f.write(f"PROGRESS: {n}\n")

def append_to_excel(rows):
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

def js_click(driver, element):
    driver.execute_script("arguments[0].click();", element)

def safe_find(driver, by, value, timeout=10):
    wait = WebDriverWait(driver, timeout)
    return wait.until(EC.element_to_be_clickable((by, value)))

def wait_for_results(driver, timeout=10):
    try:
        wait = WebDriverWait(driver, timeout)
        wait.until(lambda d: 
            d.find_elements(By.XPATH, "//p[contains(text(), 'لا توجد أية نتيجة للبحث')]") or
            d.find_elements(By.ID, "pr_id_16-table") or
            d.find_elements(By.XPATH, "//*[contains(text(), '/') and contains(text(), '2025')]")
        )
        return True
    except TimeoutException:
        return False

def select_dropdown(driver, placeholder_text, option_to_select, step_number):
    info_lines = [f"=== STEP {step_number}: Selecting {placeholder_text} ==="]
    wait = WebDriverWait(driver, 15)
    dropdown_xpath = f"//span[contains(@class, 'p-dropdown-label') and contains(@class, 'p-placeholder') and contains(text(), '{placeholder_text}')]/ancestor::div[contains(@class, 'p-dropdown')]"
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
    info_lines.append(f"Found dropdown: {placeholder_text}")
    js_click(driver, dropdown)
    time.sleep(1)
    wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-dropdown-panel')]")))
    time.sleep(0.5)
    option_xpath = f"//li[contains(@class, 'p-dropdown-item')]//span[contains(text(), '{option_to_select}')]"
    try:
        option = wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        js_click(driver, option)
        info_lines.append(f"✓ Selected: {option_to_select}")
        time.sleep(1)
    except TimeoutException:
        options = driver.find_elements(By.XPATH, "//li[contains(@class, 'p-dropdown-item')]")
        available_options = [opt.text for opt in options if opt.text.strip()]
        info_lines.append(f"✗ Option '{option_to_select}' not found. Available: {available_options}")
        dropdown.click()
        time.sleep(0.5)
    finally:
        write_progress(0, "\n".join(info_lines))

def click_checkbox_with_progress(driver):
    info_lines = ["=== STEP 2: Clicking Checkbox ==="]
    wait = WebDriverWait(driver, 10)
    checkbox_xpath = "//div[contains(@class, 'p-checkbox-box')]"
    checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
    js_click(driver, checkbox)
    info_lines.append("✓ Checkbox clicked")
    write_progress(0, "\n".join(info_lines))
    time.sleep(1)

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
    if wait_for_results(driver, 10):
        print("  ✅ Results loaded")
    else:
        print("  ⚠️  Results timeout")
    time.sleep(1)

def robust_table_detection(driver, case_number):
    print(f"\n Checking for data in case {case_number}...")
    try:
        no_results = driver.find_elements(By.XPATH, "//p[contains(text(), 'لا توجد أية نتيجة للبحث')]")
        if no_results:
            for element in no_results:
                if element.is_displayed():
                    return "no_results", []
    except: pass
    try:
        table = driver.find_element(By.ID, "pr_id_16-table")
        if table.is_displayed():
            data = parse_table_by_element(table, case_number)
            if data:
                return "has_data", data
    except: pass
    try:
        all_tables = driver.find_elements(By.TAG_NAME, "table")
        for table in all_tables:
            if table.is_displayed():
                data = parse_table_by_element(table, case_number)
                if data:
                    return "has_data", data
    except: pass
    try:
        case_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '/') and contains(text(), '2025')]")
        if case_elements:
            return "possible_data", []
    except: pass
    try:
        loading = driver.find_elements(By.XPATH, "//*[contains(text(), 'جاري') or contains(text(), 'تحميل') or contains(text(), 'loading')]")
        if loading:
            return "loading", []
    except: pass
    return "unknown", []

def parse_table_by_element(table_element, case_number):
    data_rows = []
    rows = table_element.find_elements(By.TAG_NAME, "tr")
    for row in rows:
        if not row.is_displayed():
            continue
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) == 6 and not any(cell.get_attribute("colspan") for cell in cells):
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
            if row_data["case_number"] and '/' in row_data["case_number"]:
                data_rows.append(row_data)
    return data_rows

def run_scraper():
    start_resume = read_progress()
    if start_resume:
        start_n = start_resume + 1
        print(f"🔄 Resuming from number {start_n} (last progress: {start_resume})")
    else:
        start_n = START_NUM
        print(f" Starting from number {start_n}")
    driver = init_driver()
    driver.get(TARGET_URL)
    time.sleep(5)
    select_dropdown(driver, "اختيار محكمة الاستئناف", "محكمة الاستئناف بمراكش", 1)
    click_checkbox_with_progress(driver)
    select_dropdown(driver, "اختيار المحكمة الإبتدائية", "المحكمة الابتدائية بمراكش", 3)
    print("\n=== STEP 4: Selecting Police Unit ===")
    blank_dropdowns = driver.find_elements(By.XPATH, "//span[contains(@class, 'p-dropdown-label') and contains(@class, 'p-placeholder') and contains(text(), '---')]/ancestor::div[contains(@class, 'p-dropdown')]")
    if blank_dropdowns:
        first_dropdown = blank_dropdowns[0]
        js_click(driver, first_dropdown)
        time.sleep(1)
        options = driver.find_elements(By.XPATH, "//li[contains(@class, 'p-dropdown-item')]")
        available_options = [opt.text for opt in options if opt.text.strip()]
        selected_unit = "الدرك الملكي" if "الدرك الملكي" in available_options else available_options[0]
        for opt in options:
            if selected_unit in opt.text:
                js_click(driver, opt)
                time.sleep(1)
                write_progress(0, f"=== STEP 4: Selecting Police Unit ===\nAvailable police units: {available_options}\n✓ Selected police unit: {selected_unit}")
                break
    print("\n=== STEP 5: Selecting Police Station ===")
    if len(blank_dropdowns) >= 2:
        second_dropdown = blank_dropdowns[1]
        js_click(driver, second_dropdown)
        time.sleep(1)
        options = driver.find_elements(By.XPATH, "//li[contains(@class, 'p-dropdown-item')]")
        available_options = [opt.text for opt in options if opt.text.strip()]
        selected_station = "قائد مركز الدرك الملكي بايت اورير" if "قائد مركز الدرك الملكي بايت اورير" in available_options else available_options[0]
        for opt in options:
            if selected_station in opt.text:
                js_click(driver, opt)
                time.sleep(1)
                write_progress(0, f"=== STEP 5: Selecting Police Station ===\nAvailable police stations: {available_options}\n✓ Selected police station: {selected_station}")
                break
    for n in range(start_n, END_NUM + 1):
        write_progress(n)
        attempt = 0
        success = False
        while not success and attempt < RETRIES:
            attempt += 1
            try:
                fill_case_details(driver, n, YEAR)
                status, rows_data = robust_table_detection(driver, n)
                if status == "no_results":
                    print(f"🚫 [NO RESULTS] {n}")
                    success = True
                elif status == "has_data" and rows_data:
                    append_to_excel(rows_data)
                    print(f"✅ [FOUND] {n} -> {len(rows_data)} rows")
                    success = True
                elif status == "possible_data":
                    print(f"🔍 [POSSIBLE DATA] {n}")
                    if attempt == RETRIES:
                        success = True
                elif status == "loading":
                    print(f"⏳ [STILL LOADING] {n} - Retrying...")
                    time.sleep(2)
                else:
                    print(f"❓ [UNKNOWN: {status}] {n}")
                    if attempt == RETRIES:
                        success = True
                time.sleep(MIN_DELAY + random.random() * (MAX_DELAY - MIN_DELAY))
            except Exception as e:
                print(f"Error {n}: {e}")
    print("\n✅ Scraping completed!")
    driver.quit()

if __name__ == "__main__":
    run_scraper()

