import time
import random
import pandas as pd
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# =========================================================
# CONFIG
# =========================================================
# NOTE: Update these paths for your machine before running.
# Keep them relative in GitHub projects (recommended).
input_file = r"ve_links.xlsx"                 # Input Excel containing the list of links
output_file = r"ve_links_output.xlsx"         # Output Excel where results are saved

links_column = "Link"
batch_size = 3
delay_min = 9
delay_max = 13
WAIT_TIME = 15

# =========================================================
# VERIFIED XPATHS (Fields to extract)
# =========================================================
xpaths = {
    "Exhibitor Name": ['//div[contains(@class,"details-header")]//h1'],
    "Stand(s)": ['//div[contains(@class,"stand-details")]//span[@class="stand-reference-label"][2]'],
    "Company Website": ['//div[contains(@class,"exhibitor-details-contact-us-links")]//a[starts-with(@href,"http")]'],
    "Email": ['//a[starts-with(@href,"mailto:")]'],
    "Contact Number": ['//a[starts-with(@href,"tel:")]'],
    "Full Address": ['//div[@id="exhibitor_details_address"]//span'],
    "LinkedIn": ['//a[contains(@data-dtm,"follow_linkedin")]'],
    "Description": ['//div[@id="exhibitor_details_description"]//p']
}

# =========================================================
# FUNCTIONS
# =========================================================
def init_driver():
    """Initialize Chrome WebDriver with basic settings."""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    return webdriver.Chrome(service=Service(), options=chrome_options)


def get_element_texts(driver, xpath_list, get_href=False, join_with=", "):
    """
    Extract text or href values for a list of XPaths.
    Returns a joined string or None if nothing found.
    """
    texts = []
    for xp in xpath_list:
        elements = driver.find_elements(By.XPATH, xp)
        for el in elements:
            txt = el.get_attribute("href") if get_href else el.text.strip()
            if txt:
                texts.append(txt.strip())
    return join_with.join(texts) if texts else None


def wait_for_element(driver, by, value):
    """Wait for a specific element to appear."""
    try:
        WebDriverWait(driver, WAIT_TIME).until(
            EC.presence_of_element_located((by, value))
        )
        return True
    except TimeoutException:
        return False


def wait_for_dom_stability(driver, stable_seconds=3, timeout=30):
    """
    Wait until page HTML stops changing (useful for Vue/React re-rendering).
    Applied once per page.
    """
    start_time = time.time()
    last_html = driver.page_source
    stable_start = time.time()

    while True:
        time.sleep(0.5)
        current_html = driver.page_source

        if current_html != last_html:
            last_html = current_html
            stable_start = time.time()
        elif time.time() - stable_start >= stable_seconds:
            return True

        if time.time() - start_time > timeout:
            return False


def detect_captcha(driver):
    """
    Detect reCAPTCHA iframe. If found, pause so user can solve manually.
    """
    try:
        driver.find_element(By.XPATH, '//iframe[contains(@src,"recaptcha")]')
        print("⚠ CAPTCHA detected. Solve manually...")
        while True:
            time.sleep(3)
            try:
                driver.find_element(By.XPATH, '//iframe[contains(@src,"recaptcha")]')
            except NoSuchElementException:
                break
        return True
    except NoSuchElementException:
        return False


# =========================================================
# RESUME SUPPORT (Continue from existing output)
# =========================================================
if os.path.exists(output_file):
    existing_df = pd.read_excel(output_file)
    processed_links = set(existing_df["Link"].dropna().tolist())
    results = existing_df.to_dict("records")
    print(f"🔁 Resuming… {len(processed_links)} links already processed")
else:
    processed_links = set()
    results = []

# =========================================================
# MAIN SCRIPT
# =========================================================
df_links = pd.read_excel(input_file)
links = df_links[links_column].dropna().unique().tolist()
total_links = len(links)
count = len(processed_links)

for i in range(0, total_links, batch_size):
    driver = init_driver()

    for link in links[i:i + batch_size]:
        if link in processed_links:
            continue

        count += 1
        print(f"\n[{count}/{total_links}] Processing: {link}")
        processed_links.add(link)

        try:
            driver.get(link)

            if detect_captcha(driver):
                continue

            # Base signal: wait for header
            WebDriverWait(driver, WAIT_TIME).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//div[contains(@class,"details-header")]//h1')
                )
            )

            # DOM stability wait (once per page)
            wait_for_dom_stability(driver)

            data_row = {"Link": link}

            for field, xp_list in xpaths.items():

                # Explicit waits only where needed
                if field == "Full Address":
                    wait_for_element(driver, By.ID, "exhibitor_details_address")
                    value = get_element_texts(driver, xp_list, join_with=", ")

                elif field == "Description":
                    wait_for_element(driver, By.ID, "exhibitor_details_description")
                    value = get_element_texts(driver, xp_list)

                elif field in ["Company Website", "LinkedIn"]:
                    value = get_element_texts(driver, xp_list, get_href=True)

                elif field == "Email":
                    value = get_element_texts(driver, xp_list, get_href=True)
                    if value:
                        value = value.replace("mailto:", "")

                elif field == "Contact Number":
                    value = get_element_texts(driver, xp_list, get_href=True)
                    if value:
                        value = value.replace("tel:", "")

                else:
                    value = get_element_texts(driver, xp_list)

                data_row[field] = value if value else "No Data On Link"
                print(f"{field}: {data_row[field]}")

            results.append(data_row)

            # Save after every link (crash-safe)
            pd.DataFrame(results).to_excel(output_file, index=False)

        except Exception as e:
            print(f"❌ Error on {link}: {e}")
            pd.DataFrame(results).to_excel(output_file, index=False)
            continue

    driver.quit()
    time.sleep(random.uniform(delay_min, delay_max))

print(f"\n✅ COMPLETED. Data saved to: {output_file}")