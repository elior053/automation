import os
import re
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class SurplexScraper:
    def __init__(self, output_file):
        self.output_file = output_file
        self.driver = self.setup_driver()

    def setup_driver(self):
        service = Service(r"\\SERVER2022-DC\Workload_Data\pythonProject\chromedriver.exe")
        driver = webdriver.Chrome(service=service)
        return driver

    def get_next_id(self):
        if os.path.exists(self.output_file):
            df = pd.read_excel(self.output_file)
            if not df.empty:
                return df['ID'].max() + 1
        return 1

    def url_already_exists(self, url):
        if not os.path.exists(self.output_file):
            return False
        df = pd.read_excel(self.output_file)
        return url in df['full_url'].values

    def scrape(self, search_term):
        next_id = self.get_next_id()
        self.driver.get("https://www.surplex.com/")
        self.driver.maximize_window()
        time.sleep(3)
        wait = WebDriverWait(self.driver, 10)

        if not os.path.exists(self.output_file):
            df = pd.DataFrame(columns=["ID", "manufacture", "model", "year", "price", "currency", "image", "full_url"])
        else:
            df = pd.read_excel(self.output_file)

        self.accept_cookies(wait)
        search_results = self.perform_search(wait, search_term)

        for result in search_results:
            href = result.get_attribute("href")
            full_url = href
            self.driver.execute_script("window.open(arguments[0]);", full_url)
            self.driver.switch_to.window(self.driver.window_handles[-1])
            print(f"Success: Opened {full_url}")

            price, currency = self.extract_price()
            if price == "Not available":
                print("Price not available, skipping to the next item")
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[0])
                continue

            year = self.extract_year_of_manufacture()
            model = self.extract_model(search_term)
            manufacture = search_term

            if self.url_already_exists(full_url):
                print(f"URL already exists in the file: {full_url}")
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[0])
                continue

            screenshot_filename = self.save_screenshot(next_id)
            new_row = pd.DataFrame({
                "ID": [next_id],
                "manufacture": [manufacture],
                "model": [model],
                "year": [year],
                "price": [price],
                "currency": [currency],
                "image": [screenshot_filename],
                "full_url": [href],
                "Date": [datetime.now().date()]
            })
            df = pd.concat([df, new_row], ignore_index=True)
            next_id += 1

            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(2)

        df.to_excel(self.output_file, index=False)

    def accept_cookies(self, wait):
        try:
            accept_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn.btn--primary.js-consentText.js-acceptAllBtn")))
            accept_button.click()
            print("Cookie consent button clicked successfully.")
        except Exception as e:
            print(f"Error clicking consent button: {e}")

    def perform_search(self, wait, search_term):
        element_search = wait.until(EC.visibility_of_element_located((By.ID, "searchInput")))
        element_search.clear()
        element_search.send_keys(search_term)
        element_search.send_keys(Keys.ENTER)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".js-article-watch--item")))
        return self.driver.find_elements(By.CSS_SELECTOR, ".js-article-watch--item .cardProduct__title a.item-url")

    def extract_price(self):
        try:
            price_element = self.driver.find_element(By.CLASS_NAME, "bidBox__price")
            price_text = price_element.text.strip()
            match = re.search(r"([\d.,]+)\s*([€$£])", price_text)
            if match:
                price = match.group(1)
                currency = match.group(2)
            else:
                price = "Not available"
                currency = ""
        except NoSuchElementException:
            print("no price")
            price = "Not available"
            currency = ""
        return price, currency

    def extract_year_of_manufacture(self):
        try:
            year_of = self.driver.find_element(By.XPATH, "//dt[contains(text(), 'Year of manufacture')]")
            year_of = year_of.find_element(By.XPATH, "following-sibling::dd")
            year = year_of.text
        except NoSuchElementException:
            year = "Not available"
        return year

    def extract_model(self, search_term):
        try:
            name_dt = self.driver.find_element(By.XPATH, "//dt[contains(text(), 'Name')]")
            model_dd = name_dt.find_element(By.XPATH, "following-sibling::dd")
            model = model_dd.text.replace(search_term, "").strip()
        except NoSuchElementException:
            model = "Not available"
        return model

    def save_screenshot(self, next_id):
        self.driver.execute_script("document.body.style.transform = 'scale(0.45)';")
        self.driver.execute_script("document.body.style.transformOrigin = '0 0';")
        screenshot_filename = f"{next_id}.png"
        screenshot_path = f"\\\\SERVER2022-DC\\Workload_Data\\pythonProject\\screenshots\\{screenshot_filename}"
        self.driver.save_screenshot(screenshot_path)
        return screenshot_filename

    def quit_driver(self):
        self.driver.quit()


def surplex():
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    print(f'Hey Elior, I\'m on = {current_time}')

    output_file = 'OUTPUT.xlsx'
    file_path = 'manufacture.xlsx'

    if not os.path.exists(file_path):
        print(f"File {file_path} not found.")
        return

    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        search_term = row["שמות יצרני מכונות"]
        scraper = SurplexScraper(output_file)
        scraper.scrape(search_term)
        scraper.quit_driver()


if __name__ == "__main__":
    surplex()