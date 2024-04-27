import os
import re
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class BimedisScraper:
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
        self.driver.get("https://bimedis.com/")
        self.driver.maximize_window()
        time.sleep(3)
        wait = WebDriverWait(self.driver, 10)

        if not os.path.exists(self.output_file):
            df = pd.DataFrame(columns=["ID", "manufacture", "model", "year", "price", "currency", "image", "full_url"])
        else:
            df = pd.read_excel(self.output_file)

        self.accept_cookies(wait)
        search_results = self.perform_search(wait, search_term)

        if not search_results:
            print(f"No search results found for '{search_term}'")
            return

        original_window = self.driver.current_window_handle
        processed_urls = set()

        for result in search_results:
            href = result.get_attribute("href")
            if href not in processed_urls:
                processed_urls.add(href)
                self.driver.execute_script("window.open(arguments[0]);", href)
                self.driver.switch_to.window(self.driver.window_handles[-1])
                print(f"Success: Opened {href} in a new tab.")

                price, currency = self.extract_price(wait)
                year = self.extract_year_of_manufacture()
                model = self.extract_model()
                manufacture = search_term

                if self.url_already_exists(href):
                    print(f"URL already exists in the file: {href}")
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

        self.driver.quit()

    def accept_cookies(self, wait):
        try:
            agree_button = wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                   "//button[contains(@class, 'v-btn') and contains("
                                                                   "@class, 'yellow') and contains(., 'I AGREE')]")))
            agree_button.click()
            print("Cookie consent button clicked successfully.")
        except Exception as e:
            print(f"Error clicking consent button: {e}")

    def perform_search(self, wait, search_term):
        element_search = wait.until(EC.visibility_of_element_located((By.ID, "input-275")))
        element_search.clear()
        element_search.send_keys(search_term)
        element_search.send_keys(Keys.ENTER)

        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[target='_blank'][href^='/a-item/']"))
            )
            search_results = self.driver.find_elements(By.CSS_SELECTOR, "a[target='_blank'][href^='/a-item/']")
        except Exception:
            search_results = []

        return search_results

    def extract_price(self, wait):
        try:
            price_container = wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "div[data-v-233ee7d8] span.cur")))
            price_text = price_container.text.strip().replace(" ", "")
            currency_symbol = price_container.get_attribute("data-before")

            match = re.search(r"([\d.,]+)", price_text)
            if match:
                price = match.group(1)
                currency = currency_symbol if currency_symbol else ""
            else:
                price = "Not available"
                currency = ""
            print(f"Price: {price}, Currency: {currency}")
        except NoSuchElementException:
            print("Price information not available for this item.")
            price = "Not available"
            currency = ""
        return price, currency

    def extract_year_of_manufacture(self):
        try:
            year_container = self.driver.find_element(By.XPATH,
                                                      "//div[h5[@class='advert_properties']//div[contains(text(), 'Year:')]]/span[@class='list-value']")
            year = year_container.text.strip()
        except NoSuchElementException:
            year = "Not available"
        return year

    def extract_model(self):
        try:
            model_container = self.driver.find_element(By.XPATH, "//tr[.//div[contains(text(), 'Model')]]")
            model_link = model_container.find_element(By.XPATH, ".//td[@align='right']/a")
            model = model_link.text.strip()
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


def bimedis():
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
        search_term = row["שמות יצרני ציוד רפואי"]
        scraper = BimedisScraper(output_file)
        scraper.scrape(search_term)


if __name__ == "__main__":
    bimedis()