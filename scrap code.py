import subprocess
import sys

import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import pandas as pd
import re
import time

# parameter setting
def int_to_date_format(num):
    if not isinstance(num, int):
        raise ValueError("Input must be an integer.")
    year = (num // 10000) + 2000  # after 2000
    month = (num // 100) % 100
    day = num % 100
    date_str = f"{year}-{month:02d}-{day:02d}"
    return date_str

now = datetime.datetime.now()
formatdate = now.strftime("%y%m%d")
if now.hour > 15:
    DATE_Select = int(formatdate)
else:
    DATE_Select = int(formatdate) - 1

# Setup ChromeDriver in my computer
chrome_driver_path = 'C:/Program Files/Google/Chrome/Application/chromedriver.exe'
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service)

# Function to calculate the starting index for each page
def get_page_index(page):
    return (page - 1) * 24 # page added pattern

total_pages=42
# this example is about "Flat" in Hillingdon Borough
base_url = 'https://www.rightmove.co.uk/property-for-sale/find.html?locationIdentifier=REGION%5E93959&sortType=6&propertyTypes=Flat&includeSSTC=false&mustHave=&dontShow=&furnishTypes=&keywords=&index='
data = []

# Loop through pages
for page in range(1, total_pages + 1):
    index = get_page_index(page)
    url = base_url + str(index)
    driver.get(url)

    # Wait for the page to load
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, 'propertyCard')))
    except TimeoutException:
        print(f"Timeout occurred for page {page}")
        continue

    # Extract id, price and link information using Selenium
    properties = driver.find_elements(By.CLASS_NAME, 'propertyCard')
    for property in properties:
        try:
            anchor_tag = property.find_element(By.CLASS_NAME, 'propertyCard-anchor')
            property_id_full = anchor_tag.get_attribute('id')
            property_id = re.findall(r'\d+', property_id_full)[0] if property_id_full else 'N/A'
            price = property.find_element(By.CLASS_NAME, 'propertyCard-priceValue').text.strip()
            href_tag = property.find_element(By.CLASS_NAME, 'propertyCard-link')
            href = href_tag.get_attribute('href')
            property_link = href if href else 'N/A'

            # open and scrap in new web

            # Visit the property link to extract house type, bed, bath, size, tenure, council tax, parking, garden, and nearest station information
            if property_link != 'N/A':
                driver.get(property_link)
                time.sleep(3)  # Wait for the new page to load
                try:
                    house_type_tag = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, '_1hV1kqpVceE9m-QrX_hWDN'))
                    )
                    house_type = house_type_tag.text.strip() if house_type_tag else 'N/A'
                except TimeoutException:
                    house_type = 'N/A'

                try:
                    bedrooms = driver.find_element(By.XPATH, "//span[text()='BEDROOMS']/following::p").text.strip()
                except NoSuchElementException:
                    bedrooms = 'N/A'

                try:
                    bathrooms = driver.find_element(By.XPATH, "//span[text()='BATHROOMS']/following::p").text.strip()
                except NoSuchElementException:
                    bathrooms = 'N/A'

                try:
                    size = driver.find_element(By.XPATH, "//span[text()='SIZE']/following::p").text.strip()
                except NoSuchElementException:
                    size = 'N/A'

                try:
                    tenure = driver.find_element(By.XPATH, "//span[text()='TENURE']/following::p").text.strip()
                except NoSuchElementException:
                    tenure = 'N/A'

                try:
                    council_tax = driver.find_element(By.XPATH,
                                                      "//dt[text()='COUNCIL TAX']/following::dd[1]").text.strip()
                except NoSuchElementException:
                    council_tax = 'N/A'

                try:
                    parking = driver.find_element(By.XPATH, "//dt[text()='PARKING']/following::dd[1]").text.strip()
                except NoSuchElementException:
                    parking = 'N/A'

                try:
                    garden = driver.find_element(By.XPATH, "//dt[text()='GARDEN']/following::dd[1]").text.strip()
                except NoSuchElementException:
                    garden = 'N/A'

                try:
                    nearest_stations_container = driver.find_element(By.CLASS_NAME, '_2f-e_tRT-PqO8w8MBRckcn')
                    nearest_stations_info = nearest_stations_container.text.strip()
                except NoSuchElementException:
                    nearest_stations_info = 'N/A'

                driver.back()  # Go back to the main listings page
                time.sleep(2)  # Wait for the page to reload
            else:
                house_type = 'N/A'
                bedrooms = 'N/A'
                bathrooms = 'N/A'
                size = 'N/A'
                tenure = 'N/A'
                council_tax = 'N/A'
                parking = 'N/A'
                garden = 'N/A'
                nearest_stations_info = 'N/A'

            data.append(
                {'ID': property_id, 'Price': price, 'HouseType': house_type, 'Beds': bedrooms, 'Baths': bathrooms,
                 'Size': size,
                 'Tenure': tenure, 'CouncilTax': council_tax, 'Parking': parking, 'Garden': garden,
                 'NearestStations': nearest_stations_info, 'Link': property_link})
        except (NoSuchElementException, IndexError, StaleElementReferenceException):
            continue

# Convert to DataFrame and save to Excel
df = pd.DataFrame(data)
df.to_excel('D:\\Hillingdon.xlsx', index=False)
# When we crawled the data for the three types of rooms in the other 32 boroughs separately
# the borough id and the type of house were replaced

# Close the driver
driver.quit()
