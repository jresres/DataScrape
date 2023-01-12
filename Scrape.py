# IMPORTS
import time

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
from pandas import ExcelWriter
from xlsxwriter.utility import xl_rowcol_to_cell
import numpy as np
import time
import csv


# DIV TREE CLASS OF ALL CATEGORIES
class_name = "_p13n-zg-nav-tree-all_style_zg-browse-item__1rdKf _p13n-zg-nav-tree-all_style_zg-browse-height-small__nleKL"

i = 0                   # index
category = ''           # category string
product_name_list = []  # product name list array
ranking_list = []       # ranking list array

# Get user input on which category to scrape data on
while (i == 0):
    category = input("Enter '1' for Health Care. Enter '2' for Personal Care\n")
    if (category == '1' or category == '2'):
        break

# Initialize chrome web driver location and url
driver = webdriver.Chrome("/Users/jaredreiss/Desktop/Programming/Amazon_Datascrape/chromedriver")
driver.get("https://www.amazon.com/Best-Sellers/zgbs/ref=zg_bs_unv_hpc_0_3764461_3")

# Explicit wait for page to find class element necessary
try:
    test = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div[2]/div/div[2]/div/div/div[1]/span')))
finally:
    print("FATAL ERROR: Page never loaded. Try looking at the URL entered as driver root")

health_household_button = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div/div/div/div[2]/div/div[2]/div/div/div[2]/div[25]/a')
health_household_button.click()

writer = ExcelWriter('product_data.xlsx', engine='xlsxwriter')

time.sleep(5)

if (category == 'health'):
    health_care_button = driver.find_element_by_css_selector('#zg_browseRoot > ul > ul > li:nth-child(2) > a')
    health_care_button.click()

if (category == 'personal'):
    personal_care_button = driver.find_element_by_css_selector('#zg_browseRoot > ul > ul > li:nth-child(4) > a')
    personal_care_button.click()

time.sleep(5)

while i < 2:
    try:
        items = driver.find_elements_by_css_selector('li.zg-item-immersion')

        for item in items:
            product_name = item.find_element_by_css_selector('div.p13n-sc-truncated').text
            ranking = int(item.find_element_by_css_selector('span.zg-badge-text').text.lstrip('#'))

            product_name_list.append(product_name)
            ranking_list.append(ranking)

        driver.find_element_by_css_selector('li.a-last').click()
        time.sleep(5)
        i += 1

    except exceptions.StaleElementReferenceException:
        pass

#print(ranking_list)
data = {'Product Name': product_name_list, 'Ranking': ranking_list}
df = pd.DataFrame(data = data)
df.to_excel(writer, 'Report', index = False) #MAKE SURE TO COPY DATA IF ERROR OCCURS AND PUT IN NEW SPREADSHEET

workbook = writer.book
worksheet = writer.sheets['Report']

worksheet.set_zoom(90)
worksheet.set_column('A:A', 80)

ranking_format = workbook.add_format({'align': 'middle', 'bold': True, 'bottom':6})

print(df)

writer.save()
time.sleep(10)

driver.close()
