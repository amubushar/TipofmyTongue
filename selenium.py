# Author: Abdullah Mubushar 6/7/2020
# Purpose: Utilizes scripted dom based browser to retrieve data
# Installation:
# 1) pip install selenium
# 2) Download chromedriver.exe (https://sites.google.com/a/chromium.org/chromedriver/downloads) for your chrome version (probably 83)
# 3) Extract and copy driver to your Python scripts folder, usually %appdata%/Python/Python37/Scripts
# 4) Modify the following variables:
#       "driver_path" to specify your driver location as in step 3
#       "download_dir" to where it should save the download file
#       "start_date", "end_date" to specify the data retrieval range
# 5) Save, Run, ???, Profit!
# Note: script has several sleep delays implemented to allow website to load properly, if you are on site and not using a VPN you may reduce them, it 

import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from os import path

#Driver location
driver_path = path.expandvars(r"%appdata%/Python/Python37/Scripts/chromedriver.exe")

#Download location
download_dir = path.expandvars(r'%USERPROFILE%\Desktop')

#Data fetch date
start_date = "2020-Mar-02"
end_date = "2020-Jun-04"
scotford_id = '3'

#Base urls
url = ""

#Account details
acct = "acc"
passw = "password"

#Run chrome in domless mode
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')

#Maximum time before cancellation
timeout = 60 

try:
    driver = webdriver.Chrome()
except:
    driver = webdriver.Chrome(driver_path, options=chrome_options)

driver.get(url)

try:
    element_present = EC.presence_of_element_located((By.NAME, "USER"))
    WebDriverWait(driver, timeout).until(element_present)
    print("Successfully connected to server")
except TimeoutException:
    print("Timed out waiting for page to load")
    exit()

#Authentication sequence
elem = driver.find_element_by_name("USER")
elem.send_keys(acct)
elem = driver.find_element_by_name("PASSWORD")
elem.send_keys(passw)

#Submit login form
elem.send_keys(Keys.RETURN)

try:
    element_present = EC.presence_of_element_located((By.NAME, "ctl00$contentPlaceHolder$ctl00$availableLocations"))
    WebDriverWait(driver, timeout).until(element_present)
    print("Successfully authenticated")
except TimeoutException:
    print("Timed out waiting for page to load")
    exit()

#Set Location
elem = driver.find_element_by_name("ctl00$contentPlaceHolder$ctl00$availableLocations")
elem.send_keys(scotford_id)

#Set start date
elem = driver.find_element_by_name('ctl00$contentPlaceHolder$ctl00$startGasDayPicker$gasDayTextBox')
elem.clear()
elem.send_keys(start_date)

#Set end date
elem = driver.find_element_by_name('ctl00$contentPlaceHolder$ctl00$endGasDayPicker$gasDayTextBox')
elem.clear()
elem.send_keys(end_date)

#Submit data select form
elem.send_keys(Keys.RETURN)

#Browse to pop-up window once loaded
while(len(driver.window_handles) < 2):
    pass
driver.switch_to.window(driver.window_handles[-1])

#set download behavior
params = {'behavior': 'allow', 'downloadPath': download_dir}
driver.execute_cdp_cmd('Page.setDownloadBehavior', params)

#Locate csv download by full xpath
try:
    element_present = EC.presence_of_element_located((By.XPATH, '/html/body/form/div[3]/span/div/table/tbody/tr[4]/td/div/div/div[4]/table/tbody/tr/td/div[2]/div[1]/a'))
    WebDriverWait(driver, timeout).until(element_present)
    print("Successfully retrieved data")
except TimeoutException:
    print("Timed out waiting for page to load")
    exit()

#download csv file, can be changed to pdf
while(True):
    try:
        driver.execute_script("$find('ctl00_reportViewer').exportReport('CSV');")
        break
    except:
        time.sleep(0.5)
        pass

print("Downloaded file successfully at: " + str(download_dir))