import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Initialize ChromeOptions and specify the user profile directory
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("user-data-dir=C:\\Users\\dharshini\\AppData\\Local\\Google\\Chrome\\User Data")
chrome_options.add_argument("profile-directory=Profile 7")  # Specify the exact profile directory here

# Initialize Chrome WebDriver with the specified profile
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Read Excel file
excelFilePath = "Portfoli_testing.xlsx"
excelData = pd.read_excel(excelFilePath)
# Extract project URL
driver.get(excelData["PROJECT URL"][0])
time.sleep(3)
resultList = []

# Iterate over each row in the DataFrame and perform actions in Selenium
for index, row in excelData.iterrows():
    action = row['ACTION']
    element_type = row['ELEMENT TYPE']
    element_path = row['ELEMENT PATH']
    try:
        if row['ELEMENT TYPE'] == "XPATH":
            clicked_element = driver.find_element(By.XPATH, element_path)
            print(clicked_element.text)
            clicked_element.click()
        elif row['ELEMENT TYPE'] == "ID":
            clicked =  driver.find_element(By.ID, element_path)
            print(clicked.text)
            clicked.click()
        resultList.append("Pass")
    except Exception as e:
        resultList.append("Fail")
    time.sleep(3)

excelData["RESULT"] = resultList
excelData.to_excel(excelFilePath, index=False)

