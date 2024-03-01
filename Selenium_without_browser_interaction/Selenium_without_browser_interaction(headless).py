import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

# Set Chrome options for headless mode and user data directory
chrome_options = Options()
chrome_options.add_argument('--headless')  # Run in headless mode
chrome_options.add_argument("--no-sandbox")  # Bypass OS security model
chrome_options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
chrome_options.add_argument('--disable-gpu')  # Disable GPU acceleration


# Set user data directory and profile directory
user_data_dir = 'C:\\Users\\dharshini\\AppData\\Local\\Google\\Chrome\\User Data'
profile_directory = 'Profile 7'
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
chrome_options.add_argument(f"--profile-directory={profile_directory}")

# Create Chrome webdriver
driver = webdriver.Chrome(options=chrome_options)

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
            clicked = driver.find_element(By.ID, element_path)
            print(clicked.text)
            clicked.click()
        resultList.append("Pass")
    except Exception as e:
        resultList.append("Fail")
    time.sleep(3)

# Update Excel with results
excelData["RESULT"] = resultList
excelData.to_excel(excelFilePath, index=False)

# Close the browser
driver.quit()