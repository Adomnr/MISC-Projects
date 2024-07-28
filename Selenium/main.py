from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import os

# Bypass SSL verification
os.environ['WDM_SSL_VERIFY'] = '0'

# Initialize the WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the target URL
url = "https://us.soccerway.com/teams/comparison/?team_ids%5B%5D=14495&team_ids%5B%5D=16210"
driver.get(url)

# Wait for the page to load completely
driver.implicitly_wait(10)

# Function to extract options from a dropdown
def get_dropdown_options(dropdown):
    select = Select(dropdown)
    options = [option.text for option in select.options]
    return options

# Locate the dropdown elements
dropdowns = driver.find_elements(By.CSS_SELECTOR, "select")

# Extract data from each dropdown
dropdown_data = {}
for i, dropdown in enumerate(dropdowns):
    dropdown_data[f"Dropdown {i+1}"] = get_dropdown_options(dropdown)

# Print the extracted data
for dropdown, options in dropdown_data.items():
    print(f"{dropdown}: {options}")

# Close the WebDriver
driver.quit()
