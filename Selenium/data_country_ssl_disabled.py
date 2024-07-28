from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import ssl


import os

# Bypass SSL verification
os.environ['WDM_SSL_VERIFY'] = '0'

# Disable SSL verification
#ssl._create_default_https_context = ssl._create_unverified_context

# Initialize the WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the target URL
url = "https://us.soccerway.com/matches/2024/07/23/korea-republic/k-league/seongnam-ilhwa/chunnam-dragons/4308862/head2head/"  # replace with your actual URL
driver.get(url)

# Wait for the page to load completely
driver.implicitly_wait(10)

# Function to extract options from a dropdown
def get_dropdown_options(dropdown):
    select = Select(dropdown)
    options = [option.text for option in select.options]
    return options

# Locate the country dropdown element
country_dropdown = Select(driver.find_element(By.CSS_SELECTOR, "select[name='country']"))

# Extract all country options
countries = [option.text for option in country_dropdown.options]

# Dictionary to store all dropdown data for each country
all_dropdown_data = {}

# Iterate through each country and extract dropdown data
for country in countries:
    # Select the country
    country_dropdown.select_by_visible_text(country)
    
    # Wait for the page to update (You might need to add more wait time depending on the website's response time)
    driver.implicitly_wait(5)
    
    # Locate all dropdown elements again as the DOM might have updated
    dropdowns = driver.find_elements(By.CSS_SELECTOR, "select")
    
    # Extract data from each dropdown
    dropdown_data = {}
    for i, dropdown in enumerate(dropdowns):
        dropdown_data[f"Dropdown {i+1}"] = get_dropdown_options(dropdown)
    
    # Store the data for the current country
    all_dropdown_data[country] = dropdown_data

# Print the extracted data for each country
for country, data in all_dropdown_data.items():
    print(f"Data for {country}:")
    for dropdown, options in data.items():
        print(f"  {dropdown}: {options}")

# Close the WebDriver
driver.quit()
