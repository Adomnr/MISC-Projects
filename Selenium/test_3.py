from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import time

# Bypass SSL verification
os.environ['WDM_SSL_VERIFY'] = '0'

# Initialize the WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the target URL
url = "https://us.soccerway.com/teams/comparison/?team_ids%5B%5D=16210&team_ids%5B%5D=16210"
driver.get(url)

# Function to extract options from a dropdown
def get_dropdown_options(dropdown):
    select = Select(dropdown)
    options = [option.text for option in select.options]
    return options

# Function to print options for debugging
def print_dropdown_options(dropdown_index):
    dropdown = Select(driver.find_elements(By.CSS_SELECTOR, "select")[dropdown_index])
    options = [option.text for option in dropdown.options]
    print(f"Options in Dropdown {dropdown_index + 1}: {options}")
    return options

# Function to select an option and extract data from dependent dropdowns
def extract_data_for_selection(dropdown_index, option):
    try:
        dropdowns = driver.find_elements(By.CSS_SELECTOR, "select")
        dropdown = Select(dropdowns[dropdown_index])
        
        # Ensure dropdown is visible and options are loaded
        WebDriverWait(driver, 10).until(EC.visibility_of(dropdown.first_selected_option))
        
        # Retry mechanism for selecting options
        attempts = 3
        while attempts > 0:
            try:
                dropdown.select_by_visible_text(option)
                break
            except Exception as e:
                print(f"Retrying selection due to: {e}")
                time.sleep(1)
                attempts -= 1

        # Wait for the page to update
        WebDriverWait(driver, 10).until(EC.staleness_of(dropdown))

        # Re-find the dependent dropdown elements
        dependent_dropdowns = driver.find_elements(By.CSS_SELECTOR, "select")

        # Extract data from the dependent dropdowns
        dependent_data = {}
        dependent_data['Dropdown 8'] = get_dropdown_options(dependent_dropdowns[7])
        dependent_data['Dropdown 9'] = get_dropdown_options(dependent_dropdowns[8])
        dependent_data['Dropdown 11'] = get_dropdown_options(dependent_dropdowns[10])
        dependent_data['Dropdown 12'] = get_dropdown_options(dependent_dropdowns[11])
        
        return dependent_data
    except Exception as e:
        print(f"Error selecting {option} in Dropdown {dropdown_index + 1}: {e}")
        return {}

# Locate the initial dropdown elements
dropdown7 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[6])
dropdown10 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[9])

# Print options for debugging
print_dropdown_options(6)  # Dropdown 7
print_dropdown_options(9)  # Dropdown 10

dropdown7_options = [option.text for option in dropdown7.options if option.text]
dropdown10_options = [option.text for option in dropdown10.options if option.text]

# Dictionary to store the extracted data
all_data = {}

# Iterate over each option in Dropdown 7 and extract data
for option in dropdown7_options:
    print(f"Selecting {option} in Dropdown 7")
    all_data[f"Dropdown 7 - {option}"] = extract_data_for_selection(6, option)

# Iterate over each option in Dropdown 10 and extract data
for option in dropdown10_options:
    print(f"Selecting {option} in Dropdown 10")
    all_data[f"Dropdown 10 - {option}"] = extract_data_for_selection(9, option)

# Print the extracted data
for key, value in all_data.items():
    print(f"{key}:")
    for dropdown, options in value.items():
        print(f"  {dropdown}: {options}")

# Close the WebDriver
driver.quit()
