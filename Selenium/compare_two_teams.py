from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import time

# Bypass SSL verification
os.environ['WDM_SSL_VERIFY'] = '0'

# Initialize the WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the target URL
url = "https://us.soccerway.com/matches/2024/07/23/korea-republic/k-league/seongnam-ilhwa/chunnam-dragons/4308862/head2head/"
driver.get(url)

# Wait for the page to load completely
driver.implicitly_wait(4)


# Function to extract options from a dropdown
def get_dropdown_options(dropdown):
    select = Select(dropdown)
    options = [option.text for option in select.options]
    return options



# Function to select an option and extract data from dependent dropdowns
def extract_data_for_selection(dropdown_index, option):
    dropdown = Select(driver.find_elements(By.CSS_SELECTOR, "select")[dropdown_index])
    dropdown.select_by_visible_text(option)
    time.sleep(2)  # Wait for the page to update

    # Re-find the dependent dropdown elements
    dependent_dropdowns = driver.find_elements(By.CSS_SELECTOR, "select")

    # Extract data from the dependent dropdowns
    dependent_data = {}
    dependent_data['Dropdown 8'] = get_dropdown_options(dependent_dropdowns[7])
    dependent_data['Dropdown 9'] = get_dropdown_options(dependent_dropdowns[8])
    dependent_data['Dropdown 11'] = get_dropdown_options(dependent_dropdowns[10])
    dependent_data['Dropdown 12'] = get_dropdown_options(dependent_dropdowns[11])

    return dependent_data

# Locate the initial dropdown elements
dropdown7 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[6])
# dropdown10 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[9])

dropdown7_options = [option.text for option in dropdown7.options if option.text]
# dropdown10_options = [option.text for option in dropdown10.options if option.text]

# Dictionary to store the extracted data
all_data = {}

def compare_two_teams_function(selection,country1,competition1,team1,country2,competition2,team2):
    if selection == "Club":

        for option in dropdown7_options:
            print(f"Selecting {option} in Dropdown 7")
            all_data[f"Dropdown 7 - {option}"] = extract_data_for_selection(6, option)
            if option == country1:
                break
        time.sleep(1)
        dropdown8 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[7])
        dropdown8_options = [option2.text for option2 in dropdown8.options if option2.text]
        print(dropdown8_options)

        for option in dropdown8_options:
            print(f"Selecting {option} in Dropdown 8")
            all_data[f"Dropdown 8 - {option}"] = extract_data_for_selection(7, option)
            if option == competition1:
                break
        time.sleep(1)
        dropdown9 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[8])
        dropdown9_options = [option.text for option in dropdown9.options if option.text]
        for option in dropdown9_options:
            print(f"Selecting {option} in Dropdown 9")
            all_data[f"Dropdown 9 - {option}"] = extract_data_for_selection(8, option)
            if option == team1:
                break
        dropdown10 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[9])
        dropdown10_options = [option.text for option in dropdown10.options if option.text]
        for option in dropdown10_options:
            print(f"Selecting {option} in Dropdown 10")
            all_data[f"Dropdown 10 - {option}"] = extract_data_for_selection(9, option)
            if option == country2:
                break
        dropdown11 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[10])
        dropdown11_options = [option.text for option in dropdown11.options if option.text]
        for option in dropdown11_options:
            print(f"Selecting {option} in Dropdown 11")
            all_data[f"Dropdown 11 - {option}"] = extract_data_for_selection(10, option)
            if option == competition2:
                break
        dropdown12 = Select(driver.find_elements(By.CSS_SELECTOR, "select")[11])
        dropdown12_options = [option.text for option in dropdown12.options if option.text]
        for option in dropdown12_options:
            print(f"Selecting {option} in Dropdown 12")
            all_data[f"Dropdown 12 - {option}"] = extract_data_for_selection(11, option)
            if option == team2:
                break
        inputs = driver.find_elements(By.XPATH, "//input[@type='button' or @type='submit']")

        # Look for the specific button with text "Compare two teams" and click it
        for input_elem in inputs:
            input_text = input_elem.get_attribute("value")
            if input_text and "compare two teams" in input_text.lower():
                print(f"Found 'Compare two teams' button: {input_text}")
                input_elem.click()  # Click the button
                break
    else:
        pass

driver.quit()
