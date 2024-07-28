from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os

# Bypass SSL verification
os.environ['WDM_SSL_VERIFY'] = '0'

# Initialize the WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the target URL
url = "https://us.soccerway.com/matches/2024/07/23/korea-republic/k-league/seongnam-ilhwa/chunnam-dragons/4308862/head2head/"
driver.get(url)

# Wait for the page to load completely
driver.implicitly_wait(10)

# Find all input elements with type="button" or type="submit"
inputs = driver.find_elements(By.XPATH, "//input[@type='button' or @type='submit']")

# Look for the specific button with text "Compare two teams"
compare_button_found = False
for input_elem in inputs:
    input_text = input_elem.get_attribute("value")
    if input_text and "compare two teams" in input_text.lower():
        print(f"Found 'Compare two teams' button: {input_text}")
        compare_button_found = True
        break

if not compare_button_found:
    print("'Compare two teams' button not found.")

# Close the WebDriver
driver.quit()
