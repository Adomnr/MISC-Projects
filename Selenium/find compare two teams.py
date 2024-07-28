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

# Function to print button details
def print_button_details(elements, element_type):
    for idx, element in enumerate(elements):
        element_text = element.text.strip()
        if not element_text:  # If no text, check for value attribute
            element_text = element.get_attribute("value")
        if not element_text:  # If no value, check for aria-label attribute
            element_text = element.get_attribute("aria-label")
        if not element_text:  # If no aria-label, check for title attribute
            element_text = element.get_attribute("title")
        print(f"{element_type} {idx + 1}: '{element_text}'")

# Find all button, input[type=button], input[type=submit], and anchor elements that act as buttons
buttons = driver.find_elements(By.TAG_NAME, "button")
inputs = driver.find_elements(By.XPATH, "//input[@type='button' or @type='submit']")
anchors = driver.find_elements(By.TAG_NAME, "a")

# Print the details of all found elements
print("Buttons:")
print_button_details(buttons, "Button")

print("\nInput elements acting as buttons:")
print_button_details(inputs, "Input")

print("\nAnchor elements acting as buttons:")
print_button_details(anchors, "Anchor")

# Close the WebDriver
driver.quit()
