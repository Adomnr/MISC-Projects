from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Initialize the WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the target URL
url = "https://us.soccerway.com/teams/comparison/?team_ids%5B%5D=14495&team_ids%5B%5D=16210"
driver.get(url)

# Wait for the page to load completely
driver.implicitly_wait(10)

# Find all dropdown elements
dropdown_elements = driver.find_elements(By.TAG_NAME, "select")

# Print the number of dropdowns and their details
for index, dropdown in enumerate(dropdown_elements):
    print(f"Dropdown {index + 1}:")
    options = dropdown.find_elements(By.TAG_NAME, "option")
    for option in options:
        print(f" - {option.text}")

# Close the driver
driver.quit()
