from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl

# Initialize the WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the target URL
url = "https://us.soccerway.com/teams/comparison/?competition_ids%5B%5D=1269&team_ids%5B%5D=3414&competition_ids%5B%5D=953&team_ids%5B%5D=113"
driver.get(url)

# Wait for the page to load completely
driver.implicitly_wait(10)

# Extract all table elements
tables = driver.find_elements(By.TAG_NAME, "table")

# Print the number of tables found
print(f"Number of tables found: {len(tables)}")

# Prepare data storage
extracted_data = []

# Check if the 15th table exists
if len(tables) >= 15:
    # Extract data from the 15th table
    table = tables[14]  # Index 14 for the 15th table
    print("\nTable 15:")
    rows = table.find_elements(By.CSS_SELECTOR, "tr")

    # Specify the row indices we are interested in (0-based index)
    row_indices = [3, 10, 11, 12, 15]  # Row 4, 11, 12, 13, and 16

    for index in row_indices:
        if index < len(rows):
            row = rows[index]
            data = row.find_elements(By.CSS_SELECTOR, "td")
            data_text = [d.text for d in data]
            extracted_data.append(data_text)
            print(f"Row {index + 1}: {data_text}")
        else:
            print(f"Row {index + 1} does not exist.")
else:
    print("The 15th table does not exist on this page.")

# Close the driver
driver.quit()

# Load an existing workbook
workbook_path = "extracted_data.xlsx"  # Replace with your actual workbook path
workbook = openpyxl.load_workbook(workbook_path)
sheet = workbook.active  # Replace with your actual sheet name if needed

# Write the extracted data to the Excel sheet starting from row 2 and column 2
start_row = 4
start_col = 4

for i, data in enumerate(extracted_data):
    for j, value in enumerate(data):
        try:
            # Try converting to integer first
            numeric_value = int(value)
        except ValueError:
            try:
                # If integer conversion fails, try converting to float
                numeric_value = float(value.replace('m', ''))  # Handle minute notation
            except ValueError:
                # If both conversions fail, keep it as a string
                numeric_value = value
        cell = sheet.cell(row=start_row + i, column=start_col + j)
        cell.value = numeric_value  # Only change the value, retain the existing format

# Save the updated workbook
workbook.save(workbook_path)
print(f"Data has been updated in {workbook_path}")
