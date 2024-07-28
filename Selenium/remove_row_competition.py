import openpyxl

# List of countries to check
countries_to_check = [
    "Select competition"
]

# Load the workbook and select the active sheet
wb = openpyxl.load_workbook("dropdown_options_filtered.xlsx")
ws = wb.active

# Function to check if all countries are present in a row
def row_contains_all_countries(row):
    cell_values = [cell.value for cell in row]
    return all(country in cell_values for country in countries_to_check)

# Iterate over the rows and remove rows that contain all the countries
rows_to_remove = []
for row in ws.iter_rows():
    if row_contains_all_countries(row):
        rows_to_remove.append(row[0].row)  # Store the row number

# Remove the rows in reverse order to prevent shifting issues
for row_num in sorted(rows_to_remove, reverse=True):
    ws.delete_rows(row_num)

# Save the modified workbook
wb.save("dropdown_options_filtered_2.xlsx")
