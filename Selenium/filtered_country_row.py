import pandas as pd

# List of countries
countries = [
    "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", "Armenia", "Aruba",
    "Australia", "Austria", "Azerbaijan", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium",
    "Belize", "Benin", "Bermuda", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Brazil",
    "British Virgin Islands", "Brunei Darussalam", "Bulgaria", "Burkina Faso", "Burundi", "Cambodia",
    "Cameroon", "Canada", "Cape Verde", "Cayman Islands", "Chile", "China PR", "Chinese Taipei",
    "Colombia", "Congo DR", "Congo", "Cook Islands", "Costa Rica", "Croatia", "Cuba", "Curaçao",
    "Cyprus", "Czechia", "Côte d'Ivoire", "Denmark", "Djibouti", "Dominican Republic", "Ecuador",
    "Egypt", "El Salvador", "England", "Estonia", "Eswatini", "Ethiopia", "Faroe Islands", "Fiji",
    "Finland", "France", "French Guiana", "Gabon", "Gambia", "Georgia", "Germany", "Ghana",
    "Gibraltar", "Greece", "Grenada", "Guadeloupe", "Guam", "Guatemala", "Guinea", "Guyana", "Haiti",
    "Honduras", "Hong Kong, China", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq",
    "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Korea Republic", "Kosovo",
    "Kuwait", "Kyrgyz Republic", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya",
    "Lithuania", "Luxembourg", "Macao", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali",
    "Malta", "Martinique", "Mauritania", "Mauritius", "Mexico", "Moldova", "Mongolia", "Montenegro",
    "Morocco", "Mozambique", "Myanmar", "Nepal", "Netherlands", "New Caledonia", "New Zealand",
    "Nicaragua", "Nigeria", "North Macedonia", "Northern Ireland", "Norway", "Oman", "Pakistan",
    "Palestine", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", "Portugal",
    "Puerto Rico", "Qatar", "Republic of Ireland", "Romania", "Russia", "Rwanda", "Réunion",
    "San Marino", "Saudi Arabia", "Scotland", "Senegal", "Serbia", "Sierra Leone", "Singapore",
    "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "Spain", "Sri Lanka",
    "St. Kitts and Nevis", "Sudan", "Suriname", "Sweden", "Switzerland", "Syria", "São Tomé e Príncipe",
    "Tahiti", "Tajikistan", "Tanzania", "Thailand", "Togo", "Trinidad and Tobago", "Tunisia",
    "Turkmenistan", "Turks and Caicos Islands", "Türkiye", "USA", "Uganda", "Ukraine",
    "United Arab Emirates", "Uruguay", "Uzbekistan", "Venezuela", "Vietnam", "Wales", "Yemen",
    "Zambia", "Zimbabwe"
]

file_path = 'dropdown_options_filtered_club.xlsx'  # Update this with the correct file path

# Read the Excel file
df = pd.read_excel(file_path, header=None)

# Drop rows with all NaN values
df.dropna(how='all', inplace=True)

# Prepare formatted rows with row numbers
formatted_rows = []

# Iterate over each row and collect the data
for index, row in df.iterrows():
    if row[0] in countries:
        row_data = [str(cell) for cell in row if pd.notna(cell)]
        formatted_row = f"Row {index + 1}: " + ", ".join(row_data)
        formatted_rows.append(formatted_row)

# Print the formatted rows
for row in formatted_rows:
    print(row)
