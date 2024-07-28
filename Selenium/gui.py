import pandas as pd
import tkinter as tk
from tkinter import ttk

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

# Data dictionary to store competitions and teams
data_dict = {}

# Variable to keep track of the current country and processing state
current_country = None
is_competition_row = False

# Iterate over each row and collect the data
for index, row in df.iterrows():
    row_data = [str(cell) for cell in row if pd.notna(cell)]

    if row[0] in countries:
        # New country detected
        current_country = row[0]
        is_competition_row = True
        data_dict[current_country] = {'competitions': [], 'teams': {}}
    elif current_country:
        if is_competition_row:
            # Row is a competition
            competition = row[0]
            data_dict[current_country]['competitions'].append(competition)
            data_dict[current_country]['teams'][competition] = []
            is_competition_row = False
        else:
            # Row is a team
            data_dict[current_country]['teams'][competition].extend(row_data)
            is_competition_row = True

# GUI Creation
def update_competitions(event, dropdown_box, dropdown_comp):
    selected_country = dropdown_box.get()
    competitions = data_dict[selected_country]['competitions']
    dropdown_comp['values'] = competitions

def update_teams(event, dropdown_comp, dropdown_team, country_dropdown):
    selected_competition = dropdown_comp.get()
    selected_country = country_dropdown.get()
    teams = data_dict[selected_country]['teams'][selected_competition]
    dropdown_team['values'] = teams

# Create the main window
root = tk.Tk()
root.title("Football Teams Comparison")

# Dropdown Box 1
label_1 = tk.Label(root, text="Select Type")
label_1.pack()
dropdown_1 = ttk.Combobox(root, values=["Club", "National"])
dropdown_1.pack()

# Dropdown Box 2
label_2 = tk.Label(root, text="Country 1")
label_2.pack()
dropdown_2 = ttk.Combobox(root, values=countries)
dropdown_2.pack()
dropdown_2.bind("<<ComboboxSelected>>", lambda event: update_competitions(event, dropdown_2, dropdown_4))

# Dropdown Box 3
label_3 = tk.Label(root, text="Country 2")
label_3.pack()
dropdown_3 = ttk.Combobox(root, values=countries)
dropdown_3.pack()
dropdown_3.bind("<<ComboboxSelected>>", lambda event: update_competitions(event, dropdown_3, dropdown_5))

# Dropdown Box 4
label_4 = tk.Label(root, text="Competition 1")
label_4.pack()
dropdown_4 = ttk.Combobox(root)
dropdown_4.pack()
dropdown_4.bind("<<ComboboxSelected>>", lambda event: update_teams(event, dropdown_4, dropdown_6, dropdown_2))

# Dropdown Box 5
label_5 = tk.Label(root, text="Competition 2")
label_5.pack()
dropdown_5 = ttk.Combobox(root)
dropdown_5.pack()
dropdown_5.bind("<<ComboboxSelected>>", lambda event: update_teams(event, dropdown_5, dropdown_7, dropdown_3))

# Dropdown Box 6
label_6 = tk.Label(root, text="Team 1")
label_6.pack()
dropdown_6 = ttk.Combobox(root)
dropdown_6.pack()

# Dropdown Box 7
label_7 = tk.Label(root, text="Team 2")
label_7.pack()
dropdown_7 = ttk.Combobox(root)
dropdown_7.pack()

# Compare Button
compare_button = tk.Button(root, text="Compare two teams")
compare_button.pack()

root.mainloop()
