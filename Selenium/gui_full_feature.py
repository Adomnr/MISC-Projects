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

file_path_club = 'dropdown_options_filtered_club.xlsx'  # Update this with the correct file path for club
file_path_national = 'dropdown_options_national.xlsx'  # Update this with the correct file path for national

# Function to read and process Excel data for club
def read_excel_data(file_path):
    df = pd.read_excel(file_path, header=None)
    df.dropna(how='all', inplace=True)
    data_dict = {}
    current_country = None
    is_competition_row = False
    for index, row in df.iterrows():
        row_data = [str(cell) for cell in row if pd.notna(cell)]
        if row[0] in countries:
            current_country = row[0]
            is_competition_row = True
            data_dict[current_country] = {'competitions': [], 'teams': {}}
        elif current_country:
            if is_competition_row:
                competition = row[0]
                data_dict[current_country]['competitions'].append(competition)
                data_dict[current_country]['teams'][competition] = []
                is_competition_row = False
            else:
                data_dict[current_country]['teams'][competition].extend(row_data)
                is_competition_row = True
    return data_dict

# Function to read and process Excel data for national
def read_excel_data_national(file_path):
    df = pd.read_excel(file_path, header=None)
    df.dropna(how='all', inplace=True)
    data_dict = {}
    current_country = None
    for index, row in df.iterrows():
        row_data = [str(cell) for cell in row if pd.notna(cell)]
        if row[0] in countries:
            current_country = row[0]
            data_dict[current_country] = []
        elif current_country:
            data_dict[current_country].extend(row_data)
    print(f"National Data Dict: {data_dict}")
    return data_dict

# Read club and national data
data_dict_club = read_excel_data(file_path_club)
data_dict_national = read_excel_data_national(file_path_national)

# GUI Creation
def update_competitions(event, dropdown_box, dropdown_comp, data_dict):
    selected_country = dropdown_box.get()
    competitions = data_dict[selected_country]['competitions']
    dropdown_comp['values'] = competitions

def update_teams(event, dropdown_comp, dropdown_team, country_dropdown, data_dict):
    selected_competition = dropdown_comp.get()
    selected_country = country_dropdown.get()
    teams = data_dict[selected_country]['teams'][selected_competition]
    dropdown_team['values'] = teams

def update_ui(event):
    selected_type = dropdown_1.get()
    if selected_type == "National":
        dropdown_6.pack_forget()
        dropdown_7.pack_forget()
        dropdown_2.set('')
        dropdown_3.set('')
        dropdown_4.set('')
        dropdown_5.set('')
        dropdown_2.bind("<<ComboboxSelected>>", lambda event: update_teams_national(event, dropdown_2, dropdown_4, data_dict_national))
        dropdown_3.bind("<<ComboboxSelected>>", lambda event: update_teams_national(event, dropdown_3, dropdown_5, data_dict_national))
    else:
        dropdown_6.pack()
        dropdown_7.pack()
        dropdown_2.set('')
        dropdown_3.set('')
        dropdown_4.set('')
        dropdown_5.set('')
        dropdown_2.bind("<<ComboboxSelected>>", lambda event: update_competitions(event, dropdown_2, dropdown_4, data_dict_club))
        dropdown_3.bind("<<ComboboxSelected>>", lambda event: update_competitions(event, dropdown_3, dropdown_5, data_dict_club))
        dropdown_4.bind("<<ComboboxSelected>>", lambda event: update_teams(event, dropdown_4, dropdown_6, dropdown_2, data_dict_club))
        dropdown_5.bind("<<ComboboxSelected>>", lambda event: update_teams(event, dropdown_5, dropdown_7, dropdown_3, data_dict_club))

def update_teams_national(event, country_dropdown, dropdown_comp, data_dict):
    selected_country = country_dropdown.get()
    print(f"Selected Country: {selected_country}")
    if selected_country in data_dict:
        teams = data_dict[selected_country]  # Get the row under the country
        print(f"Teams for {selected_country}: {teams}")
        dropdown_comp['values'] = teams
    else:
        dropdown_comp['values'] = []

# Create the main window
root = tk.Tk()
root.title("Football Teams Comparison")

# Dropdown Box 1
label_1 = tk.Label(root, text="Select Type")
label_1.pack()
dropdown_1 = ttk.Combobox(root, values=["Club", "National"])
dropdown_1.pack()
dropdown_1.bind("<<ComboboxSelected>>", update_ui)

# Dropdown Box 2
label_2 = tk.Label(root, text="Country 1")
label_2.pack()
dropdown_2 = ttk.Combobox(root, values=countries)
dropdown_2.pack()

# Dropdown Box 3
label_3 = tk.Label(root, text="Country 2")
label_3.pack()
dropdown_3 = ttk.Combobox(root, values=countries)
dropdown_3.pack()

# Dropdown Box 4
label_4 = tk.Label(root, text="Competition 1 / Team 1")
label_4.pack()
dropdown_4 = ttk.Combobox(root)
dropdown_4.pack()

# Dropdown Box 5
label_5 = tk.Label(root, text="Competition 2 / Team 2")
label_5.pack()
dropdown_5 = ttk.Combobox(root)
dropdown_5.pack()

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

# Compare button
compare_button = tk.Button(root, text="Compare two teams")
compare_button.pack()

root.mainloop()
