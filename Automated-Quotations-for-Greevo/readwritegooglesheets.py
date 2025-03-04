# Library of gspread and oauth2client for the api
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Scopes Required for the running of the Google Docs api
scopes = [
    'https://www.googleapis.com/auth/spreadsheets'
    'https://www.googleapis.com/auth/drive'
]

# File which contains secret key should be in Data Folder
creds = ServiceAccountCredentials.from_json_keyfile_name("./Data/secret_key.json")

# Authorizing the api with secret key file
file = gspread.authorize(creds)

# Importing all three sheets with Hybrid Off grid and Grid Tied and Also Solar Panels Prices

Greevo_Data = file.open("Greevo Data")


# Importing sheet 1 of all of them
Hybrid_Inverters = Greevo_Data.worksheet("Hybrid Inverters Rates")
Grid_Tied_Inverters = Greevo_Data.worksheet("Grid Tie Inverters Rates")
Solar_Panels = Greevo_Data.worksheet("Solar Panels Rates")
Customer_Data_Sheet = Greevo_Data.worksheet("Customer Data")

# Importing Inverter Names with separated values
Grid_Tie_name_of_inverters = Grid_Tied_Inverters.col_values(1)

# Importing name wattage and price of hybrid off grid inverters
Hybrid_name_of_inverters = Hybrid_Inverters.col_values(1)

# Importing name wattage and price of Solar Panels
Solar_Panels_Names = Solar_Panels.col_values(1)
Solar_Panel_Price = Solar_Panels.col_values(2)
Solar_Panel_Wattage = Solar_Panels.col_values(3)

# clear selection of name and price and wattage of the inverters and solar panels.
Grid_Tie_Inverter_Names = Grid_Tie_name_of_inverters[1::2]

# clear selection of name and price and wattage of the inverters and solar panels.
Hybrid_Inverter_Names = Hybrid_name_of_inverters[1::2]
