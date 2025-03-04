from spire.doc import *
import os
from datetime import datetime

def document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                    Inverter_Watt, Inverter2_Watt, Number_of_Panels,
                    Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                    TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords,
                    Inverter_Name, Inverter2_Name,
                    Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces, BatterySpecs,
                    Number_of_inverters, panelwattage,
                    Carriage_Cost, Installation, Foundation, Earthing_val,
                    panelprice, structure_rate_normal, structure_rate_raised, InverterPrice,
                    pv_balance, Structure_TYP, advancepanelnames, advanceinverternames,advancecablename, NumberofDifferentInverters):
    # Create a Document object
    document = Document()

    # Load the template
    document.LoadFromFile(template_file)
    NOS = int(int(Number_of_Panels)/2)

    print("was here")
    if BatteryPrice != "" and BatteryPieces != "":
        BatteryPrice = int(BatteryPrice) * int(BatteryPieces)

    panelname = Name_of_Panels[:-6]

    TCN = ('{:,}'.format(int(TotalCostNormal)))
    TCR = ('{:,}'.format(int(TotalCostRaised)))
    TCNNI = ('{:,}'.format(int(TotalCostNormalNNI)))
    TCRNI = ('{:,}'.format(int(TotalCostRaisedNNI)))

    Net_Metering_cost = ('{:,}'.format(int(Net_Metering)))

    net_metering_symbol = ""

    if int(Net_Metering) == 0:
        net_metering_symbol = "N/A"
    else:
        net_metering_symbol = "01 Job"

    inverter_warranty_year = ""

    if str(Inverter_Name) == "FoxESS":
        inverter_warranty_year = "10 years replacement"
    else:
        inverter_warranty_year = "5 years"

    carriage_symbol = ""
    if int(Carriage_Cost) == 0:
        carriage_symbol = "N/A"
    else:
        carriage_symbol = "01 Job"

    foundation_symbol = ""
    if int(Foundation) == 0:
        foundation_symbol = "N/A"
    else:
        foundation_symbol = "01 Job"

    batteryprice = ""
    if BatteryPrice == "" or int(BatteryPrice) == 0:
        batteryprice = "N/A"
    else:
        batteryprice = str(BatteryPrice)

    if Installation == 0:
        installation_symbol = "N/A"
    else:
        installation_symbol = "01 Job"

    if Earthing_val == 0:
        earthing_symbol = "N/A"
    else:
        earthing_symbol = "01 Set"

    TotalPanelPrice = int(panelprice) * int(panelwattage) * int(Number_of_Panels)

    structure_rate = 0
    if Structure_TYP == "Normal":
        structure_rate = int(structure_rate_normal) * int((int(Number_of_Panels)/2))
    else:
        structure_rate = int(structure_rate_raised) * int(Number_of_Panels) * int(panelwattage)

    def convert_to_int_and_format(value):
        try:
            return '{:,}'.format(int(value))
        except ValueError:
            return None  # or handle the error as you wish

    # Perform conversion and formatting for each variable
    TotalPanelPrice = convert_to_int_and_format(TotalPanelPrice)
    structure_rate = convert_to_int_and_format(structure_rate)
    pv_balance = convert_to_int_and_format(pv_balance)
    Earthing_val = convert_to_int_and_format(Earthing_val)
    Installation = convert_to_int_and_format(Installation)
    Carriage_Cost = convert_to_int_and_format(Carriage_Cost)
    Foundation = convert_to_int_and_format(Foundation)
    InverterPrice = convert_to_int_and_format(InverterPrice)
    batteryprice = convert_to_int_and_format(batteryprice)

    # Execute statements only if conversion was successful
    if all(variable is not None for variable in
           [TotalPanelPrice, structure_rate, pv_balance, Earthing_val, Installation, Carriage_Cost, Foundation,
            InverterPrice, batteryprice]):
        # Execute statements here
        TotalPanelPrice = f'{TotalPanelPrice}'
        structure_rate = f'{structure_rate}'
        pv_balance = f'{pv_balance}'
        Earthing_val = f'{Earthing_val}'
        Installation = f'{Installation}'
        Carriage_Cost = f'{Carriage_Cost}'
        Foundation = f'{Foundation}'
        InverterPrice = f'{InverterPrice}'
        batteryprice = f'{batteryprice}'
    print("Structure Price: "+ structure_rate)

    # Store the placeholders and new strings in a dictionary
    dictionary = {
                    '[uid]': str(UniqueID),
                    '[sw]': str(SystemSize),
                    '[nop]': str(Number_of_Panels),
                    '[nos]': str(NOS),
                    '[noi]': str(Number_of_inverters),
                    '[iw]': str(Inverter_Watt),
                    '[iw2]': str(Inverter2_Watt),
                    '[iwy]': str(inverter_warranty_year),
                    '[calculated_price]': str(TCR),
                    '[calculated_price_normal]': str(TCN),
                    '[calculated_price_NNI]': str(TCRNI),
                    '[calculated_price_normal_NNI]': str(TCNNI),
                    '[value_in_words]': str(valueinwords),
                    '[bn]': str(BatteryName),
                    '[bp]': str(batteryprice),
                    '[bv]': str(BatteryPieces),
                    '[bs]': str(BatterySpecs),
                    '[nmc]': str(Net_Metering_cost),
                    '[nms]': str(net_metering_symbol),
                    '[pn]': str(panelname),
                    '[pw]': str(panelwattage),
                    '[in]': str(Inverter_Name),
                    '[in2]': str(Inverter2_Name),
                    '[ip]': str(InverterPrice),
                    '[fs]': str(foundation_symbol),
                    '[cs]': str(carriage_symbol),
                    '[es]': str(earthing_symbol),
                    '[is]': str(installation_symbol),
                    '[pp]': str(TotalPanelPrice),
                    '[sp]': str(structure_rate),
                    '[pvp]': str(pv_balance),
                    '[ep]': str(Earthing_val),
                    '[isp]': str(Installation),
                    '[cp]': str(Carriage_Cost),
                    '[fp]': str(Foundation),
                    '[apn]': str(advancepanelnames),
                    '[ain]': str(advanceinverternames),
                    '[acn]': str(advancecablename)
                }

    # Loop through the items in the dictionary
    for key, value in dictionary.items():
        # Replace a placeholder (key) with a new string (value)
        document.Replace(key, value, False, True)

    filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
        ClientLocation) + " Quotation" + str(UniqueID) + ".docx"

    current_date = datetime.now()

    # Define the directory structure
    directory_year_month = current_date.strftime("%B %Y")
    directory_day = current_date.strftime("%B %d")

    # Create the directories if they don't exist
    os.makedirs(os.path.join(directory_year_month, directory_day, "Word Files"), exist_ok=True)
    os.makedirs(os.path.join(directory_year_month, directory_day, "TradeMark Removed File"), exist_ok=True)
    os.makedirs(os.path.join(directory_year_month, directory_day, "PDF Files"), exist_ok=True)

    # Combine the directory paths
    folder_path = os.path.join(directory_year_month, directory_day)

    # Path to the new file
    file_path = os.path.join(folder_path, "Word Files", filename)

    print(file_path)

    # Save the resulting document
    document.SaveToFile(file_path, FileFormat.Docx2016)
    document.Close()