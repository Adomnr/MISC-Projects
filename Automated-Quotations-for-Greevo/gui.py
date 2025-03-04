# Importing Libraries
import tkinter
from tkinter import ttk
from tkinter import messagebox
from readwritegooglesheets import *
# For Value in words.
import inflect
from datetime import datetime
from doctest import *
import uuid
import time
from trademarkremover import *
import math
from pathlib import Path

Inverter_name, Inverter_wattage, Inverter2_name, Inverter2_wattage, Inverter3_name, Inverter3_wattage = [], [], [], [], [], []

Number_of_Inverter1, Number_of_Inverter2 = 0, 0

template_file = ""
InverterPrice, Inverter2Price = 0, 0

Number_of_Panels, panelprice, panelwattage = 0, 0, 0

TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI = 0, 0, 0, 0

UniqueID = 0
serialNumber = 0

valueinwords = ""

generalrow, generalrowentry, inverterselectionrow, inverterselectionrowentry, inverterrow, inverterrowentry, inverter2row, inverter2rowentry, inverter3row, inverter3rowentry = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9

panelrow, panelrowentry, batteryrow, batteryentryrow, cinrow, cinrowentry = 10, 11, 12, 13, 14, 15


def GetSerialNumber():
    global serialNumber
    index = []
    for x in Customer_Data_Sheet.col_values(1):
        index.append(x)
    serialNumber = int(index[-1])


def enter_data():
    global valueinwords
    accepted = accept_var.get()
    total_cost_calculator()
    UniqueID = generate_unique_id()
    if Quotation_type_combobox.get() == "General Net Metering Not Included" or Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
        if structure_type_combobox.get() == "Raised":
            valueinwords = convert_to_words(TotalCostRaised)
        else:
            valueinwords = convert_to_words(TotalCostNormal)
    if Quotation_type_combobox.get() == "General Net Metering Included" or Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
        if structure_type_combobox.get() == "Raised":
            valueinwords = convert_to_words(TotalCostRaisedNNI)
        else:
            valueinwords = convert_to_words(TotalCostNormalNNI)
    if accepted == "Accepted":
        # User info
        SystemSize = System_Size_combobox.get()
        ClientName = Client_Name_combobox.get()
        ClientLocation = Client_Location_combobox.get()
        ReferredBy = Reffered_combobox.get()
        if SystemSize and ClientName and ClientLocation and ReferredBy:
            Inverter_TYP = inverter_type_combobox.get()
            Inverter_Name = inverter_name_combobox.get()
            Inverter2_Name = inverter2_name_combobox.get()
            Inverter_Watt = inverter_wattage_combobox.get()
            Inverter2_Watt = inverter2_wattage_combobox.get()
            Number_of_Inverter1 = inverter_number_Entry.get()
            Number_of_Inverter2 = inverter2_number_Entry.get()
            Number_of_inverters = int(Number_of_Inverter1) + int(Number_of_Inverter2)
            if Inverter_TYP and Inverter_Name and Inverter_Watt:
                Name_of_Panels = panel_name_combobox.get()
                # Course info
                Structure_Type = structure_type_combobox.get()
                pv_balance = pv_balance_combobox.get()
                if Structure_Type and pv_balance:
                    Foundation = foundation_work_entry.get()
                    Carriage_Cost = carriage_entry.get()
                    Installation = installation_entry.get()
                    Net_Metering = net_metering_entry.get()
                    BatteryPrice = Battery_Price_Entry.get()
                    BatteryName = Battery_Name_Entry.get()
                    BatteryPieces = Number_of_Batteries_Entry.get()
                    PanelName = panel_name_combobox.get()
                    if Carriage_Cost and Installation and Net_Metering:
                        print("System Size: ", SystemSize, "Client Name: ", ClientName, "Client Location: ",
                              ClientLocation)
                        print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP, "Inverter Name: ",
                              Inverter_Name)
                        print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels, "Panel Price: ",
                              panelprice)
                        print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                        print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost, "Installation Cost: ",
                              Installation)
                        print("Net Metering: ", Net_Metering, "Template File", template_file, "Inverter Price",
                              InverterPrice)
                        print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ", TotalCostRaised,
                              "Unique ID: ", UniqueID)
                        print("------------------------------------------")

                        document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP, Inverter_Watt,
                                         Inverter2_Watt, Number_of_Panels,
                                         Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                         TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords, Inverter_Name,
                                         Inverter2_Name,
                                         Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces, Number_of_inverters,
                                         panelwattage,
                                         Carriage_Cost, Installation, Foundation)

                        filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                            ClientLocation) + " Quotation" + str(UniqueID)
                        tradeMarkRemover(filename)

                        record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy, Inverter_TYP,
                                    Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice, Number_of_Panels,
                                    Structure_Type, pv_balance, Carriage_Cost, Installation, Net_Metering,
                                    template_file, InverterPrice, TotalCostNormal, TotalCostRaised, TotalCostNormalNNI,
                                    TotalCostRaisedNNI)
                    else:
                        tkinter.messagebox.showwarning(title="Error",
                                                       message="Enter Carriage, Installation and Net Metering Cost")
                else:
                    tkinter.messagebox.showwarning(title="Error", message="Enter Structure Type and PV Balance.")
            else:
                tkinter.messagebox.showwarning(title="Error", message="Enter All box of Inverters.")
        else:
            tkinter.messagebox.showwarning(title="Error", message="Enter All boxes of Client Information.")
    else:
        tkinter.messagebox.showwarning(title="Error", message="You have not accepted the terms")


# This updates the template which is going to be subsituited the placeholder into GTGN Grid Tied General Normal

def round_up_to_nearest_thousand(number):
    if number % 1000 == 0:
        return number
    else:
        return ((number // 1000) + 1) * 1000


def total_cost_calculator(*args):
    global TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI
    pv_balance = pv_balance_combobox.get()
    Carriage_Cost = carriage_entry.get()
    Installation = installation_entry.get()
    Net_Metering = net_metering_entry.get()
    Foundation_Work = foundation_work_entry.get()
    TotalCostRaised = round_up_to_nearest_thousand(
        (int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) +
        (20 * int(Number_of_Panels) * int(panelwattage)) + (
                int(panelwattage) * int(panelprice) * int(Number_of_Panels)) +
        int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Net_Metering) + int(Foundation_Work))

    TotalCostNormal = round_up_to_nearest_thousand(
        (int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) +
        int((6500 * (int(Number_of_Panels) / 2))) + (
                int(panelwattage) * int(panelprice) * int(Number_of_Panels)) +
        int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Net_Metering) + int(Foundation_Work))

    TotalCostRaisedNNI = (
            round_up_to_nearest_thousand(int(InverterPrice) * int(Number_of_Inverter1)) + (
                int(Inverter2Price) * int(Number_of_Inverter2)) +
            (20 * int(Number_of_Panels) * int(panelwattage)) + (
                    int(panelwattage) * int(panelprice) * int(Number_of_Panels))
            + int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Foundation_Work))

    TotalCostNormalNNI = round_up_to_nearest_thousand(
        (int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) +
        int((6500 * (int(Number_of_Panels) / 2))) + (
                int(panelwattage) * int(panelprice) * int(Number_of_Panels)) +
        int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Foundation_Work))


def convert_to_words(number):
    p = inflect.engine()
    words = p.number_to_words(number)
    return words


def generate_unique_id():
    # Get current timestamp in seconds
    timestamp = int(time.time())
    # Generate a random component within the range of 3 digits
    random_component = uuid.uuid4().int % 1000  # 3-digit random component
    # Combine timestamp and random component to create a unique ID
    unique_id = int(f"{timestamp:03d}{random_component:03d}") % 10000000
    return unique_id


def record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy, Inverter_TYP, Inverter_Name,
                Inverter_Watt, Name_of_Panels, panelprice, Number_of_Panels, Structure_Type, pv_balance, Carriage_Cost,
                Installation, Net_Metering, template_file, InverterPrice, TotalCostNormal, TotalCostRaised,
                TotalCostNormalNNI, TotalCostRaisedNNI):
    serial_list = []
    current_date_time = datetime.now()
    current_date_time = str(current_date_time)
    for x in Customer_Data_Sheet.col_values(1):
        serial_list.append(x)
    SerialNumber = int(serial_list[-1]) + 1
    if SerialNumber == 0:
        SerialNumber = 1
    Customer_Data_Sheet.append_row([SerialNumber, current_date_time, UniqueID, SystemSize, ClientName, ClientLocation,
                                    ReferredBy, Inverter_TYP, Inverter_Name, Inverter_Watt, Name_of_Panels, panelprice,
                                    Number_of_Panels, Structure_Type, pv_balance, Carriage_Cost, Installation,
                                    Net_Metering,
                                    InverterPrice, TotalCostNormal, TotalCostRaised, TotalCostNormalNNI,
                                    TotalCostRaisedNNI])


def update_template_type(*args):
    home_dir = Path.home()
    print(home_dir)
    global template_file
    if inverter_selection_combobox.get() == "1":
        if inverter_type_combobox.get() == "Grid Tie":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGR_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGRNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGRSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGRNISPI_Template.docx")
        if inverter_type_combobox.get() == "Hybrid":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGNSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_Template.docx")
    if inverter_selection_combobox.get() == "2":
        if inverter_type_combobox.get() == "Grid Tie" and inverter2_type_combobox.get() == "Hybrid":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_WHI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_WHI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WHI_Template.docx")
        if inverter_type_combobox.get() == "Grid Tie" and inverter2_type_combobox.get() == "Grid Tie":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_WGTI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WGTI_Template.docx")
        if inverter_type_combobox.get() == "Hybrid" and inverter2_type_combobox.get() == "Grid Tie":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGNSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_WGTI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_WGTI_Template.docx")
        if inverter_type_combobox.get() == "Hybrid" and inverter2_type_combobox.get() == "Hybrid":
            if structure_type_combobox.get() == "Normal":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_WHI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_WHI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_WHI_Template.docx")
    print(template_file)


def inverter_selection(*args):
    battery_area_selection()
    inverter_name_combobox.set('')
    inverter_wattage_combobox.set('')
    if inverter_type_combobox.get() == "Grid Tie":
        Inverter_name = Grid_Tie_Inverter_Names.copy()
    else:
        Inverter_name = Hybrid_Inverter_Names.copy()
    inverter_name_combobox['values'] = Inverter_name


def inverter2_selection(*args):
    battery_area_selection()
    inverter2_name_combobox.set('')
    inverter2_wattage_combobox.set('')
    if inverter2_type_combobox.get() == "Grid Tie":
        Inverter2_name = Grid_Tie_Inverter_Names.copy()
    else:
        Inverter2_name = Hybrid_Inverter_Names.copy()
    inverter2_name_combobox['values'] = Inverter2_name


def inverter3_selection(*args):
    battery_area_selection()
    inverter3_name_combobox.set('')
    inverter3_wattage_combobox.set('')
    if inverter3_type_combobox.get() == "Grid Tie":
        Inverter3_name = Grid_Tie_Inverter_Names.copy()
    else:
        Inverter3_name = Hybrid_Inverter_Names.copy()
    inverter3_name_combobox['values'] = Inverter3_name


def update_panel(*args):
    global panelprice
    global panelwattage
    i = panel_name_combobox.current()
    panelprice = Solar_Panel_Price[i]
    print(panelprice)
    for x in Solar_Panel_Wattage:
        panelwattage = Solar_Panel_Wattage[i]
    PanelWattageInt = int(panelwattage)
    print(panelwattage)
    Panels_Numbers(PanelWattageInt)


def Panels_Numbers(panelwattage):
    global Number_of_Panels
    SZ = int(System_Size_combobox.get()) * 1000
    Number_of_Panels = SZ / int(panelwattage)
    Number_of_Panels = math.ceil(Number_of_Panels)
    print(Number_of_Panels)


def inverter_wattage_selection(*args):
    inverter_wattage_combobox.set('')
    new_list = []
    if inverter_type_combobox.get() == "Grid Tie":
        if inverter_name_combobox.get() != "":
            index = (inverter_name_combobox.current() + 1) * 2
            for x in Grid_Tied_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter_wattage_combobox.config(values=new_list)
    else:
        if inverter_name_combobox.get() != "":
            index = (inverter_name_combobox.current() + 1) * 2
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter_wattage_combobox.config(values=new_list)


def inverter2_wattage_selection(*args):
    inverter2_wattage_combobox.set('')
    new_list = []
    if inverter2_type_combobox.get() == "Grid Tie":
        if inverter2_name_combobox.get() != "":
            index = (inverter2_name_combobox.current() + 1) * 2
            for x in Grid_Tied_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter2_wattage_combobox.config(values=new_list)
    else:
        if inverter2_name_combobox.get() != "":
            index = (inverter2_name_combobox.current() + 1) * 2
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter2_wattage_combobox.config(values=new_list)


def inverter3_wattage_selection(*args):
    inverter3_wattage_combobox.set('')
    new_list = []
    if inverter3_type_combobox.get() == "Grid Tie":
        if inverter3_name_combobox.get() != "":
            index = (inverter3_name_combobox.current() + 1) * 2
            for x in Grid_Tied_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter3_wattage_combobox.config(values=new_list)
    else:
        if inverter3_name_combobox.get() != "":
            index = (inverter3_name_combobox.current() + 1) * 2
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                new_list.append(x)
            while ("" in new_list):
                new_list.remove("")
            new_list.pop(0)
            inverter3_wattage_combobox.config(values=new_list)


def inverter_price(*args):
    index = ((inverter_name_combobox.current() + 2) * 2) - 1
    list = []
    if inverter_type_combobox.get() == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
        list.pop(0)
        while ("" in list):
            list.remove("")
    else:
        if inverter_type_combobox.get() == "Hybrid":
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print(list)
    if inverter_wattage_combobox.current() >= 0 and inverter_wattage_combobox.current() <= 20:
        index2 = int(inverter_wattage_combobox.current())
        global InverterPrice
        InverterPrice = list[index2]
        print(InverterPrice)


def inverter2_price(*args):
    index = ((inverter2_name_combobox.current() + 2) * 2) - 1
    list = []
    if inverter2_type_combobox.get() == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
        list.pop(0)
        while ("" in list):
            list.remove("")
    else:
        if inverter2_type_combobox.get() == "Hybrid":
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print(list)
    if inverter2_wattage_combobox.current() >= 0 and inverter2_wattage_combobox.current() <= 20:
        index2 = int(inverter2_wattage_combobox.current())
        global Inverter2Price
        Inverter2Price = list[index2]
        print(Inverter2Price)


def inverter3_price(*args):
    index = ((inverter3_name_combobox.current() + 2) * 2) - 1
    list = []
    if inverter3_type_combobox.get() == "Grid Tie":
        for x in Grid_Tied_Inverters.row_values(int(index)):
            list.append(x)
        list.pop(0)
        while ("" in list):
            list.remove("")
    else:
        if inverter3_type_combobox.get() == "Hybrid":
            for x in Hybrid_Offgrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print(list)
    if inverter3_wattage_combobox.current() >= 0 and inverter3_wattage_combobox.current() <= 20:
        index2 = int(inverter3_wattage_combobox.current())
        global Inverter3Price
        Inverter3Price = list[index2]
        print(Inverter3Price)


def inverter_number_selection(*args):
    if inverter_selection_combobox.get() == '1':
        inverter2_frame.grid_remove()
        inverter3_frame.grid_remove()
    else:
        if inverter_selection_combobox.get() == '2':
            inverter2_frame.grid()
            if inverter3_frame.winfo_ismapped():
                inverter3_frame.grid_remove()
        else:
            if inverter2_frame != inverter2_frame.grid():
                inverter2_frame.grid()
            inverter3_frame.grid()


def battery_area_selection(*args):
    if inverter_type_combobox.get() == "Grid Tie" or inverter2_type_combobox.get() == "Grid Tie" or inverter3_type_combobox.get() == "Grid Tie":
        battery_frame.grid_remove()
    if inverter_type_combobox.get() == "Hybrid" or inverter2_type_combobox.get() == "Hybrid" or inverter3_type_combobox.get() == "Hybrid":
        if not battery_frame.winfo_ismapped():
            battery_frame.grid()


# Inverter_name,name_of_solar_panels,wattage_of_solar_panels
window = tkinter.Tk()
window.title("Quotation Automation")

frame = tkinter.Frame(window)
frame.pack()

# Saving User Info
user_info_frame = tkinter.LabelFrame(frame, text="Customer Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

System_Size_label = tkinter.Label(user_info_frame, text="System Size")
System_Size_combobox = ttk.Entry(user_info_frame)
System_Size_label.grid(row=generalrow, column=0)
System_Size_combobox.grid(row=generalrowentry, column=0)

Client_Name_label = tkinter.Label(user_info_frame, text="Client Name")
Client_Name_combobox = ttk.Entry(user_info_frame)
Client_Name_label.grid(row=generalrow, column=1)
Client_Name_combobox.grid(row=generalrowentry, column=1)

Client_Location_label = tkinter.Label(user_info_frame, text="Location")
Client_Location_combobox = ttk.Entry(user_info_frame)
Client_Location_label.grid(row=generalrow, column=2)
Client_Location_combobox.grid(row=generalrowentry, column=2)

Reffered_label = tkinter.Label(user_info_frame, text="Referred By")
Reffered_combobox = ttk.Combobox(user_info_frame,
                                 values=["Madam Rafia", "Engr Sajjad", "Engr Shaban", "Engr Abid", "Engr Ammar Butt",
                                         "Engr Ubaid", "Sir Nabeel", "Engr Osama"])
Reffered_label.grid(row=generalrow, column=3)
Reffered_combobox.grid(row=generalrowentry, column=3)

inverter_selection_frame = tkinter.LabelFrame(frame, text="DIS")
inverter_selection_frame.grid(row=inverterselectionrow, column=0, sticky="news", padx=20, pady=15)

tracker_inverters = tkinter.StringVar(inverter_selection_frame)

inverter_selection_label = tkinter.Label(inverter_selection_frame, text="Number of Different Inverters")
inverter_selection_combobox = ttk.Combobox(inverter_selection_frame, values=['1', '2', '3'],
                                           textvariable=tracker_inverters)
inverter_selection_label.grid(row=inverterselectionrow, column=0)
inverter_selection_combobox.grid(row=inverterselectionrowentry, column=0)

tracker_inverters.trace('w', inverter_number_selection)

inverter1_frame = tkinter.LabelFrame(frame, text="Inverter 1 Area")
inverter1_frame.grid(row=inverterrow, column=0, sticky="news", padx=20, pady=15)

sel = tkinter.StringVar(inverter1_frame)
sel3 = tkinter.StringVar(inverter1_frame)

sel5 = tkinter.StringVar(inverter1_frame)

inverter_type_label = tkinter.Label(inverter1_frame, text="Inverter 1 Type")
inverter_type_combobox = ttk.Combobox(inverter1_frame, values=["Grid Tie", "Hybrid"], textvariable=sel)
inverter_type_label.grid(row=inverterrow, column=0, padx=5)
inverter_type_combobox.grid(row=inverterrowentry, column=0, padx=5)

sel.trace('w', inverter_selection)

inverter_name_label = tkinter.Label(inverter1_frame, text="Inverter 1 Name")
inverter_name_combobox = ttk.Combobox(inverter1_frame, values=Inverter_name, textvariable=sel3)
inverter_name_label.grid(row=inverterrow, column=1, padx=5)
inverter_name_combobox.grid(row=inverterrowentry, column=1, padx=5)

inverter_wattage_label = tkinter.Label(inverter1_frame, text=" Inverter 1 Wattage")
inverter_wattage_combobox = ttk.Combobox(inverter1_frame, values=Inverter_wattage, textvariable=sel5)
inverter_wattage_label.grid(row=inverterrow, column=2, padx=5)
inverter_wattage_combobox.grid(row=inverterrowentry, column=2, padx=5)

sel3.trace('w', inverter_wattage_selection)

inverter_number_Label = tkinter.Label(inverter1_frame, text="No of Type 1 Inverters")
inverter_number_Entry = ttk.Entry(inverter1_frame)
inverter_number_Label.grid(row=inverterrow, column=3, padx=5)
inverter_number_Entry.grid(row=inverterrowentry, column=3, padx=5)

if inverter_number_Entry.get() == "":
    inverter_number_Entry.insert(0, '1')

inverter2_frame = tkinter.LabelFrame(frame, text="Inverter 2 Area")
inverter2_frame.grid(row=inverter2row, column=0, sticky="news", padx=20, pady=15)

sel6 = tkinter.StringVar(inverter2_frame)
sel7 = tkinter.StringVar(inverter2_frame)
sel8 = tkinter.StringVar(inverter2_frame)

inverter2_type_label = tkinter.Label(inverter2_frame, text="Inverter 2 Type")
inverter2_type_combobox = ttk.Combobox(inverter2_frame, values=["Grid Tie", "Hybrid"], textvariable=sel6)
inverter2_type_label.grid(row=inverter2row, column=0, padx=5)
inverter2_type_combobox.grid(row=inverter2rowentry, column=0, padx=5)

sel6.trace('w', inverter2_selection)

inverter2_name_label = tkinter.Label(inverter2_frame, text="Inverter 2 Name")
inverter2_name_combobox = ttk.Combobox(inverter2_frame, values=Inverter2_name, textvariable=sel7)
inverter2_name_label.grid(row=inverter2row, column=1, padx=5)
inverter2_name_combobox.grid(row=inverter2rowentry, column=1, padx=5)

inverter2_wattage_label = tkinter.Label(inverter2_frame, text=" Inverter 2 Wattage")
inverter2_wattage_combobox = ttk.Combobox(inverter2_frame, values=Inverter2_wattage, textvariable=sel8)
inverter2_wattage_label.grid(row=inverter2row, column=2, padx=5)
inverter2_wattage_combobox.grid(row=inverter2rowentry, column=2, padx=5)

inverter2_number_Label = tkinter.Label(inverter2_frame, text="No of Type 2 Inverters")
inverter2_number_Entry = ttk.Entry(inverter2_frame)
inverter2_number_Label.grid(row=inverter2row, column=3, padx=5)
inverter2_number_Entry.grid(row=inverter2rowentry, column=3, padx=5)

if inverter2_number_Entry.get() == "":
    inverter2_number_Entry.insert(0, '1')

sel7.trace('w', inverter2_wattage_selection)

inverter3_frame = tkinter.LabelFrame(frame, text="Inverter 3 Area")
inverter3_frame.grid(row=inverter3row, column=0, sticky="news", padx=20, pady=15)

sel9 = tkinter.StringVar(inverter3_frame)
sel10 = tkinter.StringVar(inverter3_frame)
sel11 = tkinter.StringVar(inverter3_frame)

inverter3_type_label = tkinter.Label(inverter3_frame, text="Inverter 3 Type")
inverter3_type_combobox = ttk.Combobox(inverter3_frame, values=["Grid Tie", "Hybrid"], textvariable=sel9)
inverter3_type_label.grid(row=inverter3row, column=0, padx=5)
inverter3_type_combobox.grid(row=inverter3rowentry, column=0, padx=5)

sel9.trace('w', inverter3_selection)

inverter3_name_label = tkinter.Label(inverter3_frame, text="Inverter 3 Name")
inverter3_name_combobox = ttk.Combobox(inverter3_frame, values=Inverter2_name, textvariable=sel10)
inverter3_name_label.grid(row=inverter3row, column=1, padx=5)
inverter3_name_combobox.grid(row=inverter3rowentry, column=1, padx=5)

inverter3_wattage_label = tkinter.Label(inverter3_frame, text=" Inverter 3 Wattage")
inverter3_wattage_combobox = ttk.Combobox(inverter3_frame, values=Inverter2_wattage, textvariable=sel11)
inverter3_wattage_label.grid(row=inverter3row, column=2, padx=5)
inverter3_wattage_combobox.grid(row=inverter3rowentry, column=2, padx=5)

inverter3_number_Label = tkinter.Label(inverter3_frame, text="No of Type 3 Inverters")
inverter3_number_Entry = ttk.Entry(inverter3_frame)
inverter3_number_Label.grid(row=inverter3row, column=3, padx=5)
inverter3_number_Entry.grid(row=inverter3rowentry, column=3, padx=5)

if inverter3_number_Entry.get() == "":
    inverter3_number_Entry.insert(0, '1')

sel10.trace('w', inverter2_wattage_selection)

if inverter_selection_combobox.get() == "":
    inverter_selection_combobox.set('1')

panel_frame = tkinter.LabelFrame(frame, text="Panel Area")
panel_frame.grid(row=panelrow, column=0, sticky="news", padx=20, pady=15)

sel2 = tkinter.StringVar(panel_frame)

panel_name_label = tkinter.Label(panel_frame, text="Panel Name")
panel_name_combobox = ttk.Combobox(panel_frame, values=Solar_Panels_Names, textvariable=sel2)
panel_name_label.grid(row=panelrow, column=0, padx=5)
panel_name_combobox.grid(row=panelrowentry, column=0, padx=5)

sel5.trace('w', inverter_price)
sel8.trace('w', inverter2_price)
sel11.trace('w', inverter3_price)
sel2.trace('w', update_panel)
sel4 = tkinter.StringVar(panel_frame)

structure_type = tkinter.Label(panel_frame, text="Structure Type")
structure_type_combobox = ttk.Combobox(panel_frame, values=["Normal", "Raised"], textvariable=sel4)
structure_type.grid(row=panelrow, column=2, padx=5)
structure_type_combobox.grid(row=panelrowentry, column=2, padx=5)

if structure_type_combobox.get() == "":
    structure_type_combobox.set("Normal")

pv_balance_label = tkinter.Label(panel_frame, text="PV Balance")
pv_balance_combobox = ttk.Entry(panel_frame)
pv_balance_label.grid(row=panelrow, column=1, padx=5)
pv_balance_combobox.grid(row=panelrowentry, column=1, padx=5)

Quotation_type = tkinter.Label(panel_frame, text="Quotation Type")
Quotation_type_combobox = ttk.Combobox(panel_frame,
                                       values=["General Net Metering Included", "Specify Brand Net Metering Included",
                                               "General Net Metering Not Included",
                                               "Specify Brand Net Metering Not Included"])
Quotation_type.grid(row=panelrow, column=3, padx=5)
Quotation_type_combobox.grid(row=panelrowentry, column=3, padx=5)

if Quotation_type_combobox.get() == "":
    Quotation_type_combobox.set("Net Metering Not Included")

sel4.trace('w', update_template_type)

battery_frame = tkinter.LabelFrame(frame, text="Battery Area")
battery_frame.grid(row=batteryrow, column=0, sticky="news", padx=20, pady=15)

Battery_Name_Label = tkinter.Label(battery_frame, text="Battery Name")
Battery_Name_Entry = ttk.Entry(battery_frame)
Battery_Name_Label.grid(row=batteryrow, column=0, padx=5)
Battery_Name_Entry.grid(row=batteryentryrow, column=0, padx=5)

if Battery_Name_Entry.get() == "":
    Battery_Name_Entry.insert(0, "Daewoo Deep Cycle")

Battery_Price_Label = tkinter.Label(battery_frame, text="Battery Price")
Battery_Price_Entry = ttk.Entry(battery_frame)
Battery_Price_Label.grid(row=batteryrow, column=1, padx=5)
Battery_Price_Entry.grid(row=batteryentryrow, column=1, padx=5)

if Battery_Price_Entry.get() == "":
    Battery_Price_Entry.insert(0, "180000")

Number_of_Batteries_Label = tkinter.Label(battery_frame, text="Number of Batteries")
Number_of_Batteries_Entry = ttk.Entry(battery_frame)
Number_of_Batteries_Label.grid(row=batteryrow, column=2, padx=5)
Number_of_Batteries_Entry.grid(row=batteryentryrow, column=2, padx=5)

if Number_of_Batteries_Entry.get() == "":
    Number_of_Batteries_Entry.insert(0, '4')

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving Course Info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=cinrow, column=0, sticky="news", padx=20, pady=10)

carriage = tkinter.Label(courses_frame, text="Carriage")
carriage_entry = ttk.Entry(courses_frame)
carriage.grid(row=cinrow, column=0)
carriage_entry.grid(row=cinrowentry, column=0)

if carriage_entry.get() == "":
    carriage_entry.insert(0, '0')

installation = tkinter.Label(courses_frame, text="Installation")
installation_entry = ttk.Entry(courses_frame)
installation.grid(row=cinrow, column=1)
installation_entry.grid(row=cinrowentry, column=1)

if installation_entry.get() == "":
    installation_entry.insert(0, '0')

net_metering = tkinter.Label(courses_frame, text="Net Metering")
net_metering_entry = ttk.Entry(courses_frame)
net_metering.grid(row=cinrow, column=2)
net_metering_entry.grid(row=cinrowentry, column=2)

if net_metering_entry.get() == "":
    net_metering_entry.insert(0, '0')

foundation_work = tkinter.Label(courses_frame, text="Foundation Work")
foundation_work_entry = ttk.Entry(courses_frame)
foundation_work.grid(row=cinrow, column=3)
foundation_work_entry.grid(row=cinrowentry, column=3)

if foundation_work_entry.get() == "":
    foundation_work_entry.insert(0, "0")

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=12, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Last Check")
terms_frame.grid(row=16, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="Checked All.", variable=accept_var, onvalue="Accepted",
                                  offvalue="Not Accepted")
terms_check.grid(row=17, column=0)

# Button
button = tkinter.Button(frame, text="Enter data", command=enter_data)
button.grid(row=18, column=0, sticky="news", padx=20, pady=10)

window.mainloop()
