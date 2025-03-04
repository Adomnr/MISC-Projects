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
from spire.doc import *
import os
from pathlib import Path
from docx import Document
from docx2pdf import convert

Inverter_name, Inverter_wattage, Inverter2_name, Inverter2_wattage, Inverter3_name, Inverter3_wattage = [], [], [], [], [], []

Number_of_Inverter1, Number_of_Inverter2 = 0, 0

template_file = ""
InverterPrice, Inverter2Price = 0, 0

Number_of_Panels, panelprice, panelwattage, panel_structure_rate = 0, 0, 0, 0

TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI = 0, 0, 0, 0

TotalCostNormalAmend, TotalCostRaisedAmend, TotalCostNormalNNIAmend, TotalCostRaisedNNIAmend = 0, 0, 0, 0

advancepanelnames, advanceinverternames, advancecablename = "", "", "AC Single Core 6mm Cable"

UniqueID = 0
serialNumber = 0
Exist_Unique_ID = 0

structure_rate_normal, structure_rate_raised = 6500, 20

valueinwords = ""

generalrow, generalrowentry, inverterselectionrow, inverterselectionrowentry, inverterrow, inverterrowentry, inverter2row, inverter2rowentry, inverter3row, inverter3rowentry = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9

panelrow, panelrowentry, batteryrow, batteryentryrow, cinrow, cinrowentry = 10, 11, 12, 13, 14, 15


def capitalize_first_character_in_each_word(sentence):
    capitalized_sentence = sentence.title()
    return capitalized_sentence


def GetSerialNumber():
    global serialNumber
    index = []
    for x in Customer_Data_Sheet.col_values(1):
        index.append(x)
    serialNumber = int(index[-1])


def itemised_exist(*args):
    existwindow = tkinter.Tk()
    existwindow.title("Enter Unique ID")

    existframe = tkinter.Frame(existwindow)
    existframe.pack()

    Exist_Frame = tkinter.LabelFrame(existframe, text="Unique ID Submission")
    Exist_Frame.grid(row=0, column=0, padx=40, pady=40)

    UniqueID_Label = tkinter.Label(existframe, text="Unique ID")
    UniqueID_Entry = tkinter.Entry(existframe)
    UniqueID_Label.grid(row=0, column=0, padx=20, pady=10)
    UniqueID_Entry.grid(row=1, column=0, padx=20, pady=10)

    def update_window(*args):
        global Exist_Unique_ID
        index = 0
        id_list = Customer_Data_Sheet.col_values(3)
        id_list.pop(0)
        print(id_list)
        for x in id_list:
            print(x)
            if int(UniqueID_Entry.get()) == int(x):
                Exist_Unique_ID = int(x)
                existwindow.destroy()
                exist_unique_id()
                index = 1
        if index == 0:
            tkinter.messagebox.showinfo(title="Error", message="Unique ID not Found")
        else:
            tkinter.messagebox.showinfo(title="Success", message="Unique ID Found")

    button = tkinter.Button(existframe, text="Enter UID", command=update_window)
    button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

def replace_row(sheet, row_number, row_data):
    range_to_update = f"A{row_number}:AL{row_number}"  # Assuming data range spans columns A to Z, adjust as needed
    sheet.update(range_name=range_to_update, values=[row_data])

def trace_id():
    global Exist_Unique_ID
    print(Exist_Unique_ID)
    index = 2
    id_list = Customer_Data_Sheet.col_values(3)
    id_list.pop(0)
    for x in id_list:
        print("NOice" + x)
        if int(Exist_Unique_ID) == int(x):
            break
        index += 1
    return index


def exist_unique_id():
    row = trace_id()
    print(row)
    row_data = Customer_Data_Sheet.row_values(row)
    print(row_data)
    windowID = tkinter.Tk()
    windowID.title("Unique ID: " + str(row_data[2]))

    frame = tkinter.Frame(windowID)
    frame.pack()

    # Saving User Info
    user_info_frame = tkinter.LabelFrame(frame, text="Customer Information")
    user_info_frame.grid(row=0, column=0, padx=20, pady=10)

    System_Size_label = tkinter.Label(user_info_frame, text="System Size")
    System_Size_Entry = ttk.Entry(user_info_frame)
    System_Size_label.grid(row=generalrow, column=0)
    System_Size_Entry.grid(row=generalrowentry, column=0)

    if System_Size_Entry.get() == "":
        System_Size_Entry.insert(0, str(row_data[3]))

    Client_Name_label = tkinter.Label(user_info_frame, text="Client Name")
    Client_Name_Entry = ttk.Entry(user_info_frame)
    Client_Name_label.grid(row=generalrow, column=1)
    Client_Name_Entry.grid(row=generalrowentry, column=1)

    if Client_Name_Entry.get() == "":
        Client_Name_Entry.insert(0, str(row_data[4]))

    Client_Location_label = tkinter.Label(user_info_frame, text="Location")
    Client_Location_Entry = ttk.Entry(user_info_frame)
    Client_Location_label.grid(row=generalrow, column=2)
    Client_Location_Entry.grid(row=generalrowentry, column=2)

    if Client_Location_Entry.get() == "":
        Client_Location_Entry.insert(0, str(row_data[5]))

    Reffered_label = tkinter.Label(user_info_frame, text="Referred By")
    Reffered_Entry = ttk.Entry(user_info_frame)
    Reffered_label.grid(row=generalrow, column=3)
    Reffered_Entry.grid(row=generalrowentry, column=3)

    if Reffered_Entry.get() == "":
        Reffered_Entry.insert(0, str(row_data[6]))

    inverter_selection_frame = tkinter.LabelFrame(frame, text="DIS")
    inverter_selection_frame.grid(row=inverterselectionrow, column=0, sticky="news", padx=20, pady=15)

    foundation_work = tkinter.Label(inverter_selection_frame, text="Foundation Work")
    foundation_work_entry = ttk.Entry(inverter_selection_frame)
    foundation_work.grid(row=inverterselectionrow, column=1)
    foundation_work_entry.grid(row=inverterselectionrowentry, column=1)

    if foundation_work_entry.get() == "":
        foundation_work_entry.insert(0, str(row_data[24]))

    Quotation_type = tkinter.Label(inverter_selection_frame, text="Quotation Type")
    Quotation_type_Entry = ttk.Combobox(inverter_selection_frame,
                                           values=["General Net Metering Included",
                                                   "Specify Brand Net Metering Included",
                                                   "General Net Metering Not Included",
                                                   "Specify Brand Net Metering Not Included",
                                                   "Itemised General Net Metering Included",
                                                   "Itemised Specify Brand Net Metering Included",
                                                   "Itemised General Net Metering Not Included",
                                                   "Itemised Specify Brand Net Metering Not Included"],
                                           )
    Quotation_type.grid(row=inverterselectionrow, column=2, padx=5)
    Quotation_type_Entry.grid(row=inverterselectionrowentry, column=2, padx=5)

    inverter1_frame = tkinter.LabelFrame(frame, text="Inverter 1 Area")
    inverter1_frame.grid(row=inverterrow, column=0, sticky="news", padx=20, pady=15)

    inverter_type_label = tkinter.Label(inverter1_frame, text="Inverter 1 Type")
    inverter_type_Entry = ttk.Entry(inverter1_frame)
    inverter_type_label.grid(row=inverterrow, column=0, padx=5)
    inverter_type_Entry.grid(row=inverterrowentry, column=0, padx=5)

    if inverter_type_Entry.get() == "":
        inverter_type_Entry.insert(0, str(row_data[7]))

    inverter_name_label = tkinter.Label(inverter1_frame, text="Inverter 1 Name")
    inverter_name_Entry = ttk.Entry(inverter1_frame)
    inverter_name_label.grid(row=inverterrow, column=1, padx=5)
    inverter_name_Entry.grid(row=inverterrowentry, column=1, padx=5)

    if inverter_name_Entry.get() == "":
        inverter_name_Entry.insert(0, str(row_data[8]))

    inverter_wattage_label = tkinter.Label(inverter1_frame, text=" Inverter 1 Wattage")
    inverter_wattage_Entry = ttk.Entry(inverter1_frame)
    inverter_wattage_label.grid(row=inverterrow, column=2, padx=5)
    inverter_wattage_Entry.grid(row=inverterrowentry, column=2, padx=5)

    if inverter_wattage_Entry.get() == "":
        inverter_wattage_Entry.insert(0, str(row_data[9]))

    inverter1_price_label = tkinter.Label(inverter1_frame, text="Inverter 1 Price")
    inverter1_price_Entry = ttk.Entry(inverter1_frame)
    inverter1_price_label.grid(row=inverterrow, column=3, padx=5)
    inverter1_price_Entry.grid(row=inverterrowentry, column=3, padx=5)

    if inverter1_price_Entry.get() == "":
        inverter1_price_Entry.insert(0, str(row_data[10]))

    inverter_number_Label = tkinter.Label(inverter1_frame, text="No of Type 1 Inverters")
    inverter_number_Entry = ttk.Entry(inverter1_frame)
    inverter_number_Label.grid(row=inverterrow, column=4, padx=5)
    inverter_number_Entry.grid(row=inverterrowentry, column=4, padx=5)

    if inverter_number_Entry.get() == "":
        inverter_number_Entry.insert(0, str(row_data[11]))

    inverter2_frame = tkinter.LabelFrame(frame, text="Inverter 2 Area")
    inverter2_frame.grid(row=inverter2row, column=0, sticky="news", padx=20, pady=15)

    inverter2_type_label = tkinter.Label(inverter2_frame, text="Inverter 2 Type")
    inverter2_type_Entry = ttk.Entry(inverter2_frame)
    inverter2_type_label.grid(row=inverter2row, column=0, padx=5)
    inverter2_type_Entry.grid(row=inverter2rowentry, column=0, padx=5)

    if inverter2_type_Entry.get() == "":
        inverter2_type_Entry.insert(0, str(row_data[12]))

    inverter2_name_label = tkinter.Label(inverter2_frame, text="Inverter 2 Name")
    inverter2_name_Entry = ttk.Entry(inverter2_frame)
    inverter2_name_label.grid(row=inverter2row, column=1, padx=5)
    inverter2_name_Entry.grid(row=inverter2rowentry, column=1, padx=5)

    if inverter2_name_Entry.get() == "":
        inverter2_name_Entry.insert(0, str(row_data[13]))

    inverter2_wattage_label = tkinter.Label(inverter2_frame, text=" Inverter 2 Wattage")
    inverter2_wattage_Entry = ttk.Entry(inverter2_frame)
    inverter2_wattage_label.grid(row=inverter2row, column=2, padx=5)
    inverter2_wattage_Entry.grid(row=inverter2rowentry, column=2, padx=5)

    if inverter2_wattage_Entry.get() == "":
        inverter2_wattage_Entry.insert(0, str(row_data[14]))

    inverter2_price_label = tkinter.Label(inverter2_frame, text="Inverter 2 Price")
    inverter2_price_Entry = ttk.Entry(inverter2_frame)
    inverter2_price_label.grid(row=inverter2row, column=3, padx=5)
    inverter2_price_Entry.grid(row=inverter2rowentry, column=3, padx=5)

    if inverter2_price_Entry.get() == "":
        inverter2_price_Entry.insert(0, str(row_data[15]))

    inverter2_number_Label = tkinter.Label(inverter2_frame, text="No of Type 2 Inverters")
    inverter2_number_Entry = ttk.Entry(inverter2_frame)
    inverter2_number_Label.grid(row=inverter2row, column=4, padx=5)
    inverter2_number_Entry.grid(row=inverter2rowentry, column=4, padx=5)

    if inverter2_number_Entry.get() == "":
        inverter2_number_Entry.insert(0, str(row_data[16]))

    panel_frame = tkinter.LabelFrame(frame, text="Panel Area")
    panel_frame.grid(row=panelrow, column=0, sticky="news", padx=20, pady=15)

    panel_name_label = tkinter.Label(panel_frame, text="Panel Name")
    panel_name_Entry = ttk.Entry(panel_frame)
    panel_name_label.grid(row=panelrow, column=0, padx=5)
    panel_name_Entry.grid(row=panelrowentry, column=0, padx=5)

    if panel_name_Entry.get() == "":
        panel_name_Entry.insert(0, str(row_data[17]))

    panel_price_label = tkinter.Label(panel_frame, text="Panel Price Per Watt")
    panel_price_entry = tkinter.Entry(panel_frame)
    panel_price_label.grid(row=panelrow, column=1, padx=5)
    panel_price_entry.grid(row=panelrowentry, column=1, padx=5)

    if panel_price_entry.get() == "":
        panel_price_entry.insert(0, str(row_data[18]))

    Numberofpanels_label = tkinter.Label(panel_frame, text="Number of Panels")
    Numberofpanels_entry = tkinter.Entry(panel_frame)
    Numberofpanels_entry.grid(row=panelrow, column=2, padx=5)
    Numberofpanels_entry.grid(row=panelrowentry, column=2, padx=5)

    if Numberofpanels_entry.get() == "":
        Numberofpanels_entry.insert(0, str(row_data[19]))

    pv_balance_label = tkinter.Label(panel_frame, text="PV Balance")
    pv_balance_Entry = ttk.Entry(panel_frame)
    pv_balance_label.grid(row=panelrow, column=3, padx=5)
    pv_balance_Entry.grid(row=panelrowentry, column=3, padx=5)

    if pv_balance_Entry.get() == "":
        pv_balance_Entry.insert(0, str(row_data[20]))

    structure_type = tkinter.Label(panel_frame, text="Structure Type")
    structure_type_Entry = ttk.Entry(panel_frame)
    structure_type.grid(row=panelrow, column=4, padx=5)
    structure_type_Entry.grid(row=panelrowentry, column=4, padx=5)

    if structure_type_Entry.get() == "":
        structure_type_Entry.insert(0, str(row_data[26]))

    battery_frame = tkinter.LabelFrame(frame, text="Battery Area")
    battery_frame.grid(row=batteryrow, column=0, sticky="news", padx=20, pady=15)

    Battery_Name_Label = tkinter.Label(battery_frame, text="Battery Name")
    Battery_Name_Entry = ttk.Entry(battery_frame)
    Battery_Name_Label.grid(row=batteryrow, column=0, padx=5)
    Battery_Name_Entry.grid(row=batteryentryrow, column=0, padx=5)

    if Battery_Name_Entry.get() == "":
        Battery_Name_Entry.insert(0, str(row_data[30]))

    Battery_Price_Label = tkinter.Label(battery_frame, text="Battery Price")
    Battery_Price_Entry = ttk.Entry(battery_frame)
    Battery_Price_Label.grid(row=batteryrow, column=1, padx=5)
    Battery_Price_Entry.grid(row=batteryentryrow, column=1, padx=5)

    if Battery_Price_Entry.get() == "":
        Battery_Price_Entry.insert(0, str(row_data[29]))

    Number_of_Batteries_Label = tkinter.Label(battery_frame, text="Number of Batteries")
    Number_of_Batteries_Entry = ttk.Entry(battery_frame)
    Number_of_Batteries_Label.grid(row=batteryrow, column=2, padx=5)
    Number_of_Batteries_Entry.grid(row=batteryentryrow, column=2, padx=5)

    Battery_Specification_Label = tkinter.Label(battery_frame, text="Specifications of Battery")
    Battery_Specification_Entry = tkinter.Entry(battery_frame)
    Battery_Specification_Label.grid(row=batteryrow, column=3, padx=10)
    Battery_Specification_Entry.grid(row=batteryentryrow, column=3, padx=10)

    if Number_of_Batteries_Entry.get() == "":
        Number_of_Batteries_Entry.insert(0, str(row_data[31]))

    if Battery_Specification_Entry.get() == "":
        Battery_Specification_Entry.insert(0, str(row_data[32]))

    for widget in user_info_frame.winfo_children():
        widget.grid_configure(padx=10, pady=5)

    # Saving Course Info
    courses_frame = tkinter.LabelFrame(frame)
    courses_frame.grid(row=cinrow, column=0, sticky="news", padx=20, pady=10)

    carriage = tkinter.Label(courses_frame, text="Carriage")
    carriage_entry = ttk.Entry(courses_frame)
    carriage.grid(row=cinrow, column=0, padx=5)
    carriage_entry.grid(row=cinrowentry, column=0, padx=5)

    if carriage_entry.get() == "":
        carriage_entry.insert(0, str(row_data[21]))

    installation = tkinter.Label(courses_frame, text="Installation")
    installation_entry = ttk.Entry(courses_frame)
    installation.grid(row=cinrow, column=1, padx=5)
    installation_entry.grid(row=cinrowentry, column=1, padx=5)

    if installation_entry.get() == "":
        installation_entry.insert(0, str(row_data[22]))

    net_metering = tkinter.Label(courses_frame, text="Net Metering")
    net_metering_entry = ttk.Entry(courses_frame)
    net_metering.grid(row=cinrow, column=2, padx=5)
    net_metering_entry.grid(row=cinrowentry, column=2, padx=5)

    if net_metering_entry.get() == "":
        net_metering_entry.insert(0, str(row_data[23]))

    Earthing = tkinter.Label(courses_frame, text="Earthing")
    Earthing_entry = ttk.Entry(courses_frame)
    Earthing.grid(row=cinrow, column=3, padx=5)
    Earthing_entry.grid(row=cinrowentry, column=3, padx=5)

    if Earthing_entry.get() == "":
        Earthing_entry.insert(0, str(row_data[25]))

    def total_cost(*args):
        global TotalCostNormalAmend, TotalCostRaisedAmend, TotalCostNormalNNIAmend, TotalCostRaisedNNIAmend

        index = 1
        panelamendwattage = ""

        for x in Solar_Panels_Names:
            if panel_name_Entry.get() == x:
                panelamendwattage = Solar_Panel_Wattage[index]
            index += 1

        TotalCostRaisedAmend = round_up_to_nearest_thousand(
                            (int(inverter1_price_Entry.get()) * int(inverter_number_Entry.get())) + (int(inverter2_price_Entry.get()) * int(inverter2_number_Entry.get())) +
                            (int(structure_rate_raised) * int(Numberofpanels_entry.get()) * int(panelamendwattage)) + (int(panelamendwattage) * int(panel_price_entry.get()) *
                            int(Numberofpanels_entry.get())) + int(pv_balance_Entry.get()) + int(carriage_entry.get()) + int(installation_entry.get()) + int(net_metering_entry.get()) + int(foundation_work_entry.get()))

        print(TotalCostRaisedAmend)

        TotalCostRaisedNNIAmend = round_up_to_nearest_thousand((int(inverter1_price_Entry.get()) * int(inverter_number_Entry.get())) + (int(inverter2_price_Entry.get()) *
                                                        int(inverter2_number_Entry.get())) + (int(structure_rate_raised) * int(Numberofpanels_entry.get()) *
                                                        int(panelamendwattage)) + (int(panelamendwattage) * int(panel_price_entry.get()) * int(Numberofpanels_entry.get())) +
                                                        int(pv_balance_Entry.get()) + int(carriage_entry.get()) + int(installation_entry.get()) + int(foundation_work_entry.get()))
        print(TotalCostRaisedNNIAmend)
        TotalCostNormalAmend = round_up_to_nearest_thousand( (int(inverter1_price_Entry.get()) * int(inverter_number_Entry.get())) + (int(inverter2_price_Entry.get()) * int(inverter2_number_Entry.get())) +
                            int((int(structure_rate_normal) * (int(Numberofpanels_entry.get()) / 2))) + (int(panelamendwattage) * int(panel_price_entry.get()) * int(Numberofpanels_entry.get())) +
                            int(pv_balance_Entry.get()) + int(carriage_entry.get()) + int(installation_entry.get()) + int(net_metering_entry.get()) + int(foundation_work_entry.get()))

        print(TotalCostNormalAmend)

        TotalCostNormalNNIAmend = round_up_to_nearest_thousand((int(inverter1_price_Entry.get()) * int(inverter_number_Entry.get())) + (int(inverter2_price_Entry.get()) * int(inverter2_number_Entry.get())) +
                                                          int( (int(structure_rate_normal) * (int(Numberofpanels_entry.get()) / 2))) + (int(panelamendwattage) * int(panel_price_entry.get()) *
                                                        int(Numberofpanels_entry.get())) + int(carriage_entry.get()) + int( int(installation_entry.get()) + int(foundation_work_entry.get())))
        print(TotalCostNormalNNIAmend)

    def make_quotation():
        template = get_template_type(row_data[37], structure_type_Entry.get(), Quotation_type_Entry.get(), inverter_type_Entry.get(), inverter2_type_Entry.get())
        valueinwordsamend = ""
        if Quotation_type_Entry.get() == "General Net Metering Not Included" or Quotation_type_Entry.get() == "Specify Brand Net Metering Not Included":
            if structure_type_Entry.get() == "Raised":
                valueinwordsamend = capitalize_first_character_in_each_word(convert_to_words(TotalCostRaised))
            else:
                valueinwordsamend = capitalize_first_character_in_each_word(convert_to_words(TotalCostNormal))
        if Quotation_type_Entry.get() == "General Net Metering Included" or Quotation_type_Entry.get() == "Specify Brand Net Metering Included":
            if structure_type_Entry.get() == "Raised":
                valueinwordsamend = capitalize_first_character_in_each_word(convert_to_words(TotalCostRaisedNNI))
            else:
                valueinwordsamend = capitalize_first_character_in_each_word(convert_to_words(TotalCostNormalNNI))
        data = [row_data[0],  row_data[1], row_data[2], System_Size_Entry.get(), Client_Name_Entry.get(),Client_Location_Entry.get(),
                Reffered_Entry.get(), inverter_type_Entry.get(), inverter_name_Entry.get(), inverter_wattage_Entry.get(),
                inverter1_price_Entry.get(), inverter_number_Entry.get(), inverter2_type_Entry.get(), inverter2_name_Entry.get(),
                inverter2_wattage_Entry.get(), inverter2_price_Entry.get(), inverter2_number_Entry.get(), panel_name_Entry.get(),
                panel_price_entry.get(), Numberofpanels_entry.get(), pv_balance_Entry.get(), carriage_entry.get(), installation_entry.get(),
                net_metering_entry.get(), foundation_work_entry.get(), Earthing_entry.get(), structure_type_Entry.get(), template, Quotation_type_Entry.get(),
                Battery_Price_Entry.get(), Battery_Name_Entry.get(), Number_of_Batteries_Entry.get(), Battery_Specification_Entry.get()]

        Number_of_inverters = int(data[11])+int(data[16])

        index = 1
        panelamendwattage = 0
        for x in Solar_Panels_Names:
            if data[17] == x:
                panelamendwattage = Solar_Panel_Wattage[index]
            index += 1
        total_cost()
        document_creater(data[2], data[4], data[5], data[3], data[7], data[9], data[14], data[19],
                         data[23], template, TotalCostNormalAmend, TotalCostRaisedAmend,
                         TotalCostNormalNNIAmend, TotalCostRaisedNNIAmend, valueinwordsamend,
                         data[8],data[13], data[17], data[29], data[30], data[31], data[32],
                         str(Number_of_inverters), panelamendwattage, data[21], data[22], data[24], data[25],
                         data[18], structure_rate_normal, structure_rate_raised, data[10],
                         data[20], data[20], advancepanelnames, advanceinverternames, advancecablename,
                         row_data[37])

        filename = str(data[3]) + "kW " + str(data[7]) + " " + str(
            data[5]) + " Quotation" + str(data[2]) + ".docx"
        tradeMarkRemover(filename)

        record_data(data[3], data[2], data[4], data[5], data[6],
                data[7], data[8], data[9], data[10], data[11],
                data[12], data[13], data[14], data[15], data[16],
                data[17], data[18], data[19], data[20], data[21],
                data[22], data[23], data[24], data[25], data[26], data[27],
                data[28], data[29], data[30], data[31], data[32],
                TotalCostNormalAmend, TotalCostRaisedAmend, TotalCostNormalNNIAmend, TotalCostRaisedNNIAmend, row_data[37])
        tkinter.messagebox.showinfo(title="Success",
                                    message="File is Created and Recorded\nFileName: " + filename)
        data = [row_data[0], row_data[1], row_data[2], System_Size_Entry.get(), Client_Name_Entry.get(),
                Client_Location_Entry.get(),
                Reffered_Entry.get(), inverter_type_Entry.get(), inverter_name_Entry.get(),
                inverter_wattage_Entry.get(), inverter1_price_Entry.get(), inverter_number_Entry.get(), inverter2_type_Entry.get(),
                inverter2_name_Entry.get(), inverter2_wattage_Entry.get(), inverter2_price_Entry.get(), inverter2_number_Entry.get(),
                panel_name_Entry.get(), panel_price_entry.get(), Numberofpanels_entry.get(), pv_balance_Entry.get(), carriage_entry.get(),
                installation_entry.get(), net_metering_entry.get(), foundation_work_entry.get(), Earthing_entry.get(), structure_type_Entry.get(),
                row_data[27], row_data[28], Battery_Price_Entry.get(), Battery_Name_Entry.get(), Number_of_Batteries_Entry.get(),
                Battery_Specification_Entry.get(), TotalCostNormalAmend, TotalCostRaisedAmend, TotalCostNormalNNIAmend, TotalCostRaisedNNIAmend]
        print(data)
        replace_row(Customer_Data_Sheet, row, data)
        windowID.destroy()

    def amend_option():
        print(row)
        total_cost()
        data = [row_data[0], row_data[1], row_data[2], System_Size_Entry.get(), Client_Name_Entry.get(), Client_Location_Entry.get(),
                Reffered_Entry.get(), inverter_type_Entry.get(), inverter_name_Entry.get(), inverter_wattage_Entry.get(),
                inverter1_price_Entry.get(), inverter_number_Entry.get(), inverter2_type_Entry.get(), inverter2_name_Entry.get(),
                inverter2_wattage_Entry.get(), inverter2_price_Entry.get(), inverter2_number_Entry.get(), panel_name_Entry.get(),
                panel_price_entry.get(), Numberofpanels_entry.get(), pv_balance_Entry.get(), carriage_entry.get(), installation_entry.get(),
                net_metering_entry.get(), foundation_work_entry.get(), Earthing_entry.get(), structure_type_Entry.get(), row_data[27], row_data[28],
                Battery_Price_Entry.get(), Battery_Name_Entry.get(), Number_of_Batteries_Entry.get(), Battery_Specification_Entry.get(),TotalCostNormalAmend,
                TotalCostRaisedAmend,TotalCostNormalNNIAmend,TotalCostRaisedNNIAmend,row_data[37]]
        print(data)
        replace_row(Customer_Data_Sheet,row,data)
        windowID.destroy()


    amend_button = tkinter.Button(frame, text="Amend", command=amend_option)
    amend_button.grid(row=18, column=0, sticky="news", padx=10, pady=10)


    make_button = tkinter.Button(frame, text="Create Quotation", command=make_quotation)
    make_button.grid(row=19, column=0, sticky="news", padx=10, pady=10)

def override_rates(*args):
    newwindow = tkinter.Tk()
    newwindow.title("Rates Overrider")

    newframe = tkinter.Frame(newwindow)
    newframe.pack()

    # Saving User Info
    Rates_Frames = tkinter.LabelFrame(newframe, text="Inverter and Panel Rates")
    Rates_Frames.grid(row=0, column=0, padx=20, pady=10)

    Inverter_Rate_Label = tkinter.Label(Rates_Frames, text="Inverter 1 Rate")
    Inverter_Rate_Entry = tkinter.Entry(Rates_Frames)
    Inverter_Rate_Label.grid(row=0, column=0, padx=20, pady=10)
    Inverter_Rate_Entry.grid(row=1, column=0, padx=20, pady=10)

    if Inverter_Rate_Entry.get() == "":
        Inverter_Rate_Entry.insert(0, str(InverterPrice))

    PanelRateLabel = tkinter.Label(Rates_Frames, text="Panel Rate")
    PanelRateValue = tkinter.Entry(Rates_Frames)
    PanelRateLabel.grid(row=2, column=0, padx=20, pady=10)
    PanelRateValue.grid(row=3, column=0, padx=20, pady=10)

    if PanelRateValue.get() == "":
        PanelRateValue.insert(0, str(panelprice))

    structure_rate_normal_label = tkinter.Label(Rates_Frames, text="Normal Structure Rate")
    structure_rate_normal_entry = tkinter.Entry(Rates_Frames)
    structure_rate_normal_label.grid(row=4, column=0, padx=20, pady=10)
    structure_rate_normal_entry.grid(row=5, column=0, padx=20, pady=10)

    if structure_rate_normal_entry.get() == "":
        structure_rate_normal_entry.insert(0, "6500")

    structure_rate_raised_label = tkinter.Label(Rates_Frames, text="Raised Structure Rate")
    structure_rate_raised_entry = tkinter.Entry(Rates_Frames)
    structure_rate_raised_label.grid(row=6, column=0, padx=20, pady=10)
    structure_rate_raised_entry.grid(row=7, column=0, padx=20, pady=10)

    if structure_rate_raised_entry.get() == "":
        structure_rate_raised_entry.insert(0, "20")

    def update_rates(*args):
        update_panel_price(PanelRateValue.get())
        update_inverter_price(Inverter_Rate_Entry.get())
        update_structure_rate_normal_rates(structure_rate_normal_entry.get())
        update_structure_rate_raised_rates(structure_rate_raised_entry.get())
        print(panelprice)
        print(InverterPrice)
        print(structure_rate_raised)
        print(structure_rate_normal)
        newwindow.destroy()

    ChangeButton = tkinter.Button(newframe, text="Override", command=update_rates)
    ChangeButton.grid(row=4, column=0)


def update_structure_rate_normal_rates(structurepricenormal):
    global structure_rate_normal
    structure_rate_normal = int(structurepricenormal)


def update_structure_rate_raised_rates(structurepriceraised):
    global structure_rate_raised
    structure_rate_raised = int(structurepriceraised)


def update_panel_price(PanelRateValue):
    global panelprice
    panelprice = int(PanelRateValue)


def update_inverter_price(Inverter_Rate_Entry):
    global InverterPrice
    InverterPrice = int(Inverter_Rate_Entry)


def advance_options(*args):
    new2window = tkinter.Tk()
    new2window.title("Rates Overrider")

    new2frame = tkinter.Frame(new2window)
    new2frame.pack()

    # Saving User Info
    Rates_Frames = tkinter.LabelFrame(new2frame, text="Inverter and Panel General Names")
    Rates_Frames.grid(row=0, column=0, padx=20, pady=10)

    Inverter_Name_Label = tkinter.Label(Rates_Frames, text="Names of Inverters")
    Inverter_Name_Entry = tkinter.Entry(Rates_Frames)
    Inverter_Name_Label.grid(row=0, column=0, padx=20, pady=10)
    Inverter_Name_Entry.grid(row=1, column=0, padx=20, pady=10)

    if Inverter_Name_Entry.get() == "":
        Inverter_Name_Entry.insert(0, str(advanceinverternames))

    PanelNameLabel = tkinter.Label(Rates_Frames, text="Names of Panels")
    PanelNameValue = tkinter.Entry(Rates_Frames)
    PanelNameLabel.grid(row=2, column=0, padx=20, pady=10)
    PanelNameValue.grid(row=3, column=0, padx=20, pady=10)

    if PanelNameValue.get() == "":
        PanelNameValue.insert(0, str(advancepanelnames))

    CableNameLabel = tkinter.Label(Rates_Frames, text="Names of AC Cable")
    CableNameValue = tkinter.Entry(Rates_Frames)
    CableNameLabel.grid(row=4, column=0, padx=20, pady=10)
    CableNameValue.grid(row=5, column=0, padx=20, pady=10)

    if CableNameValue.get() == "":
        CableNameValue.insert(0, str(advancecablename))

    def update_names(*args):
        update_panel_names(PanelNameValue.get())
        update_inverters_name(Inverter_Name_Entry.get())
        print(advancepanelnames)
        print(advanceinverternames)
        new2window.destroy()

    ChangeButton = tkinter.Button(new2frame, text="Apply Option", command=update_names)
    ChangeButton.grid(row=4, column=0)


def update_panel_names(PanelNamesValue):
    global advancepanelnames
    advancepanelnames = PanelNamesValue


def update_inverters_name(InvertersNameValue):
    global advanceinverternames
    advanceinverternames = InvertersNameValue


def update_cable_name(CableName):
    global advancecablename
    advancecablename = CableName


def enter_data():
    global valueinwords
    total_cost_calculator()
    UniqueID = generate_unique_id()
    if Quotation_type_combobox.get() == "General Net Metering Not Included" or Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
        if structure_type_combobox.get() == "Raised":
            valueinwords = capitalize_first_character_in_each_word(convert_to_words(TotalCostRaised))
        else:
            valueinwords = capitalize_first_character_in_each_word(convert_to_words(TotalCostNormal))
    if Quotation_type_combobox.get() == "General Net Metering Included" or Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
        if structure_type_combobox.get() == "Raised":
            valueinwords = capitalize_first_character_in_each_word(convert_to_words(TotalCostRaisedNNI))
        else:
            valueinwords = capitalize_first_character_in_each_word(convert_to_words(TotalCostNormalNNI))
    print(valueinwords)
    if True:
        # User info
        SystemSize = System_Size_combobox.get()
        ClientName = Client_Name_combobox.get()
        ClientLocation = Client_Location_combobox.get()
        ReferredBy = Reffered_combobox.get()

        if SystemSize and ClientName and ClientLocation and ReferredBy:
            NumberofDifferentInverters = inverter_selection_combobox.get()
            Inverter_TYP = inverter_type_combobox.get()
            Inverter2_TYP = inverter2_type_combobox.get()
            Inverter_Name = inverter_name_combobox.get()
            Inverter2_Name = inverter2_name_combobox.get()
            Inverter_Watt = inverter_wattage_combobox.get()
            Inverter2_Watt = inverter2_wattage_combobox.get()
            Number_of_Inverter1 = inverter_number_Entry.get()
            Number_of_Inverter2 = inverter2_number_Entry.get()
            Number_of_inverters = int(Number_of_Inverter1) + int(Number_of_Inverter2)
            Quotation_TYP = Quotation_type_combobox.get()

            if Inverter_TYP and Inverter_Name and Inverter_Watt:

                if inverter_selection_combobox.get() == "2":

                    if Inverter2_TYP and Inverter2_Name and Inverter2_Watt:
                        Name_of_Panels = panel_name_combobox.get()
                        Structure_Type = structure_type_combobox.get()
                        pv_balance = pv_balance_combobox.get()

                        if Name_of_Panels and Structure_Type and pv_balance:
                            BatteryPrice = Battery_Price_Entry.get()
                            BatteryName = Battery_Name_Entry.get()
                            BatteryPieces = Number_of_Batteries_Entry.get()
                            BatterySpecs = Battery_Specification_Entry.get()

                            if Inverter_TYP == "Hybrid" or Inverter2_TYP == "Hybrid":
                                if BatteryPrice and BatteryName and BatteryPieces:
                                    Earthing_val = Earthing_entry.get()
                                    Foundation = foundation_work_entry.get()
                                    Carriage_Cost = carriage_entry.get()
                                    Installation = installation_entry.get()
                                    Net_Metering = net_metering_entry.get()

                                    if Carriage_Cost and Installation and Net_Metering and Foundation and Earthing_val:
                                        print("10")
                                        print("System Size: ", SystemSize, "Client Name: ", ClientName,
                                              "Client Location: ", ClientLocation)
                                        print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP,
                                              "Inverter Name: ", Inverter_Name)
                                        print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels,
                                              "Panel Price: ", panelprice)
                                        print("Inverter 1 Price", InverterPrice, "Inverter 2 Price", Inverter2Price)
                                        print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                                        print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost,
                                              "Installation Cost: ", Installation)
                                        print("Net Metering: ", Net_Metering, "Template File", template_file,
                                              "Inverter Price", InverterPrice)
                                        print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ",
                                              TotalCostRaised, "Unique ID: ", UniqueID)
                                        print("------------------------------------------")

                                        document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                                                         Inverter_Watt, Inverter2_Watt, Number_of_Panels,
                                                         Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                                         TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords,
                                                         Inverter_Name, Inverter2_Name, Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces,
                                                         BatterySpecs, Number_of_inverters, panelwattage,
                                                         Carriage_Cost, Installation, Foundation, Earthing_val, panelprice,
                                                         structure_rate_normal, structure_rate_raised, InverterPrice, pv_balance,
                                                         Structure_Type, advancepanelnames, advanceinverternames, advancecablename, NumberofDifferentInverters)

                                        filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                                            ClientLocation) + " Quotation" + str(UniqueID) + ".docx"
                                        tradeMarkRemover(filename)

                                        record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy,
                                                    Inverter_TYP, Inverter_Name, Inverter_Watt, InverterPrice,
                                                    Number_of_Inverter1, Inverter2_TYP, Inverter2_Name, Inverter2_Watt, Inverter2Price,
                                                    Number_of_Inverter2, Name_of_Panels, panelprice, Number_of_Panels, pv_balance,
                                                    Carriage_Cost, Installation, Net_Metering, Foundation, Earthing_val,
                                                    Structure_Type, template_file,  Quotation_TYP, BatteryPrice, BatteryName, BatteryPieces,
                                                    BatterySpecs, TotalCostNormal, TotalCostRaised, TotalCostNormalNNI,
                                                    TotalCostRaisedNNI, NumberofDifferentInverters)
                                        tkinter.messagebox.showinfo(title="Success",
                                                                    message="File is Created and Recorded\nFileName: " + filename)
                                    else:
                                        tkinter.messagebox.showwarning(title="Error",
                                                                       message="Enter Carriage, Installation and Net Metering Cost")
                                else:
                                    tkinter.messagebox.showwarning(title="Error",
                                                                   message="Battery Name, Rate and Pieces")
                            else:
                                Foundation = foundation_work_entry.get()
                                Carriage_Cost = carriage_entry.get()
                                Installation = installation_entry.get()
                                Net_Metering = net_metering_entry.get()
                                Earthing_val = Earthing_entry.get()
                                if Carriage_Cost and Installation and Net_Metering:
                                    print("System Size: ", SystemSize, "Client Name: ", ClientName, "Client Location: ",
                                          ClientLocation)
                                    print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP,
                                          "Inverter Name: ",
                                          Inverter_Name)
                                    print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels,
                                          "Panel Price: ", panelprice)
                                    print("Inverter 1 Price", InverterPrice, "Inverter 2 Price", Inverter2Price)
                                    print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                                    print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost,
                                          "Installation Cost: ",
                                          Installation)
                                    print("Net Metering: ", Net_Metering, "Template File", template_file,
                                          "Inverter Price",
                                          InverterPrice)
                                    print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ",
                                          TotalCostRaised,
                                          "Unique ID: ", UniqueID)
                                    print("------------------------------------------")

                                    document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                                                    Inverter_Watt, Inverter2_Watt, Number_of_Panels,
                                                    Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                                    TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords,
                                                    Inverter_Name, Inverter2_Name,
                                                    Name_of_Panels, "", "", "", "",
                                                    Number_of_inverters, panelwattage,
                                                    Carriage_Cost, Installation, Foundation, Earthing_val,
                                                    panelprice, structure_rate_normal, structure_rate_raised, InverterPrice,
                                                    pv_balance, Structure_Type, advancepanelnames, advanceinverternames,advancecablename, NumberofDifferentInverters)

                                    filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                                        ClientLocation) + " Quotation" + str(UniqueID) + ".docx"
                                    tradeMarkRemover(filename)

                                    record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy,
                                                Inverter_TYP, Inverter_Name, Inverter_Watt, InverterPrice,
                                                Number_of_Inverter1,
                                                Inverter2_TYP, Inverter2_Name, Inverter2_Watt, Inverter2Price,
                                                Number_of_Inverter2,
                                                Name_of_Panels, panelprice, Number_of_Panels, pv_balance, Carriage_Cost,
                                                Installation, Net_Metering, Foundation, Earthing_val, Structure_Type,
                                                template_file,
                                                Quotation_TYP, "", "", "", "",
                                                TotalCostNormal, TotalCostRaised, TotalCostNormalNNI,
                                                TotalCostRaisedNNI, NumberofDifferentInverters)
                                    tkinter.messagebox.showinfo(title="Success",
                                                                message="File is Created and Recorded\nFileName: " + filename)
                                else:
                                    tkinter.messagebox.showwarning(title="Error",
                                                                   message="Enter Carriage, Installation and Net Metering Cost")
                        else:
                            tkinter.messagebox.showwarning(title="Error",
                                                           message="Enter Structure Type and PV Balance.")
                else:
                    Name_of_Panels = panel_name_combobox.get()
                    Structure_Type = structure_type_combobox.get()
                    pv_balance = pv_balance_combobox.get()
                    if Name_of_Panels and Structure_Type and pv_balance:
                        BatteryPrice = Battery_Price_Entry.get()
                        BatteryName = Battery_Name_Entry.get()
                        BatteryPieces = Number_of_Batteries_Entry.get()
                        BatterySpecs = Battery_Specification_Entry.get()
                        if Inverter_TYP == "Hybrid":
                            if BatteryPrice and BatteryName and BatteryPieces:
                                Earthing_val = Earthing_entry.get()
                                Foundation = foundation_work_entry.get()
                                Carriage_Cost = carriage_entry.get()
                                Installation = installation_entry.get()
                                Net_Metering = net_metering_entry.get()
                                if Carriage_Cost and Installation and Net_Metering and Foundation and Earthing_val:
                                    print("System Size: ", SystemSize, "Client Name: ", ClientName, "Client Location: ",
                                          ClientLocation)
                                    print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP,
                                          "Inverter Name: ", Inverter_Name)
                                    print("Inverter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels,
                                          "Panel Price: ", panelprice)
                                    print("Inverter 1 Price", InverterPrice, "Inverter 2 Price", Inverter2Price)
                                    print("No of Panels: ", Number_of_Panels, "Structure Type: ",
                                          Structure_Type)
                                    print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost,
                                          "Installation Cost: ", Installation)
                                    print("Net Metering: ", Net_Metering, "Template File", template_file,
                                          "Inverter Price", InverterPrice)
                                    print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ",
                                          TotalCostRaised, "Unique ID: ", UniqueID)
                                    print("------------------------------------------")

                                    document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                                                     Inverter_Watt, Inverter2_Watt, Number_of_Panels,
                                                     Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                                     TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords,
                                                     Inverter_Name, Inverter2_Name,
                                                     Name_of_Panels, BatteryPrice, BatteryName, BatteryPieces,
                                                     BatterySpecs,
                                                     Number_of_inverters, panelwattage,
                                                     Carriage_Cost, Installation, Foundation, Earthing_val,
                                                     panelprice, structure_rate_normal, structure_rate_raised,
                                                     InverterPrice, pv_balance,
                                                     Structure_Type, advancepanelnames, advanceinverternames,
                                                     advancecablename, NumberofDifferentInverters)

                                    filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                                        ClientLocation) + " Quotation" + str(UniqueID) + ".docx"
                                    tradeMarkRemover(filename)

                                    record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy,
                                                Inverter_TYP, Inverter_Name, Inverter_Watt, InverterPrice,
                                                Number_of_Inverter1,
                                                Inverter2_TYP, Inverter2_Name, Inverter2_Watt, Inverter2Price,
                                                Number_of_Inverter2,
                                                Name_of_Panels, panelprice, Number_of_Panels, pv_balance, Carriage_Cost,
                                                Installation, Net_Metering, Foundation, Earthing_val, Structure_Type,
                                                template_file,
                                                Quotation_TYP, BatteryPrice, BatteryName, BatteryPieces, BatterySpecs,
                                                TotalCostNormal, TotalCostRaised, TotalCostNormalNNI,
                                                TotalCostRaisedNNI, NumberofDifferentInverters)
                                    tkinter.messagebox.showinfo(title="Success",
                                                                message="File is Created and Recorded\nFileName: " + filename)
                                else:
                                    tkinter.messagebox.showwarning(title="Error",
                                                                   message="Enter Carriage, Installation and Net Metering Cost")
                        else:
                            Foundation = foundation_work_entry.get()
                            Carriage_Cost = carriage_entry.get()
                            Installation = installation_entry.get()
                            Net_Metering = net_metering_entry.get()
                            Earthing_val = Earthing_entry.get()
                            if Carriage_Cost and Installation and Net_Metering:
                                print("System Size: ", SystemSize, "Client Name: ", ClientName,
                                      "Client Location: ", ClientLocation)
                                print("Referred by: ", ReferredBy, "Inverter Type: ", Inverter_TYP,
                                      "Inverter Name: ", Inverter_Name)
                                print("Inver4ter Wattage ", Inverter_Watt, "Panel Name: ", Name_of_Panels,
                                      "Panel Price: ", panelprice)
                                print("Inverter 1 Price", InverterPrice, "Inverter 2 Price", Inverter2Price)
                                print("No of Panels: ", Number_of_Panels, "Structure Type: ", Structure_Type)
                                print("PV Balance: ", pv_balance, "Carriage: ", Carriage_Cost,
                                      "Installation Cost: ", Installation)
                                print("Net Metering: ", Net_Metering, "Template File", template_file,
                                      "Inverter Price", InverterPrice)
                                print("Total Cost Normal: ", TotalCostNormal, "Total Cost Raised: ",
                                      TotalCostRaised, "Unique ID: ", UniqueID)
                                print("------------------------------------------")

                                document_creater(UniqueID, ClientName, ClientLocation, SystemSize, Inverter_TYP,
                                                 Inverter_Watt, Inverter2_Watt, Number_of_Panels,
                                                 Net_Metering, template_file, TotalCostNormal, TotalCostRaised,
                                                 TotalCostNormalNNI, TotalCostRaisedNNI, valueinwords,
                                                 Inverter_Name, Inverter2_Name,
                                                 Name_of_Panels, "", "", "", "",
                                                 Number_of_inverters, panelwattage,
                                                 Carriage_Cost, Installation, Foundation, Earthing_val,
                                                 panelprice, structure_rate_normal, structure_rate_raised,
                                                 InverterPrice, pv_balance,
                                                 Structure_Type, advancepanelnames, advanceinverternames,
                                                 advancecablename, NumberofDifferentInverters)

                                filename = str(SystemSize) + "kW " + str(Inverter_TYP) + " " + str(
                                    ClientLocation) + " Quotation" + str(UniqueID) + ".docx"
                                tradeMarkRemover(filename)

                                record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy,
                                            Inverter_TYP, Inverter_Name, Inverter_Watt, InverterPrice,
                                            Number_of_Inverter1,
                                            Inverter2_TYP, Inverter2_Name, Inverter2_Watt, Inverter2Price,
                                            Number_of_Inverter2,
                                            Name_of_Panels, panelprice, Number_of_Panels, pv_balance, Carriage_Cost,
                                            Installation, Net_Metering, Foundation, Earthing_val, Structure_Type,
                                            template_file,
                                            Quotation_TYP, "", "", "", "",
                                            TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI, NumberofDifferentInverters)
                                tkinter.messagebox.showinfo(title="Success",
                                                            message="File is Created and Recorded\nFileName: " + filename)
                            else:
                                tkinter.messagebox.showwarning(title="Error",
                                                               message="Enter Carriage, Installation and Net Metering Cost")
                    else:
                        tkinter.messagebox.showwarning(title="Error", message="Enter Structure Type and PV Balance.")
            else:
                tkinter.messagebox.showwarning(title="Error", message="Enter All box of Inverters.")
        else:
            tkinter.messagebox.showwarning(title="Error", message="Enter All boxes of Client Information.")


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
    Number_of_Inverter1 = inverter_number_Entry.get()
    Number_of_Inverter2 = inverter2_number_Entry.get()

    TotalCostRaised = round_up_to_nearest_thousand( (int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) +
                                                    (int(structure_rate_raised) * int(Number_of_Panels) * int(panelwattage)) + ( int(panelwattage) * int(panelprice) *
                                                    int(Number_of_Panels)) + int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Net_Metering) + int(Foundation_Work))

    print("Total Cost with Raised Structure and Net Metering Included: " + str(TotalCostRaised))

    TotalCostRaisedNNI = round_up_to_nearest_thousand((int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) +
                                                      (int(structure_rate_raised) *int(Number_of_Panels) * int(panelwattage)) + (int(panelwattage) * int(panelprice) *
                                                    int(Number_of_Panels)) + int(pv_balance) + int(Carriage_Cost) + int(Installation) + int( Foundation_Work))

    print("Total Cost with Raised Structure and Net Metering Not Included: " + str(TotalCostRaisedNNI))

    TotalCostNormal = round_up_to_nearest_thousand((int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) +
                                                    int((int(structure_rate_normal) * (int(Number_of_Panels) / 2))) + (int(panelwattage) * int(panelprice) *
                                                    int(Number_of_Panels)) + int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Net_Metering) + int(Foundation_Work))

    print("Total Cost with Normal Structure and Net Metering Included: " + str(TotalCostNormal))

    TotalCostNormalNNI = round_up_to_nearest_thousand((int(InverterPrice) * int(Number_of_Inverter1)) + (int(Inverter2Price) * int(Number_of_Inverter2)) + int(
                                                    (int(structure_rate_normal) * (int(Number_of_Panels) / 2))) + (int(panelwattage) * int(panelprice) *
                                                    int( Number_of_Panels)) + int(pv_balance) + int(Carriage_Cost) + int(Installation) + int(Foundation_Work))
    print("Total Cost with Normal Structure and Net Metering Not Included: " + str(TotalCostNormalNNI))


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
    unique_id_list = Customer_Data_Sheet.col_values(3)
    unique_id_list.pop(0)
    print(unique_id_list)
    unique_id = int(f"{timestamp:03d}{random_component:03d}") % 10000000
    index = 0
    for x in unique_id_list:
        if unique_id == int(x):
            index = 1
        else:
            index = 0
    if index == 0:
        return unique_id
    else:
        generate_unique_id()
    print(unique_id)


def record_data(SystemSize, UniqueID, ClientName, ClientLocation, ReferredBy,
                Inverter_TYP, Inverter_Name, Inverter_Watt, InverterPrice, Number_of_Inverter1,
                Inverter2_TYP, Inverter2_Name, Inverter2_Watt, Inverter2Price, Number_of_Inverter2,
                Name_of_Panels, panelprice, Number_of_Panels, pv_balance, Carriage_Cost,
                Installation, Net_Metering, Foundation, Earthing_val, Structure_Type, template_file,
                Quotation_TYP, BatteryPrice, BatteryName, BatteryPieces, BatterySpecs,
                TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI, NumberofDifferentInverters):
    serial_list = []
    current_date_time = datetime.now()
    current_date_time = str(current_date_time)
    current_date_time = current_date_time[:-7]
    for x in Customer_Data_Sheet.col_values(1):
        serial_list.append(x)
    SerialNumber = int(serial_list[-1]) + 1
    if SerialNumber == 0:
        SerialNumber = 1
    Customer_Data_Sheet.append_row(
        [SerialNumber, current_date_time, UniqueID, SystemSize, ClientName, ClientLocation, ReferredBy,
         Inverter_TYP, Inverter_Name, Inverter_Watt, InverterPrice, Number_of_Inverter1,
         Inverter2_TYP, Inverter2_Name, Inverter2_Watt, Inverter2Price, Number_of_Inverter2,
         Name_of_Panels, panelprice, Number_of_Panels, pv_balance, Carriage_Cost,
         Installation, Net_Metering, Foundation, Earthing_val, Structure_Type, template_file,
         Quotation_TYP, BatteryPrice, BatteryName, BatteryPieces, BatterySpecs,
         TotalCostNormal, TotalCostRaised, TotalCostNormalNNI, TotalCostRaisedNNI, NumberofDifferentInverters])


def update_names_of_inverters_and_panels(*args):
    global advancepanelnames
    global advanceinverternames
    global advancecablename
    if inverter_type_combobox.get() == "Hybrid":
        advanceinverternames = "Crown/SolarMax/MaxPower/Infinix"
        advancepanelnames = "Canadian/Jinko/JA/Astro/Trina/Longi"
        advancecablename = "AC Single Core 6mm Cable"
    else:
        if inverter_type_combobox.get() == "Grid Tie":
            advanceinverternames = "Sofar/Solis/Growatt/Sunways"
            advancepanelnames = "Canadian/Jinko/JA/Astro/Trina/Longi"
            advancecablename = "AC Single Core 6mm Cable"


def get_template_type(inverter_selection,structure_type,quotation_type,inverter_type,inverter2_type):
    template_file = ""
    update_names_of_inverters_and_panels()
    if inverter_selection == "1":
        if inverter_type == "Grid Tie":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (
                        ".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringNotIncluded"
                        "\\GTGRNI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (
                        ".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITR_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRNI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRNISPI_Template.docx")
        if inverter_type == "Hybrid":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGNSPI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITNSPI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_Template.docx")
    if inverter_selection == "2":
        if inverter_type == "Grid Tie" and inverter2_type == "Hybrid":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_WHI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_WHI_Template.docx")
                if quotation_type == "Itemised Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_WHI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_WHI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITR_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNISPI_WHI_Template.docx")
        if inverter_type == "Grid Tie" and inverter2_type == "Grid Tie":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_WGTI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_WGTI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTR_WGTI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTIT_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNISPI_WGTI_Template.docx")
        if inverter_type == "Hybrid" and inverter2_type == "Grid Tie":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WGTI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGNSPI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITNSPI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_WGTI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_WGTI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_WGTI_Template.docx")
        if inverter_type == "Hybrid" and inverter2_type == "Hybrid":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WHI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_WHI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_WHI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_WHI_Template.docx")
    print("Template File: " + template_file)
    return template_file

def update_template_type(*args):
    update_names_of_inverters_and_panels()
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (
                        ".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringNotIncluded"
                        "\\GTGRNI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (
                        ".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITR_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRNI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRNISPI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITNSPI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_WHI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITR_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNISPI_WHI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_WGTI_Template.docx")
            if structure_type_combobox.get() == "Raised":
                if Quotation_type_combobox.get() == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTR_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTIT_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNISPI_WGTI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITNSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_WGTI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_WGTI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_WGTI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_WHI_Template.docx")
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
                if Quotation_type_combobox.get() == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_WHI_Template.docx")
                if Quotation_type_combobox.get() == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_WHI_Template.docx")
    print("Template File: " + template_file)


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
    print("Panel Price: " + panelprice)
    for x in Solar_Panel_Wattage:
        panelwattage = Solar_Panel_Wattage[i]
    PanelWattageInt = int(panelwattage)
    print("Panel Wattage: " + panelwattage)
    Panels_Numbers(PanelWattageInt)


def Panels_Numbers(panelwattage):
    global Number_of_Panels
    SZ = float(System_Size_combobox.get()) * 1000
    Number_of_Panels = SZ / int(panelwattage)
    Number_of_Panels = math.ceil(Number_of_Panels)
    Number_of_Panels = int(Number_of_Panels)
    print("Number of Panels: " + str(Number_of_Panels))


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
            for x in Hybrid_Inverters.row_values(int(index)):
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
            for x in Hybrid_Inverters.row_values(int(index)):
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
            for x in Hybrid_Inverters.row_values(int(index)):
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
            for x in Hybrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print("Inverter List" + str(list))
    if inverter_wattage_combobox.current() >= 0 and inverter_wattage_combobox.current() <= 20:
        index2 = int(inverter_wattage_combobox.current())
        global InverterPrice
        InverterPrice = list[index2]
        print("Inverter Price" + str(InverterPrice))


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
            for x in Hybrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print("Inverter 2 List" + str(list))
    if inverter2_wattage_combobox.current() >= 0 and inverter2_wattage_combobox.current() <= 20:
        index2 = int(inverter2_wattage_combobox.current())
        global Inverter2Price
        Inverter2Price = list[index2]
        print("Inverter 2 Price" + str(InverterPrice))


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
            for x in Hybrid_Inverters.row_values(int(index)):
                list.append(x)
            while ("" in list):
                list.remove("")
    print("Inverter 3 List" + str(list))
    if inverter3_wattage_combobox.current() >= 0 and inverter3_wattage_combobox.current() <= 20:
        index2 = int(inverter3_wattage_combobox.current())
        global Inverter3Price
        Inverter3Price = list[index2]
        print("Inverter 3 Price" + str(InverterPrice))


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

systemtracker = tkinter.StringVar()

System_Size_label = tkinter.Label(user_info_frame, text="System Size")
System_Size_combobox = ttk.Entry(user_info_frame, textvariable=systemtracker)
System_Size_label.grid(row=generalrow, column=0)
System_Size_combobox.grid(row=generalrowentry, column=0)

systemtracker.trace('w', update_panel)

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
                                         "Engr Ubaid", "Sir Nabeel", "Engr Osama", "Engr Abdullah", "ELAF", "SBD"])
Reffered_label.grid(row=generalrow, column=3)
Reffered_combobox.grid(row=generalrowentry, column=3)

inverter_selection_frame = tkinter.LabelFrame(frame, text="DIS")
inverter_selection_frame.grid(row=inverterselectionrow, column=0, sticky="news", padx=20, pady=15)

tracker_inverters = tkinter.StringVar(inverter_selection_frame)

inverter_selection_label = tkinter.Label(inverter_selection_frame, text="Number of Different Inverters")
inverter_selection_combobox = ttk.Combobox(inverter_selection_frame, values=['1', '2', '3'], textvariable=tracker_inverters)
inverter_selection_label.grid(row=inverterselectionrow, column=0)
inverter_selection_combobox.grid(row=inverterselectionrowentry, column=0)

tracker_inverters.trace('w', inverter_number_selection)

foundation_work = tkinter.Label(inverter_selection_frame, text="Foundation Work")
foundation_work_entry = ttk.Entry(inverter_selection_frame)
foundation_work.grid(row=inverterselectionrow, column=1)
foundation_work_entry.grid(row=inverterselectionrowentry, column=1)

if foundation_work_entry.get() == "":
    foundation_work_entry.insert(0, "0")

OverrideRate = ttk.Button(text="Override Rates")

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

if inverter_selection_combobox.get() == '1':
    if inverter2_number_Entry.get() == "":
        inverter2_number_Entry.insert(0, '0')
else:
    if inverter2_number_Entry.get() == "":
        inverter2_number_Entry.insert(0, '0')

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

if inverter_selection_combobox.get() == "3":
    if inverter3_number_Entry.get() == "":
        inverter3_number_Entry.insert(0, '1')
else:
    if inverter3_number_Entry.get() == "":
        inverter3_number_Entry.insert(0, '0')

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

pv_balance_label = tkinter.Label(panel_frame, text="PV Balance")
pv_balance_combobox = ttk.Entry(panel_frame)
pv_balance_label.grid(row=panelrow, column=1, padx=5)
pv_balance_combobox.grid(row=panelrowentry, column=1, padx=5)

structure_type = tkinter.Label(panel_frame, text="Structure Type")
structure_type_combobox = ttk.Combobox(panel_frame, values=["Normal", "Raised"], textvariable=sel4)
structure_type.grid(row=panelrow, column=2, padx=5)
structure_type_combobox.grid(row=panelrowentry, column=2, padx=5)

if structure_type_combobox.get() == "":
    structure_type_combobox.set("Normal")

update_tracker = tkinter.StringVar()

Quotation_type = tkinter.Label(panel_frame, text="Quotation Type")
Quotation_type_combobox = ttk.Combobox(panel_frame,
                                       values=["General Net Metering Included", "Specify Brand Net Metering Included",
                                               "General Net Metering Not Included",
                                               "Specify Brand Net Metering Not Included",
                                               "Itemised General Net Metering Included",
                                               "Itemised Specify Brand Net Metering Included",
                                               "Itemised General Net Metering Not Included",
                                               "Itemised Specify Brand Net Metering Not Included"],
                                       textvariable=update_tracker)
Quotation_type.grid(row=panelrow, column=3, padx=5)
Quotation_type_combobox.grid(row=panelrowentry, column=3, padx=5)

update_tracker.trace('w', update_template_type)
if Quotation_type_combobox.get() == "":
    Quotation_type_combobox.set("Net Metering Not Included")

sel4.trace('w', update_template_type)

battery_frame = tkinter.LabelFrame(frame, text="Battery Area")
battery_frame.grid(row=batteryrow, column=0, sticky="news", padx=20, pady=15)

Battery_Name_Label = tkinter.Label(battery_frame, text="Battery Name")
Battery_Name_Entry = ttk.Entry(battery_frame)
Battery_Name_Label.grid(row=batteryrow, column=0, padx=10)
Battery_Name_Entry.grid(row=batteryentryrow, column=0, padx=10)

if Battery_Name_Entry.get() == "":
    Battery_Name_Entry.insert(0, "Daewoo Deep Cycle")

Battery_Price_Label = tkinter.Label(battery_frame, text="Battery Price")
Battery_Price_Entry = ttk.Entry(battery_frame)
Battery_Price_Label.grid(row=batteryrow, column=1, padx=10)
Battery_Price_Entry.grid(row=batteryentryrow, column=1, padx=10)

if Battery_Price_Entry.get() == "":
    Battery_Price_Entry.insert(0, "45000")

Number_of_Batteries_Label = tkinter.Label(battery_frame, text="Number of Batteries")
Number_of_Batteries_Entry = ttk.Entry(battery_frame)
Number_of_Batteries_Label.grid(row=batteryrow, column=2, padx=10)
Number_of_Batteries_Entry.grid(row=batteryentryrow, column=2, padx=10)

Battery_Specification_Label = tkinter.Label(battery_frame, text="Specifications of Battery")
Battery_Specification_Entry = tkinter.Entry(battery_frame)
Battery_Specification_Label.grid(row=batteryrow, column=3, padx=10)
Battery_Specification_Entry.grid(row=batteryentryrow, column=3, padx=10)

if Number_of_Batteries_Entry.get() == "":
    Number_of_Batteries_Entry.insert(0, '4')

if Battery_Specification_Entry.get() == "":
    Battery_Specification_Entry.insert(0, "12 V  180 AH")

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

Earthing = tkinter.Label(courses_frame, text="Earthing")
Earthing_entry = ttk.Entry(courses_frame)
Earthing.grid(row=cinrow, column=3)
Earthing_entry.grid(row=cinrowentry, column=3)

if Earthing_entry.get() == "":
    Earthing_entry.insert(0, "0")

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=12, pady=5)

options_frame = tkinter.LabelFrame(frame, text="Options Frame")
options_frame.grid(row=17, column=0, sticky="news", padx=20, pady=10)

overrride_button = tkinter.Button(options_frame, text="Override Rates", command=override_rates)
overrride_button.grid(row=18, column=0, sticky="news", padx=10, pady=10)

advance_button = tkinter.Button(options_frame, text="Advance Option", command=advance_options)
advance_button.grid(row=18, column=1, sticky="news", padx=10, pady=10)

exist_button = tkinter.Button(options_frame, text="ID Already Exist", command=itemised_exist)
exist_button.grid(row=18, column=2, sticky="news", padx=10, pady=10)
# Button
button = tkinter.Button(frame, text="Enter data", command=enter_data)
button.grid(row=19, column=0, sticky="news", padx=20, pady=10)

window.mainloop()
