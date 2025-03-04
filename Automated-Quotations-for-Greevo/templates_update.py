from main import inverter_type_combobox,structure_type_combobox

def update_template_type(*args):
    global template_file
    if inverter_type_combobox.get() == "Grid Tie":
        template_file = ".\\Templates\\GTGN_Template.docx"
        if structure_type_combobox.get() == "Raised":
            template_file = ".\\Templates\\GTGR_Template.docx"
    elif inverter_type_combobox.get() == "Hybrid":
        template_file = ".\\Templates\\HGN_Template.docx"
        if structure_type_combobox.get() == "Raised":
            template_file = ".\\Templates\\HGR_Template.docx"
    print(template_file)