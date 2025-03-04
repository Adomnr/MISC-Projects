from gui import inverter_selection_combobox,inverter_type_combobox,structure_type_combobox,Quotation_type_combobox,inverter2_type_combobox
from pathlib import Path

def update_template_type(template_file):
    home_dir = Path.home()
    print(home_dir)
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
    return template_file
