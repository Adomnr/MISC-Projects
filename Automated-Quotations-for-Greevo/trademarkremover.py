from docx import Document
from pdfConversion import *
import os
from datetime import datetime

def tradeMarkRemover(filename):
    current_date = datetime.now()

    directory_year_month = current_date.strftime("%B %Y")
    directory_day = current_date.strftime("%B %d")

    # Create the directories if they don't exist
    os.makedirs(os.path.join(directory_year_month, directory_day, "Word Files"), exist_ok=True)
    os.makedirs(os.path.join(directory_year_month, directory_day, "TradeMark Removed File"), exist_ok=True)
    os.makedirs(os.path.join(directory_year_month, directory_day, "PDF Files"), exist_ok=True)

    # Combine the directory paths
    folder_path = os.path.join(directory_year_month, directory_day)

    # Path to the new file
    file_path = os.path.join(folder_path, "Word Files")
    output_file_path_current = os.path.join(folder_path, "TradeMark Removed File")
    pdf_file_path = os.path.join(folder_path, "PDF Files")
    template_file_path = os.path.join(file_path, filename)

    pdf_file_name = filename[:-5] + ".pdf"
    print(template_file_path)
    print(filename)
    output_file_path = os.path.join(output_file_path_current,filename)
    savepath = os.path.join(pdf_file_path,pdf_file_name)

    variables = {
        'Evaluation Warning: The document was created with Spire.Doc for Python.': ' '
    }

    template_document = Document(template_file_path)

    for variable_key, variable_value in variables.items():
        for paragraph in template_document.paragraphs:
            replace_text_in_paragraph(paragraph, variable_key, variable_value)

    template_document.save(output_file_path)
    pdf_conversion(output_file_path, savepath)


def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)
