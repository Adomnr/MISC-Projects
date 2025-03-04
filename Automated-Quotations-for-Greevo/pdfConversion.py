from docx2pdf import convert

def pdf_conversion(filepath, output_path):
    convert(filepath,output_path)