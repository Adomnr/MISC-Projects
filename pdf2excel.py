from spire.pdf.common import *
from spire.pdf import *

for i in range(5768,5769):
    inputFile = "New_Report/Report_page_{}.pdf".format(i)
    outputFile = "Report_Excel/Report_page{}.xlsx".format(i)
    # Create a PdfDocument object
    pdf = PdfDocument()
    # Load a PDF document
    pdf.LoadFromFile(inputFile)
    # Save the PDF file to Excel XLSX format
    pdf.SaveToFile(outputFile, FileFormat.XLSX)
    pdf.Close()
    print("Created Report_page{} excel file".format(i))

print("complete")