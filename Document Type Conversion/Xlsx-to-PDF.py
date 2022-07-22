# -----------------------------------------------------------------------------------
# Convert Excel sheets into PDF files.

# Author: Adrienne Winter, 2022.

# README:
# Assumes that there is no file named "Excel_to_PDF" already existing in the "folderPath" directory.

# Set the folder path variable for the folder containing the Excel sheets to convert into PDF.
# Remember to add an escape \ in your file path name -> C:File\\Path\\Name
# -----------------------------------------------------------------------------------

import os
from win32com import client

# Setting variables
folderPath = "C:\\Users\\Dell\\Documents\\MASTERS\\Funding\\OMT - Oppenheimer Memorial Trust"

# Create folder to save PDFs into
os.makedirs(f"{folderPath}\\Excel_to_PDF")

# Save Excel files as PDFs
for fileName in os.listdir(folderPath):
    if fileName.split(".")[-1] != ("xlsx" or "xls"):
        print(f"Skipping {fileName} as it is not an Excel file.")
    else:
        excelObj = client.Dispatch("Excel.Application")
        excelWorkbook = excelObj.Workbooks.Open(f"{folderPath}\\{fileName}")
        PDFname = fileName.split(".")[0]
        excelWorkbook.SaveAs(f"{folderPath}\\Excel_to_PDF\\{PDFname}", FileFormat=57)
        print(f"Saving {fileName} as PDF in {folderPath}\\Excel_to_PDF")

excelWorkbook.Close()
excelObj.Quit()