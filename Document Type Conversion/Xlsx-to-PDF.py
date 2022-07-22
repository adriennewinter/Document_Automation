# -----------------------------------------------------------------------------------
# Convert Excel sheets into PDF files.

# Author: Adrienne Winter, 2022.

# README:
# You will need to install some dependencies before using:
# pip install aspose-cells
# pip install JPype1
# Make sure you have Java (JDK) installed https://www.oracle.com/java/technologies/downloads/ 
# Make sure the JAVA_HOME environment variable is set to the path where your jvm.dll is saved 
# e.g. "C:\Program Files\Java\jdk-18.0.2\bin\server"

# Set the folder path variable for the folder containing the Excel sheets to convert into PDF.
# Remember to add an escape \ in your file path name -> C:File\\Path\\Name
# -----------------------------------------------------------------------------------

import os
import shutil
import jpype
jpype.startJVM()
import asposecells
from asposecells.api import *

# Setting variables
folderPath = "C:\\Users\\Dell\\Documents\\MASTERS\\Funding\\OMT - Oppenheimer Memorial Trust"

# Create folder to save PDFs into
os.makedirs(f"{folderPath}\\Excel_to_PDF")

# Convert Excel to PDF and save in PDF folder "Excel_to_PDF"
for fileName in os.listdir(folderPath):
    fileExtension = fileName.split('.')[-1]
    if fileExtension != ("xlsx" or "xls"):
        print(f"Skipping '{fileName}' because it is not an Excel file.")
    else:
        excelFile = Workbook(f"{folderPath}\\{fileName}")
        PDFname = fileName.split('.')[0]
        excelFile.save(f"{folderPath}\\{PDFname}.pdf", SaveFormat.PDF)
        shutil.move(f"{folderPath}\\{PDFname}.pdf", f"{folderPath}\\Excel_to_PDF")
        print(f"Creating file {PDFname}.pdf in {folderPath}\\Excel_to_PDF")

# End the Java Virtual Machine
jpype.shutdownJVM()