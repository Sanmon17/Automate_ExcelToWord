# Automate_ExcelToWord
Excel to Word Automation Script This Python script automates the process of exporting data and images from an Excel sheet to a specific section in a Word document. It supports handling text and images, and can be customized for different Excel sheets, sections, and formatting requirements.

# Features:
-Text Export: Extracts data from an Excel sheet (skipping the first few rows) and inserts it into a specified section of a Word document.
-Image Extraction: Collects images (such as screenshots) pasted into the Excel sheet and inserts them into the Word document in the correct order.
-Image Resizing: Automatically resizes images to fit within a specified maximum width (6.5 inches by default).

# Requirements:
-Python 3.x
-openpyxl (for reading Excel files)
-python-docx (for manipulating Word documents)
-pywin32 (for working with Excel via COM interface)
-Pillow (for clipboard image processing)
**Installation**: `pip install openpyxl python-docx pywin32 Pillow`

# Usage:
Run the script from the command line:
  `python export_excel_to_word.py <excel_file_path> <sheet_name> <word_file_path> <section_title>`
  ==============================================================================================
  Arguments:
  *excel_file_path*: Path to the Excel file
  *sheet_name*: Name of the sheet within the Excel file
  *word_file_path*: Path to the Word document
  *section_title*: Title of the section in the Word document where the data and images will be inserted

# Example:
**python export_excel_to_word.py report.xlsx "Sheet1" report.docx "Summary"**
This command will extract data from the "Sheet1" of report.xlsx, and insert the data and any images from that sheet into the "Summary" section of report.docx.

# Note:
This script will skip data from cell 1 to 4 of the excel file.
First it transfer all texts then all images.
