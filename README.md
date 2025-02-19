# Automate_ExcelToWord
Excel to Word Automation Script. This Python script automates the process of exporting data and images from an Excel sheet to a specific section in a Word document. It supports handling text and images, and can be customized for different Excel sheets, sections, and formatting requirements.

# Features:
- **Text Export**: Extracts data from an Excel sheet (skipping the first few rows) and inserts it into a specified section of a Word document.<br>
- **Image Extraction**: Collects images (such as screenshots) pasted into the Excel sheet and inserts them into the Word document in the correct order.<br>
- **Image Resizing**: Automatically resizes images to fit within a specified maximum width (6.5 inches by default).<br>

# Requirements:
- **Python 3.x**<br>
- **openpyxl** (for reading Excel files)<br>
- **python-docx** (for manipulating Word documents)<br>
- **pywin32** (for working with Excel via COM interface)<br>
- **Pillow** (for clipboard image processing)<br><br>
**Installation**: `pip install openpyxl python-docx pywin32 Pillow`<br>

# Usage:
Run the script from the command line:<br>
**`python export_excel_to_word.py <excel_file_path> <sheet_name> <word_file_path> <section_title>`**<br>
  
  **Arguments**:<br>
  - *excel_file_path*: Path to the Excel file<br>
  - *sheet_name*: Name of the sheet within the Excel file<br>
  - *word_file_path*: Path to the Word document<br>
  - *section_title*: Title of the section in the Word document where the data and images will be inserted<br>

# Example:
**`python export_excel_to_word.py report.xlsx "Sheet1" report.docx "Summary"`**<br><br>
This command will extract data from the "Sheet1" of report.xlsx, and insert the data and any images from that sheet into the "Summary" section of report.docx.<br>

# Note:
- This script will skip data from cell 1 to 4 of the excel file.<br>
- First it transfer all texts then all images.<br>
- It uses Clipboard to transfer data, and it is recommended to clear clipboard history after many runs by using the keyboard shortcut<br> `Win + V` 🖱️.
