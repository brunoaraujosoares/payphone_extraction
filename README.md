# payphone_extraction

PDF to Excel Converter using regular expressions, pytesseract, PyMuPDF, and pdf2image.
This is a solution to address a specific situation. If a similar demand arises in the future, new features, interfaces, and other functionalities can be added.

## Description
This Python project converts PDF files into Excel spreadsheets, applying a defined set of extraction points to capture structured data even in PDFs with irregular formatting or images that alter the page layout. The original Jupyter notebook was converted into a `.py` script for easier execution and automation.

## Main Features
- Define multiple extraction points (56 pre-configured regions).  
- Handle PDFs with images that affect text positioning.  
- Export extracted data to Excel format.  
- Code adapted for batch execution or integration with other systems.  

## Technologies Used
- Python  
- PDF reader libraries (`pdfplumber`, `PyPDF2`, etc.)  
- `pandas` for data manipulation and export  

## Possible Future Improvements
- Graphical interface for selecting extraction areas.  
- Configuring extraction points via an external file.  
- Parallel processing for large volumes of PDFs.
