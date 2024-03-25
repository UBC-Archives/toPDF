# toPDF
Python document and image to PDF convertor

> Version: 0.1

> [UBC Records Management Office](https://recordsmanagement.ubc.ca)

> [GPL-3.0 License](https://www.gnu.org/licenses/gpl-3.0.en.html)

**toPDF** is a Python application designed to convert documents and images into PDF format. It provides a simple graphical interface for users to select a directory containing files to convert and then converts them into PDFs. The application also combines the converted PDFs into a single file for easy access.

![image](https://github.com/UBC-Archives/toPDF/assets/6263442/61cf78f1-fdcd-4b8d-a6db-133922a9e47e)

**Features:**
- Converts various file types including images (JPEG, PNG, etc.) and documents (DOCX) into PDF format (complete list of supported image formats can be found [here](https://pillow.readthedocs.io/en/stable/handbook/image-file-formats.html)).
- Combines the converted PDFs into a single PDF file.
- Generates a word cloud PDF from the text extracted from DOCX files.

**How to Use:**

1. Select Directory: Click on the "Browse" button to select the directory containing the files you want to convert.

2. Convert: Click the "Convert" button to start the conversion process.

3. View Results: Once the conversion is complete, the combined PDF and word cloud PDF (if applicable) will be saved in the same directory as the input files.

**Installation:**

- Install Python: Ensure Python 3.x is installed on your system. You can download Python from [here](https://www.python.org/downloads).
- Execute the following command in the Command Prompt to install dependencies:

  > pip install Pillow PyPDF2 python-docx docx2pdf wordcloud matplotlib
- Download the Script: Download the Directory Inventory Generator script from [this link](https://github.com/UBC-Archives/toPDF/blob/main/UBC-RMO_toPDF.py).
- Once downloaded, double-click on the downloaded UBC-RMO_toPDF.py file to run the program.

**Disclaimer:**

toPDF is provided as-is without any warranty. The University of British Columbia or its affiliates shall not be held responsible for any loss of data or unintended consequences resulting from the use of this application. Use it at your own risk and ensure you have backups of important data before running the conversion process.
