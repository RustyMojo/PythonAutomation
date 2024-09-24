# PDF Merger for Word Documents

## Introduction
This report details the functionality and usage of a Python script designed to convert Microsoft Word documents (.docx) into PDF format and subsequently merge these PDFs into a single document. The script leverages the `docx2pdf` library for conversion and the `PyPDF2` library for merging, providing a straightforward solution for users needing to combine multiple letters or documents into one cohesive PDF file.

## Key Functionalities
The script performs the following key functions:
1. **File Conversion**: Converts each `.docx` file in a specified directory into a corresponding PDF file.
2. **PDF Merging**: Combines all generated PDF files into a single output PDF.
3. **Directory Handling**: Automatically processes all `.docx` files found in the designated folder.

## How It Works
1. **Directory Specification**: The user specifies the directory containing the Word documents.
2. **PDF Merger Initialization**: A `PdfMerger` object is created to handle the merging process.
3. **File Iteration**: The script iterates through each file in the specified directory, checking for files with the `.docx` extension.
4. **File Conversion**:
   - Each Word document is converted to PDF format.
   - The newly created PDF file is stored in the same directory as the original Word document.
5. **Merging PDFs**: Each converted PDF is appended to the merger object.
6. **Final Output**: The combined PDF is saved to the current working directory, named `combined_document.pdf`.

## Conclusion
This PDF merger script provides a reliable and efficient method for converting and merging Word documents into a single PDF file. By automating the process, it saves time and reduces manual effort, making it a valuable tool for users who frequently handle document preparation and presentation. With straightforward usage and the ability to process multiple files seamlessly, it is an essential utility for both personal and professional use.
