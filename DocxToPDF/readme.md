# DOCX to PDF Converter Script

## Introduction
This report details the functionality and design of a Python script created for converting `.docx` files to PDF format. Utilizing the `docx2pdf` library for document conversion and the built-in `os` module for file manipulation, the script aims to simplify the process of managing document formats. With an easy-to-follow structure, this tool is intended for users who frequently work with Word documents and require an efficient method for conversion.

## Key Functionalities
The script provides two primary functionalities:
1. **Convert DOCX Files to PDF**: The script scans a specified directory for `.docx` files, converts them to PDF format, and saves the converted files in a designated output folder. This process streamlines the task of converting multiple documents at once.
2. **Rename DOCX Files**: To ensure consistency in file naming, the script renames any files with a `.DOCX` extension (uppercase) to lowercase `.docx`. This prevents potential issues when managing files on case-sensitive systems.

## Steps Involved
1. **File Iteration**: The script loops through all files in the specified directory.
2. **File Filtering**: Files with a `.DOCX` extension are targeted for renaming.
3. **New Filename Construction**: A new filename is created by changing the extension to lowercase.
4. **Renaming Process**: The file is renamed using `os.rename`.
5. **Directory Creation**: The output directory is created if it doesn't already exist.
6. **File Iteration**: The script loops through all files in the input directory.
7. **File Filtering**: Only files with a `.docx` extension are processed.
8. **Path Construction**: Full paths for the input and output files are created.
9. **Conversion**: The conversion is performed using `docx2pdf.convert`.
10. **Error Handling**: Any errors encountered during conversion are logged.

## Result
The result is a new folder containing all the Word documents converted into PDFs inside of it.
