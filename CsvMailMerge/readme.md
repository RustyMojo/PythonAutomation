# Personalized Letter Generator

## Introduction
This report describes a Python application designed to automate the generation of personalized letters using a Word template and data from an Excel or CSV file. The application features a graphical user interface (GUI) built with Tkinter, allowing users to easily select files and specify merge fields. The core functionalities include merging data fields from the spreadsheet into the Word template and generating individual letters based on the provided data.

## Key Functionalities
1. **File Selection**: Users can select a Word template, an Excel/CSV file, and an output directory through a user-friendly GUI.
2. **Custom Merge Fields**: Users can define custom merge fields, allowing for dynamic text replacement in the template.
3. **Preview Mode**: Users can preview how the generated letters will look before finalizing the generation.
4. **Letter Generation**: The tool creates individual letters for each entry in the provided spreadsheet, saving them to the specified output folder.
5. **Undo Functionality**: Users can remove letters generated in the last run, enhancing usability.

## How to Use the Letter Generator
1. **Launch the Application**: Open the Letter Generator tool.
2. **Select the Template Letter**: Click "Browse" to choose a Word template.
3. **Select the Excel/CSV File**: Use the "Browse" button to locate your data file containing the merge information.
4. **Choose Output Folder**: Select the directory where you want the generated letters to be saved.
5. **Define Custom Merge Fields**: In the provided text area, specify the merge fields in the format `ColumnName:MergeWord`.
6. **Preview (Optional)**: Click the "Preview" button to see how the letters will look before generation.
7. **Generate Letters**: Click "Generate Letters" to create the individual letters based on the data.
8. **Undo (Optional)**: If needed, use the "Undo" button to delete the last set of generated letters.

## Safety Features
The Letter Generator incorporates several safety measures to ensure a smooth user experience:
- **Input Validation**: The application checks the validity of the selected files and directories before proceeding, providing error messages for invalid paths.
- **Error Handling**: Comprehensive error handling is in place to catch and report issues during letter generation, logging errors for reference.
- **Undo Functionality**: Users can easily remove letters generated during the last run, minimizing potential mistakes.
- **Logging**: All actions and errors are logged for accountability and troubleshooting.

## How It Works
The Letter Generator operates by following these main steps:
1. **File Input**: The user selects the necessary files via the GUI. The Word template serves as the foundation for the letters, while the Excel/CSV file provides the data to fill in the custom fields.
2. **Data Processing**: The application reads the spreadsheet data and prepares it for merging. Users specify which columns to use and define corresponding merge words in the template.
3. **Dynamic Text Replacement**: During letter generation, the tool replaces the specified merge words in the template with the corresponding data from the spreadsheet, creating personalized letters.
4. **Letter Creation**: The application saves each generated letter in the designated output folder, allowing for easy access.
5. **Preview Option**: Users have the option to preview the letters before finalizing the generation, ensuring everything appears as expected.

## Conclusion
The Letter Generator tool automates the creation of personalized letters, making it an invaluable resource for anyone needing to produce bulk correspondence efficiently. Its intuitive interface, robust safety measures, and error handling capabilities ensure a smooth user experience. With features like preview mode and undo functionality, it stands out as a comprehensive solution for letter generation.
