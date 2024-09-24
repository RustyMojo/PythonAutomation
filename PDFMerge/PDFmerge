import os
import logging
import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Set up logging to record events and errors in a log file
logging.basicConfig(filename='letter_generator.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def generate_letters(template_path, excel_path, output_folder, merge_fields, is_preview=False):
    """
    Generates personalized letters by merging data from an Excel/CSV file into a Word template.

    Args:
        template_path (str): Path to the Word template file.
        excel_path (str): Path to the Excel/CSV file containing data for merging.
        output_folder (str): Directory to save the generated letters.
        merge_fields (dict): Mapping of column names to merge words for text replacement.
        is_preview (bool): If True, shows a preview instead of generating letters.
    """
    try:
        # Read data from Excel or CSV file
        df = pd.read_excel(excel_path) if excel_path.endswith('.xlsx') else pd.read_csv(excel_path)
        df = df.fillna('')  # Replace NaN values with empty strings

        if is_preview:
            # If preview mode is enabled
            preview_doc = Document(template_path)  # Load the template

            # Iterate through the first three rows of the DataFrame for preview
            for index, row in df.head(3).iterrows():
                # Replace merge words in the template with corresponding values
                for paragraph in preview_doc.paragraphs:
                    for col_name, merge_words in merge_fields.items():
                        if col_name in df.columns:
                            value = str(row[col_name])  # Get the value for the current merge field
                            for merge_word in merge_words:
                                paragraph.text = paragraph.text.replace(merge_word, value)  # Replace text

            # Display the preview in a messagebox
            preview_text = "\n".join([paragraph.text for paragraph in preview_doc.paragraphs])
            messagebox.showinfo("Preview", preview_text)
        else:
            # Generation mode
            for index, row in df.iterrows():
                doc = Document(template_path)  # Load the template for each row

                # Dictionary to store paragraphs for each unique merge case
                merge_cases = {}

                # Replace merge words in the template with corresponding values for each row
                for paragraph in doc.paragraphs:
                    for col_name, merge_words in merge_fields.items():
                        if col_name in df.columns:
                            value = str(row[col_name])  # Get the value for the current merge field
                            for merge_word in merge_words:
                                # Create a unique key for each merge case
                                merge_key = (merge_word.lower(), value.lower())

                                # Append paragraph to the corresponding merge case
                                if merge_key not in merge_cases:
                                    merge_cases[merge_key] = [paragraph.text.replace(merge_word, value)]
                                else:
                                    merge_cases[merge_key].append(paragraph.text.replace(merge_word, value))

                # Save separate letters for each unique merge case
                for merge_key, paragraphs in merge_cases.items():
                    unique_merge_word, unique_value = merge_key
                    unique_doc = Document()  # Create a new Document for each unique merge case
                    for text in paragraphs:
                        unique_doc.add_paragraph(text)  # Add paragraphs to the new document
                    
                    # Construct output file path
                    output_path = f"{output_folder}/{unique_merge_word.capitalize()}_{unique_value.replace(' ', '_')}_{index + 1}.docx"
                    unique_doc.save(output_path)  # Save the new document

            logging.info(f"Letter generation successful. Output folder: {output_folder}")
            messagebox.showinfo("Success", "Letter generation completed successfully.")

    except Exception as e:
        logging.error(f"Error during letter generation: {str(e)}")
        messagebox.showerror("Error", f"An error occurred during letter generation:\n{str(e)}")

def browse_file(entry_var, filetypes):
    """
    Opens a file dialog for the user to select a file and updates the entry variable.

    Args:
        entry_var (tk.StringVar): Variable to update with the selected file path.
        filetypes (list): List of file types to filter in the dialog.
    """
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    entry_var.set(file_path)  # Set the selected file path to the entry variable

def browse_folder(entry_var):
    """
    Opens a folder dialog for the user to select a directory and updates the entry variable.

    Args:
        entry_var (tk.StringVar): Variable to update with the selected folder path.
    """
    folder_path = filedialog.askdirectory()  # Open a dialog to select a directory
    entry_var.set(folder_path)  # Set the selected folder path to the entry variable

def run_generation(template_entry, excel_entry, output_entry, fields_entry, is_preview=False):
    """
    Initiates the letter generation process.

    Args:
        template_entry (tk.Entry): Entry widget for the template file path.
        excel_entry (tk.Entry): Entry widget for the Excel/CSV file path.
        output_entry (tk.Entry): Entry widget for the output folder path.
        fields_entry (tk.Text): Text widget for custom merge fields.
        is_preview (bool): If True, runs in preview mode.
    """
    template_path = template_entry.get()  # Get the template file path
    excel_path = excel_entry.get()  # Get the Excel/CSV file path
    output_folder = output_entry.get()  # Get the output folder path

    # Get all text from the fields text widget
    fields_text = fields_entry.get("1.0", tk.END)

    # Parse custom merge fields into a dictionary
    merge_fields = {}
    for line in fields_text.split('\n'):
        parts = line.split(':')  # Split by colon
        if len(parts) == 2:
            col_name = parts[0].strip()  # Column name
            merge_word = parts[1].strip()  # Merge word

            # If the merge_word is not empty, add it to merge_fields
            if merge_word:
                if col_name in merge_fields:
                    merge_fields[col_name].append('{{' + merge_word + '}}')  # Format the merge word
                else:
                    merge_fields[col_name] = ['{{' + merge_word + '}}']  # Initialize a new list for this column
    generate_letters(template_path, excel_path, output_folder, merge_fields, is_preview)  # Call the main function

def prefill_custom_fields(excel_entry, fields_entry):
    """
    Prefills the custom merge fields based on the columns in the selected Excel/CSV file.

    Args:
        excel_entry (tk.Entry): Entry widget for the Excel/CSV file path.
        fields_entry (tk.Text): Text widget to display the custom merge fields.
    """
    excel_path = excel_entry.get()  # Get the Excel/CSV file path
    if excel_path:
        df = pd.read_excel(excel_path) if excel_path.endswith('.xlsx') else pd.read_csv(excel_path)
        columns = df.columns  # Get the columns from the DataFrame
        # Set the columns as custom fields in the Text widget
        fields_entry.delete("1.0", tk.END)  # Clear the text area
        fields_entry.insert("1.0", "\n".join([f"{col}:" for col in columns]))  # Populate with column names

def undo_last(fields_entry, output_folder, excel_entry):
    """
    Deletes the files generated in the last run to undo the letter generation.

    Args:
        fields_entry (tk.Text): Text widget for custom merge fields.
        output_folder (str): Path to the output folder where letters were saved.
        excel_entry (tk.Entry): Entry widget for the Excel/CSV file path.
    """
    try:
        # Undo by deleting all the files generated in the last run
        files_to_delete = [f"{output_folder}/Letter_{i + 1}.docx" for i in range(len(pd.read_excel(excel_entry.get())))]

        for file_path in files_to_delete:
            if os.path.exists(file_path):  # Check if the file exists
                os.remove(file_path)  # Delete the file
                logging.info(f"Undone generation. Deleted file: {file_path}")
            else:
                logging.warning(f"Undone generation. File not found: {file_path}")

        messagebox.showinfo("Undo", "Undo operation completed successfully.")

    except Exception as e:
        logging.error(f"Error during undo operation: {str(e)}")
        messagebox.showerror("Error", f"An error occurred during the undo operation:\n{str(e)}")

def validate_paths(template_path, excel_path, output_folder):
    """
    Validates the file and folder paths to ensure they exist.

    Args:
        template_path (str): Path to the Word template file.
        excel_path (str): Path to the Excel/CSV file.
        output_folder (str): Path to the output folder.

    Returns:
        bool: True if all paths are valid, False otherwise.
    """
    # Validate paths and show appropriate messages
    if not os.path.isfile(template_path):
        messagebox.showerror("Error", "Invalid template file path.")
        return False
    elif not os.path.isfile(excel_path):
        messagebox.showerror("Error", "Invalid Excel file path.")
        return False
    elif not os.path.isdir(output_folder):
        messagebox.showerror("Error", "Invalid output folder path.")
        return False
    else:
        return True  # All paths are valid
    
def validate_and_run(template_entry, excel_entry, output_entry, fields_entry):
    """
    Validates paths and initiates the letter generation process.

    Args:
        template_entry (tk.Entry): Entry widget for the template file path.
        excel_entry (tk.Entry): Entry widget for the Excel/CSV file path.
        output_entry (tk.Entry): Entry widget for the output folder path.
        fields_entry (tk.Text): Text widget for custom merge fields.
    """
    template_path = template_entry.get()  # Get the template file path
    excel_path = excel_entry.get()  # Get the Excel/CSV file path
    output_folder = output_entry.get()  # Get the output folder path

    if validate_paths(template_path, excel_path, output_folder):
        # Get all text from the fields text widget
        fields_text = fields_entry.get("1.0", tk.END)

        # Parse custom merge fields
        merge_fields = {}
        for line in fields_text.split('\n'):
            parts = line.split(':')
            if len(parts) == 2:
                col_name = parts[0].strip()  # Column name
                merge_word = parts[1].strip()  # Merge word

                # If the merge_word is not empty, add it to merge_fields
                if merge_word:
                    if col_name in merge_fields:
                        merge_fields[col_name].append('{{' + merge_word + '}}')  # Format the merge word
                    else:
                        merge_fields[col_name] = ['{{' + merge_word + '}}']  # Initialize a new list for this column

        generate_letters(template_path, excel_path, output_folder, merge_fields)  # Call the main function

def create_gui():
    """
    Creates the graphical user interface for the Letter Generator application.
    """
    root = tk.Tk()  # Create the main window
    root.title("Letter Generator")  # Set the window title

    # Style configuration for buttons and entries
    style = ttk.Style()
    style.configure('TButton', padding=5, width=20)
    style.configure('TEntry', padding=(5, 5, 5, 5), width=30)

    # Template path entry
    tk.Label(root, text="Template Letter:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.E)
    template_var = tk.StringVar()
    template_entry = ttk.Entry(root, textvariable=template_var, style='TEntry')
    template_entry.grid(row=0, column=1, padx=10, pady=10, columnspan=2, sticky=tk.W)
    ttk.Button(root, text="Browse", command=lambda: browse_file(template_var, [("Word files", "*.docx"), ("All files", "*.*")])).grid(row=0, column=3, padx=10, pady=10, sticky=tk.W)

    # Excel path entry
    tk.Label(root, text="Excel Sheet:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.E)
    excel_var = tk.StringVar()
    excel_entry = ttk.Entry(root, textvariable=excel_var, style='TEntry')
    excel_entry.grid(row=1, column=1, padx=10, pady=10, columnspan=2, sticky=tk.W)
    ttk.Button(root, text="Browse", command=lambda: browse_file(excel_var, [("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")])).grid(row=1, column=3, padx=10, pady=10, sticky=tk.W)
    ttk.Button(root, text="Prefill Custom Fields", command=lambda: prefill_custom_fields(excel_entry, fields_entry)).grid(row=1, column=4, padx=10, pady=10, sticky=tk.W)

    # Output folder entry
    tk.Label(root, text="Output Folder:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.E)
    output_var = tk.StringVar()
    output_entry = ttk.Entry(root, textvariable=output_var, style='TEntry')
    output_entry.grid(row=2, column=1, padx=10, pady=10, columnspan=2, sticky=tk.W)
    ttk.Button(root, text="Browse", command=lambda: browse_folder(output_var)).grid(row=2, column=3, padx=10, pady=10, sticky=tk.W)

    # Custom merge fields entry
    tk.Label(root, text="Custom Merge Fields (Column:Merge Word):").grid(row=3, column=0, padx=10, pady=10, sticky=tk.E)
    fields_var = tk.StringVar()
    fields_entry = tk.Text(root, height=8, width=30)  # Text area for entering custom fields
    fields_entry.grid(row=3, column=1, padx=10, pady=10, columnspan=2, sticky=tk.W)

    # Undo button
    ttk.Button(root, text="Undo", command=lambda: undo_last(fields_entry, output_entry.get(), excel_entry)).grid(row=4, column=1, padx=10, pady=10, sticky=tk.W)

    # Preview button
    ttk.Button(root, text="Preview", command=lambda: run_generation(template_entry, excel_entry, output_entry, fields_entry, True)).grid(row=4, column=2, padx=10, pady=10, sticky=tk.E)

    # Generate Letters button
    ttk.Button(root, text="Generate Letters", command=lambda: validate_and_run(template_entry, excel_entry, output_entry, fields_entry)).grid(row=4, column=3, padx=10, pady=10, sticky=tk.W)

    root.mainloop()  # Start the GUI event loop

if __name__ == "__main__":
    create_gui()  # Run the GUI application
