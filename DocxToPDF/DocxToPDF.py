import os
import docx2pdf

def convert_docx_to_pdf(input_dir, output_dir):
    """
    Converts all .docx files in a directory to PDFs and saves them in a new folder.

    Args:
        input_dir (str): Path to the directory containing the .docx files.
        output_dir (str): Path to the new directory for the converted PDFs.
    """

    # Create the new output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)  # Handles existing directories gracefully

    # Iterate through each file in the input directory
    for filename in os.listdir(input_dir):
        # Print the current file being processed
        print(filename)
        # Check if the file has a .docx extension
        if filename.endswith(".docx"):
            # Construct full paths for input and output files
            input_file = os.path.join(input_dir, filename)
            output_file = os.path.join(output_dir, os.path.splitext(filename)[0] + ".pdf")

            # Convert the .docx file to .pdf
            try:
                docx2pdf.convert(input_file, output_file)
                print(f"Converted '{filename}' to PDF successfully.")
            except Exception as e:
                print(f"Error converting '{filename}': {e}")

def rename_docx_files(directory):
    """
    Renames all files with the .DOCX extension to lowercase .docx in a directory.

    Args:
        directory (str): Path to the directory containing the files.
    """
    # Iterate through each file in the specified directory
    for filename in os.listdir(directory):
        # Check if the file has a .DOCX extension (uppercase)
        if filename.endswith(".DOCX"):
            # Construct the new filename with lowercase extension
            new_filename = os.path.splitext(filename)[0] + ".docx"
            old_path = os.path.join(directory, filename)
            new_path = os.path.join(directory, new_filename)

            # Rename the file to the new lowercase filename
            os.rename(old_path, new_path)
            print(f"Renamed '{filename}' to '{new_filename}'.")

if __name__ == "__main__":
    # Define the input directory containing the .docx files
    input_dir = r""  # Replace with your actual directory path

    # Generate a descriptive output directory name for the converted PDFs
    output_dir = os.path.join(input_dir, "Converted_PDFs")  # Customize as needed

    # Define the directory for renaming .DOCX files
    directory = r""

    # Call the function to rename .DOCX files to .docx
    rename_docx_files(directory)

    # Call the function to convert .docx files to .pdf
    convert_docx_to_pdf(input_dir, output_dir)
