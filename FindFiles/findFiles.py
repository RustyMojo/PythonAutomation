import os
import shutil

def list_files_in_directory(directory):
    """
    List all files in a given directory, including subdirectories.

    Args:
        directory (str): The path to the directory to search.

    Returns:
        list: A list of file names found in the directory.
    """
    file_list = []  # Initialize an empty list to store file names
    # Walk through the directory and its subdirectories
    for root, _, files in os.walk(directory):
        for file in files:
            file_list.append(file)  # Store just the file names, not the full paths
    return file_list

def find_missing_files(source_dir, dest_dir):
    """
    Identify files that exist in the source directory but are missing in the destination directory.

    Args:
        source_dir (str): The path to the source directory.
        dest_dir (str): The path to the destination directory.

    Returns:
        list: A list of missing file names.
    """
    # Get sets of files from both directories
    source_files = set(list_files_in_directory(source_dir))
    dest_files = set(list_files_in_directory(dest_dir))
    # Find files that are in source but not in destination
    missing_files = source_files - dest_files
    return list(missing_files)

def copy_files_with_preview(source_dir, preview_dir, missing_files):
    """
    Copy missing files from the source directory to a preview directory, renaming them if needed.

    Args:
        source_dir (str): The path to the source directory.
        preview_dir (str): The path to the preview directory where files will be copied.
        missing_files (list): A list of missing file names to copy.
    """
    for file_name in missing_files:
        source_file_path = os.path.join(source_dir, file_name)  # Full path to the source file
        # Create a unique preview file name
        preview_file_name = f"short_{file_name}"
        preview_file_path = os.path.join(preview_dir, preview_file_name)  # Full path for the preview file

        # Copy the file from the source directory to the preview directory
        shutil.copy2(source_file_path, preview_file_path)

if __name__ == "__main__":
    # Get directory paths from the user
    source_dir = input("Enter the source directory path: ")
    dest_dir = input("Enter the destination directory path: ")
    preview_dir = input("Enter the preview directory path: ")

    # Find missing files and store them in a list
    missing_files = find_missing_files(source_dir, dest_dir)

    # Display the missing files
    print("Files in source directory that are not in the destination directory:")
    for file_name in missing_files:
        print(file_name)

    # Copy the missing files to the preview directory
    copy_files_with_preview(source_dir, preview_dir, missing_files)
