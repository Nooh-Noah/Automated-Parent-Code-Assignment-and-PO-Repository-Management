# Automated-Parent-Code-Assignment-and-PO-Repository-Management
This Python script automates the process of assigning parent codes to groups of items in a Purchase Order (PO) system. It identifies if a combination of item codes and their respective quantities already exists in the PO repository. If it exists, it reuses the existing parent code; if not, it generates a new parent code, assigns it to the group, and adds the new combination to the repository.

# Features
Existing Parent Code Reuse: Automatically detects and reuses the existing parent code if a combination of item codes and quantities is found in the repository.
New Parent Code Generation: Generates new parent codes by incrementing the highest available parent code when a new combination of item codes and quantities is detected.
Repository Update: Automatically appends new combinations of item codes and quantities along with the newly generated parent codes into the PO repository, ensuring that the repository stays up-to-date.
Excel-Based System: Both the Purchase Order data and the repository data are handled in Excel, making it convenient to use in a corporate environment.

# Requirements
To run this script, you'll need:

Python 3.x
pandas library for data manipulation
pyxlsb library for reading .xlsb files if needed
Install the dependencies with the following command:

bash
Copy code
pip install pandas pyxlsb

# Usage
1. Prepare the Input Files:
  > Ensure you have two Excel sheets ready: one for the Purchase Order (PO) and another for the PO Repository.
  > The script expects the following columns in both sheets:
      Supplier Code
      Parent Code Identifier
      QUANTITY
      T. Item Code
      T. parent Code
2. Script Execution:
  > Update the file paths for the PO and PO Repository in the script.
  > Run the script. It will:
      Assign existing parent codes to matching groups in the PO.
      Generate new parent codes for non-existing groups and append them to the PO Repository.
      The updated PO and PO Repository will be saved as new Excel files.

# Code Structure:

  The script reads the PO and PO Repository from Excel.
  It creates combinations of item codes and quantities to detect if a group already exists in the repository.
  If a match is found, the existing parent code is reused. If no match is found, a new parent code is generated and added to both the PO and PO Repository.
  The updated data is saved to new Excel files.

# File Structure
  Test Parent Code.xlsb – The input Excel file containing the original PO and PO Repository data.
  Updated_PO.xlsx – The output Excel file containing the PO data with updated parent codes.
  Updated_PO_Repository.xlsx – The output Excel file containing the PO Repository with new rows appended for newly generated parent codes.


# How to Use
Set the file paths:
  Update the file paths in the script for the input PO and PO Repository files.
  Set the output file paths for the updated PO and PO Repository.
Run the script:
  The script will automatically assign parent codes and save the updated files.

# Check the output:
The updated PO and PO Repository will be saved in the specified output paths.

# Path to the file
file_path = r"C:\path\to\Test Parent Code.xlsb"
output_file_path = r"C:\path\to\Updated_PO.xlsx"
repo_output_file_path = r"C:\path\to\Updated_PO_Repository.xlsx"
After running the script, two new files will be generated:

Updated_PO.xlsx – with the assigned parent codes.
Updated_PO_Repository.xlsx – with new rows appended for newly generated parent codes.


# Contributing
Feel free to submit pull requests or suggest improvements! If you encounter any issues or have suggestions for new features, feel free to open an issue.

# License
This project is open-source and available under the MIT License.

# Author
Name: [Mohammad Shariq]
Company: [Tresori Trading]
Contact: [sheri.khan88@gmail.com]


