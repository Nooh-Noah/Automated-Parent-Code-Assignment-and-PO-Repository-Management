import pandas as pd

# Path to the file
file_path = r"C:\Users\MohammadShariqKhan\OneDrive - TRESORI TRADING L.L.C\Desktop\Test Parent Code.xlsb"
output_file_path = r"C:\Users\MohammadShariqKhan\OneDrive - TRESORI TRADING L.L.C\Desktop\Updated_PO.xlsx"
repo_output_file_path = r"C:\Users\MohammadShariqKhan\OneDrive - TRESORI TRADING L.L.C\Desktop\Updated_PO_Repository.xlsx"

# Read both sheets using pandas with pyxlsb engine
po_df = pd.read_excel(file_path, sheet_name="PO", engine='pyxlsb')
repo_df = pd.read_excel(file_path, sheet_name="PO Repository", engine='pyxlsb')

# Ensure the columns are correct and max parent code is fetched
max_parent_code = repo_df["T. parent Code"].max()

# Step 1: Create a dictionary for the PO Repository based on Parent Code
# Each parent code will be associated with a list of (Item Code, Quantity) tuples
repo_combinations = {}
for parent_code, group in repo_df.groupby("T. parent Code"):
    item_combination = tuple(sorted(zip(group['T. Item Code'], group['QUANTITY'])))
    repo_combinations[item_combination] = parent_code

# List to store new repository entries
new_repo_entries = []

# Step 2: Function to find existing parent code or generate a new one
def find_or_generate_parent_code(group):
    # Create a combination of (Item Code, Quantity) for the PO group
    po_combination = tuple(sorted(zip(group['T. Item Code'], group['QUANTITY'])))
    
    # Step 3: Check if this combination exists in the PO Repository
    if po_combination in repo_combinations:
        # Return the existing parent code
        return repo_combinations[po_combination]
    else:
        # Generate a new parent code by incrementing the max parent code
        global max_parent_code
        max_parent_code += 1
        # Save this new combination in the repo_combinations for future reference
        repo_combinations[po_combination] = max_parent_code

        # Step 3.1: Prepare new rows for PO repository for this new parent code
        new_rows = pd.DataFrame({
            'Supplier Code': group['Supplier Code'],
            'Parent Code Identifier': group['Parent Code Identifier'],
            'QUANTITY': group['QUANTITY'],
            'T. Item Code': group['T. Item Code'],
            'T. parent Code': [max_parent_code] * len(group)
        })

        # Add these new rows to the list to append to the repository later
        new_repo_entries.append(new_rows)

        return max_parent_code

# Step 4: Apply the logic to the PO Sheet by grouping by 'Parent Code Identifier'
po_df = po_df.groupby("Parent Code Identifier").apply(assign_parent_code_to_group)

# Step 5: If there are any new repository entries, append them to the repository
if new_repo_entries:
    new_repo_df = pd.concat(new_repo_entries)
    repo_df = pd.concat([repo_df, new_repo_df], ignore_index=True)

# Save the updated PO to a new Excel file
po_df.to_excel(output_file_path, sheet_name="Updated PO", index=False)

# Save the updated PO Repository with new parent codes
repo_df.to_excel(repo_output_file_path, sheet_name="Updated PO Repository", index=False)

print(f"Updated PO has been saved to {output_file_path}")
print(f"Updated PO Repository has been saved to {repo_output_file_path}")
