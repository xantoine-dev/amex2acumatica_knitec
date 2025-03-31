import pandas as pd
import openpyxl
import xlrd
import sys
import os
import re
from pathlib import Path

# Function to ask the user for a valid file_path using pathlib
def get_valid_file_path(prompt):
    while True:
        file_path = Path(input(prompt).strip().strip('"').replace("\\", "/"))  # Use Path for better handling
        if file_path.exists() and file_path.is_file():  # Check if it's a valid file
            print(f"File exists. File path saved as: {file_path}!")
            return file_path
        with open(file_path, 'r'):
            print(f"The file does not exist or is not a valid file. Please try again. filepath is {file_path}")
#The function will handle different file types like CSV, Excel, JSON, and text files.
def file_to_dataframe(file_path):
    """
    Converts a file into a pandas DataFrame.

    Parameters:
        file_path (str): The path to the file to be converted.

    Returns:
        pd.DataFrame: A DataFrame containing the file data.
    """
    # Get the file extension
    file_extension = os.path.splitext(file_path)[-1].lower()

    try:
        if file_extension == '.csv':
            # Read CSV file
            df = pd.read_csv(file_path, header=None)
        elif file_extension in ['.xls', '.xlsx']:  # Excel files
            # Read Excel file
            df = pd.read_excel(file_path, header=None)
        elif file_extension == '.json':
            # Read JSON file
            df = pd.read_json(file_path, header=None)
        elif file_extension == '.txt':
            # Read text file assuming tab-delimited data
            df = pd.read_csv(file_path, delimiter='\t', header=None)
        else:
            raise ValueError(f"Unsupported file type: {file_extension}")
        # Check if the DataFrame is empty
        if df.empty:
            raise ValueError("The file is empty or contains no data.")
        # Print the shape of the DataFrame
        print(f"DataFrame shape: {df.shape}")
        # Print the first few rows of the DataFrame
        print("DataFrame preview:")
        print(df.head())
        # Convert the first column to string and remove leading/trailing whitespace
        df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.strip()
        # Convert the second column to string and remove leading/trailing whitespace
        df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip()
        # Convert the third column to string and remove leading/trailing whitespace
        df.iloc[:, 2] = df.iloc[:, 2].astype(str).str.strip()
        # Convert the fourth column to string and remove leading/trailing whitespace
        df.iloc[:, 3] = df.iloc[:, 3].astype(str).str.strip()
        # Convert the fifth column to string and remove leading/trailing whitespace
        df.iloc[:, 4] = df.iloc[:, 4].astype(str).str.strip()
        # Convert the sixth column to string and remove leading/trailing whitespace
        df.iloc[:, 5] = df.iloc[:, 5].astype(str).str.strip()
        return df

    except Exception as e:
        print(f"An error occurred while processing the file: {e}")
        return None
# Converts $numbers to floats        
def to_float(val):
    try:
        # Remove "$" and any surrounding whitespace, then convert to float
        return float(str(val).replace("$", "").strip())
    except (ValueError, TypeError):
        return None  # Return None (NaN) if the value cannot be converted
#saved function for importing Excel formats and using PyInstaller
def import_excel_with_openpyxl():
    # Get the directory where the executable is located
    if getattr(sys, 'frozen', False):  # Check if running as a PyInstaller executable
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    # Path to the template file
    template_path = os.path.join(base_path, "Oct_2024_Corp_Expense_Claim_Template.xlsx")

    # Load the template as template_df
    try:
        template_df = pd.read_excel(template_path, engine="openpyxl")
        print("Template file loaded successfully.")
    except Exception as e:
        print(f"Error loading the Template file: {e}")
        print("Let's try locating it manually.")
        template_path = get_valid_file_path(
            "Drag and drop the Template file (.xlsx) here and press Enter: "
        )
        try:
            template_df = pd.read_excel(template_path, engine="openpyxl")
        except Exception as second_error:
            print(f"Failed to load the Template file: {second_error}")
            sys.exit(1)
# Function to get and validate directory paths
def get_valid_directory_path(prompt):
    while True:
        directory_path = input(prompt).strip().strip('"')  # Strip spaces and quotes
        if os.path.isdir(directory_path):
            print(f"directory path validated as {directory_path}")
            return directory_path
        print("The directory does not exist. Please try again.")
#reads in the excel amex file at a given statement_path, cleaning the file's amount column data. Returns data as statement_df.
def read_in_excel_and_clean(statement_path, amount_column):
    while True:
        try:
            #read statement in; skipping to the rows with data.
            statement_df = file_to_dataframe(statement_path) #Saving input file as statement_df.
            statement_df = statement_df.reset_index(drop=True)  # Reset index and drop the old one
            statement_df.columns = statement_df.iloc[13]  # Set the 13th row as the header
            statement_df = statement_df.iloc[14:] #slicing off the top 12 rows of statement_df.
            if statement_df.empty:
                print("Error: The statement file is empty. Please check the file and try again.")
                sys.exit(1)
            print(f"⭐️ statement_df loaded successfully ⭐️ {statement_df.shape[0]} rows and {statement_df.shape[1]} columns.")
            #check if the amount column exists
            if amount_column not in statement_df.columns:
                print(f"🚩 Error: '{amount_column}' column not found in the statement file.")
                print(f"Available columns: {statement_df.columns.tolist()}")
                print(statement_df.head())
                # Convert the first column to string and remove leading/trailing whitespace
                statement_df.iloc[:, 0] = statement_df.iloc[:, 0].astype(str).str.strip()
                # Convert the second column to string and remove leading/trailing whitespace
                statement_df.iloc[:, 1] = statement_df.iloc[:, 1].astype(str).str.strip()
                # Convert the third column to string and remove leading/trailing whitespace
                statement_df.iloc[:, 2] = statement_df.iloc[:, 2].astype(str).str.strip()
                # Convert the fourth column to string and remove leading/trailing whitespace
                statement_df.iloc[:, 3] = statement_df.iloc[:, 3].astype(str).str.strip()
                # Convert the fifth column to string and remove leading/trailing whitespace
                statement_df.iloc[:, 4] = statement_df.iloc[:, 4].astype(str).str.strip()
                # Convert the sixth column to string and remove leading/trailing whitespace
                statement_df.iloc[:, 5] = statement_df.iloc[:, 5].astype(str).str.strip()
                sys.exit(1)
            #check if the amount column is empty
            if statement_df[amount_column].isnull().all():
                print(f"Error: '{amount_column}' column is empty.")
                print(f"Available columns: {statement_df.columns.tolist()}")
                sys.exit(1)
            total_rows_before = len(statement_df[amount_column])
            print(f"total rows before converting to string: {total_rows_before}")

            #check if column is saved as an 'object' data type, this is pandas' way of identifying a string. If it's a string, remove $ and , to clean.
            column_data_type = statement_df[amount_column].dtype
            if column_data_type == 'object':
                #str.replace() leaves data untouched if pattern is not found
                statement_df[amount_column] = (
                statement_df[amount_column]
                .str.replace("$", "", regex=False)
                .str.replace(",", "", regex=False)
                )
                statement_df[amount_column] = pd.to_numeric(statement_df[amount_column], errors='coerce')
            total_rows_after = len(statement_df[amount_column])
            print(f"total rows after converting to string: {total_rows_after}")

            total_sum_before = statement_df[amount_column].sum()
            print(f"total before filtering out rows with negative values: {total_sum_before}")
            # Count and print the number of rows with negative values before filtering
            negative_count_before = (statement_df[amount_column] < 0).sum()
            print(f"Number of rows with negative values in the amount column before filtering: {negative_count_before}")

            # Exclude rows with negative values
            statement_df = statement_df[statement_df[amount_column] >= 0]

            # Count and print the number of rows with negative values after filtering
            negative_count_after = (statement_df[amount_column] < 0).sum()
            print(f"Number of rows with negative values after filtering: {negative_count_after}")
            total_sum_after = statement_df[amount_column].sum()
            print(f"total after filtering: {total_sum_after}")

            print(f"---Amex statement file loaded successfully :)")
            return statement_df
        except Exception as e:
            print(f"Error reading the Amex statement file: {e}")

###stores the filepath of the Amex statement
statement_path = get_valid_file_path("""----------
Hello,

Please drag and drop, or paste the Amex file path here:

""")

amount_column = "Transaction \nAmount \nUSD"

###stores statement_df as a cleaned_statement_df.
cleaned_statement_df = read_in_excel_and_clean(statement_path, amount_column)

length_of_amount_column = len(cleaned_statement_df[amount_column])
print("""

number of rows after cleaning:""")
print(str(length_of_amount_column))

# Define a function to remove numbers from a string
def remove_numbers_from_string(val):
    if isinstance(val, str):  # Ensure the value is a string
        return re.sub(r'\d+', '', val)  # Replace all digits with an empty string
    return val  # Return the value unchanged if not a string

# Apply this function to the column
cleaned_statement_df['Transaction \nDescription 4'] = cleaned_statement_df['Transaction \nDescription 4'].apply(remove_numbers_from_string)

# Define the structure of the output as an empty DataFrame called "template_df". We will add the values from the statement to the template_df file.
template_columns = [
    "Branch", "Date", "Ref. Nbr.", "Expense Item", "Expense Account",
    "Description", "Amount", "Claim Amount", "Paid With",
    "Corporate Card", "AR Reference Nbr."
]
template_df = pd.DataFrame(columns=template_columns)

# Now the output template is created as template_df and the statement is loaded as cleaned_statement_df (rows w/ negative values excluded)

# Ask the user to select the directory to save output files
output_directory = get_valid_directory_path("Enter the directory path where you want to save the output files. You can type it out or drag the folder and your computer should put the correct directory path for processing.:\n")

# Now to create the outputs, creating a sheet per last name within the Supplemental \nCardmember Last \nName column in the Amex statement.

# Define the exact name for "Last Name" column in the statement file
last_name_column = "Supplemental \nCardmember Last \nName"
# Check if the "Last Name" column exists
if last_name_column not in cleaned_statement_df.columns:
    print(f"Error: '{last_name_column}' column not found in cleaned_statement_df.")
    exit(1)

# Identify unique last names
unique_last_names = cleaned_statement_df[last_name_column].unique()
print(f"Found {len(unique_last_names)} unique last names in '{last_name_column}' column.")

# Define the columns to be used for filling the template. These should exist in the  statement file.
required_columns = {
    "Transaction Date": "Date",
    "Transaction \nAmount \nUSD": ["Amount", "Claim Amount"],
    "Transaction \nDescription 1": "Description",
    "Transaction \nDescription 4": "Ref. Nbr."
}

#This variable will be a list of our output file names.
output_files = []

export_format = input("Enter the export format ('csv' or 'excel'): ").strip().lower()

# Validate the user's input
if export_format not in ["csv", "excel"]:
    raise ValueError("Invalid format. Please enter 'csv' or 'excel'.")

for last_name in unique_last_names:
    # Filter the statement data for the current last name
    employee_data = cleaned_statement_df[cleaned_statement_df[last_name_column] == last_name]

    # Create a copy of the template
    employee_template = template_df.copy()
    
    # Populate the template with the employee's data
    for index, row in employee_data.iterrows():
        # Initialize a new row with all template columns set to empty
        new_row = {col: "" for col in employee_template.columns}

        # Fill only the specified columns with data from the statement
        new_row["Date"] = row["Transaction Date"]
        new_row["Ref. Nbr."] = row["Transaction \nDescription 4"]
        new_row["Description"] = row["Transaction \nDescription 1"]
        new_row["Amount"] = row["Transaction \nAmount \nUSD"]
        new_row["Claim Amount"] = row["Transaction \nAmount \nUSD"]
        new_row["Paid With"] = "Corporate Card, Company Expense"
        new_row["Branch"] = "KEC"

        # Append this row to the employee's template DataFrame
        employee_template = pd.concat([employee_template, pd.DataFrame([new_row])], ignore_index=True)
    # Define the output file path within the specified directory, using the last name
    if export_format == "excel":
        output_file = os.path.join(output_directory, f"{last_name}_AMEX_Claim.xlsx")
        employee_template.to_excel(output_file, index=False)
    elif export_format == "csv":
        output_file = os.path.join(output_directory, f"{last_name}_AMEX_Claim.csv")
        employee_template.to_csv(output_file, index=False)
    print(f"Generated file for {last_name}: {output_file} in {output_directory}")
    print(f"""----------
    The files have been processed and saved to {output_directory}!
    """)

# Process each last name
#for last_name in unique_last_names:
    # Filter the statement data for the current last name
    employee_data = cleaned_statement_df[cleaned_statement_df[last_name_column] == last_name]

    # Create a copy of the template
    employee_template = template_df.copy()
    
    # Populate the template with the employee's data
    for index, row in employee_data.iterrows():
        # Initialize a new row with all template columns set to empty
        new_row = {col: "" for col in employee_template.columns}

        # Fill only the specified columns with data from the statement
        new_row["Date"] = row["Transaction Date"]
        new_row["Ref. Nbr."] = row["Transaction \nDescription 4"]
        new_row["Description"] = row["Transaction \nDescription 1"]
        new_row["Amount"] = row["Transaction \nAmount \nUSD"]
        new_row["Claim Amount"] = row["Transaction \nAmount \nUSD"]
        new_row["Paid With"] = "Corporate Card, Company Expense"
        new_row["Branch"] = "KEC"

        # Append this row to the employee's template DataFrame
        employee_template = pd.concat([employee_template, pd.DataFrame([new_row])], ignore_index=True)

    # Define the output file path within the specified directory, using the last name
    output_file = os.path.join(output_directory, f"{last_name}_AMEX_Claim.xlsx")
    employee_template.to_excel(output_file, index=False)
    output_files.append(output_file)
    print(f"Generated file for {last_name}: {output_file} in {output_directory}")
    
# Check for discrepancies, bug testing.
def bug_testing_row_counts():
    for last_name in unique_last_names:
        # Filter the original statement for this last name
        original_data = cleaned_statement_df[cleaned_statement_df[last_name_column] == last_name]
        
        # Load the corresponding output file
        output_file = os.path.join(output_directory, f"{last_name}_AMEX_Claim.xlsx")
        output_data = pd.read_excel(output_file)

        # Compare row counts
        if original_data.shape[0] == output_data.shape[0]:
            print(f"File {output_file} matches the original data for {last_name}.")
        else:
            print(f"Discrepancy found in {output_file}:")
            print(f"Original rows: {original_data.shape[0]}, Output rows: {output_data.shape[0]}")

    # Display generated files
    print("\nFiles generated:")
    for file in output_files:
        print(file)
    # Total rows in the original statement after dropping NaN values
    original_row_count = cleaned_statement_df.shape[0]
    # Total rows in the generated files
    output_row_count = 0
    for file in output_files:
        df = pd.read_excel(file)
        output_row_count += df.shape[0]
    print(f"Original Row Count: {original_row_count}")
    print(f"Output Row Count: {output_row_count}")
    if original_row_count == output_row_count:
        print("All rows accounted for in the output files!")
    else:
        print("Mismatch detected: Some rows might be missing or duplicated.")

    print(f"---The sheets per employee last name have been generated and placed in {output_directory}.")

# Load the Employee Corporate Card Number Excel file
def load_employee_corporate_card():
    corporate_card_file_path = get_valid_file_path("Drag and drop the Excel file here or type the file path and press Enter: ")
    # Check if the file exists
    if not os.path.exists(corporate_card_file_path):
        print("Error: The specified file does not exist. Please check the file path and try again.")
        exit(1)
    #Create dataframe as corporate_card_df
    corporate_card_df = pd.read_excel(corporate_card_file_path)
    # Ensure there are at least two columns
    if corporate_card_df.shape[1] < 2:
        print("Error: The file must have at least two columns.")
        exit(1)
    # Combine the first and second column values, per client request, and putting them into a new column in the dataframe corporate_card_df as corporate_card_df['Combined'].
    corporate_card_df['Combined'] = corporate_card_df.iloc[:, 0].astype(str) + " - " + corporate_card_df.iloc[:, 1].astype(str)
    # Print the resulting DataFrame for verification
    ##print("Combined Column Values:")
    ##print(corporate_card_df[['Combined']])
    
    #Updating corporate_card_df['Last Name'] column
    corporate_card_df['Last Name'] = corporate_card_df['Combined'].str.split().str[-1].str.lower()
    print("Corporate card file loaded and last names extracted.")
    return corporate_card_df

while True:
    option_load_employee_corporate_card = input("Do you want to load the employee corporate card file (from Acumatica)? (yes/no). This will add the corporate card columns.: ").strip().lower()

    if option_load_employee_corporate_card == "yes":
        try:
            corporate_card_df = load_employee_corporate_card()
            print("Corporate card file loaded successfully.")
            break  # Exit the loop on success
        except Exception as e:
            print(f"An error occurred while loading: {e}")
            retry = input("Loading failed. Do you want to try again? (yes/no): ").strip().lower()
            if retry != "yes":
                print("Exiting without loading the corporate card.")
                break  # Exit the loop if the user doesn't want to retry
    elif option_load_employee_corporate_card == "no":
        print("Okay, the function will not be executed.")
        break  # Exit the loop if the user doesn't want to proceed
    else:
        print("Invalid input. Please type 'yes' or 'no'.")

# Update the "Corporate Card" column in each output file
for file in output_files:
    # Extract employee's last name from the file name
    employee_last_name = os.path.basename(file).replace("_AMEX_Claim.xlsx", "").lower()

    # Match the employee's last name with the extracted last name in the corporate card file
    matched_row = corporate_card_df[corporate_card_df['Last Name'] == employee_last_name]

    # Load the employee's output file
    employee_df = pd.read_excel(file)

    if not matched_row.empty:
        # Fill the "Corporate Card" column with the matched value from 'Combined'
        corporate_card_value = matched_row['Combined'].iloc[0]
        employee_df['Corporate Card'] = corporate_card_value
    else:
        print(f"Warning: No corporate card found for {employee_last_name}.")
        employee_df['Corporate Card'] = "Not Assigned"

    # Save the updated file
    employee_df.to_excel(file, index=False)
    print(f"Updated 'Corporate Card' column for {employee_last_name} in {file}.")

print("⭐️ Corporate Card updates completed ⭐️")