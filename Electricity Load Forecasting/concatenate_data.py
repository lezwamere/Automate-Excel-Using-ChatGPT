import pandas as pd

# Define the Excel file name
excel_file = 'test_dataframes.xlsx'

# Create an empty DataFrame to store the concatenated data
concatenated_data = pd.DataFrame()

# List of worksheet names you want to concatenate
worksheet_names = ["Week 15, Apr 2019", "Week 21, May 2019", "Week 24, Jun 2019", "Week 29, July 2019", "Week 33, Aug 2019"]

# Iterate through each worksheet and append its data to the concatenated_data DataFrame
for sheet_name in worksheet_names:
    # Read data from the current worksheet
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Append the data to the concatenated_data DataFrame
    concatenated_data = concatenated_data.append(df, ignore_index=True)

# Create a new Excel writer object
with pd.ExcelWriter('concatenated_data.xlsx', engine='xlsxwriter') as writer:
    # Write the concatenated data to a new worksheet
    concatenated_data.to_excel(writer, sheet_name='Concatenated Data', index=False)

print("Data from the worksheets has been concatenated and saved to 'concatenated_data.xlsx'")

# Handling errors fro missing data. Define the Excel file name
excel_file = 'test_dataframes.xlsx'

# Create an empty DataFrame to store the concatenated data
concatenated_data = pd.DataFrame()

# List of worksheet names you want to concatenate
worksheet_names = ["Week 15, Apr 2019", "Week 21, May 2019", "Week 24, Jun 2019", "Week 29, July2019", "Week 33, Aug"]

# Iterate through each worksheet and append its data to the concatenated_data DataFrame
for sheet_name in worksheet_names:
    # Read data from the current worksheet
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Append the data to the concatenated_data DataFrame
    concatenated_data = concatenated_data.append(df, ignore_index=True)

# Create a new Excel writer object
with pd.ExcelWriter('concatenated_data.xlsx', engine='xlsxwriter') as writer:
    # Write the concatenated data to a new worksheet
    concatenated_data.to_excel(writer, sheet_name='Concatenated Data', index=False)

print("Data from the worksheets has been concatenated and saved to 'concatenated_data.xlsx'")
