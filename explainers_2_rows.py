"""
This module processes an Excel file and writes the output to a new sheet.
"""
import os
import pandas as pd
import numpy as np
from openpyxl.styles import Alignment
# pylint: disable=invalid-name

input_sheet_name = 'Input'
# rest of your code

# path to your Excel file.
xlsx_path = os.path.abspath('./2024-03-25b Explainers.xlsx')
workbook = pd.ExcelFile(xlsx_path)

# Read the input sheet. Replace 'Input' with the actual name of your input sheet.
input_sheet_name = 'Input'
input_df = workbook.parse(input_sheet_name, header=None)

# Initialize the list that will be used to create the DataFrame for the output sheet.
output_data_list = []
output_data = {}
cs_counter = 1

# Iterate over the rows of the input DataFrame and extract data based on labels in column B.
for _, row in input_df.iterrows():
    label = row[1]  # Label in column B
    text = row[2]  # Corresponding text in column C

    if pd.notna(label):
        if label == 'Title':
            # If this is not the first Title, append the previous set to the list
            if output_data:
                output_data_list.append(output_data)
                output_data = {}  # Start a new set
            output_data['Title'] = text
        elif label == 'TB':
            output_data['Trustie Bot says'] = text
        elif label == 'CS':
            cs_key = f'Conversation Starter {cs_counter}'
            output_data[cs_key] = text
            cs_counter += 1
        elif label == 'CTA':
            output_data['Call to Action'] = text
            cs_counter = 1  # Reset the counter for conversation starters

# Append the last set to the list
if output_data:
    output_data_list.append(output_data)

# print ("output_data_list: ", output_data_list)

output_sheet_name = 'Output'


# Initialize a list to hold all DataFrames
df_list = []
df_empty = pd.DataFrame(index=range(2), columns=['Key', 'Value'])

for output_data in output_data_list:
    # Convert the dictionary to a DataFrame with the keys as one column and the values as another column
    df = pd.DataFrame(list(output_data.items()), columns=['Key', 'Value'])

    # Append the DataFrame to df_list
    df_list.append(df)

    # Append an empty DataFrame with 2 rows to df_list
    
    df_list.append(df_empty)

# Concatenate all DataFrames in df_list
df_all = pd.concat(df_list, ignore_index=True)

# Clean up the spurious characters in the Values column
df_all['Value'] = df_all['Value'].str.replace("Title:", "")
df_all['Value'] = df_all['Value'].str.replace("**", "")
df_all['Value'] = df_all['Value'].str.replace("\"", "")
df_all['Value'] = df_all['Value'].str.replace("1\.|2\.|3\.", "", regex=True)
df_all['Value'] = df_all['Value'].str.replace("Call to Action Button:", "")

# Add a new column 'Reference' which is 1 where 'Key' is 'Title' and 0 elsewhere
df_all['Reference'] = (df_all['Key'] == 'Title').astype(int)

# Replace 0s with NaNs in 'Reference' column
df_all['Reference'] = df_all['Reference'].replace(0, np.nan)

# Forward fill NaNs in 'Reference' column
df_all['Reference'] = df_all['Reference'].ffill()

# Initialize reference number
ref_num = 0

# Create a new 'Reference' column with empty strings as default values
df_all['Reference'] = ''

# Iterate over the DataFrame
for i, row in df_all.iterrows():
    # If the 'Key' is 'Title', increment the reference number
    if row['Key'] == 'Title':
        ref_num += 1
    # Set the 'Reference' value for the current row
    df_all.at[i, 'Reference'] = ref_num if row['Key'] == 'Title' else ''

# Reorder the columns
df_all = df_all[['Reference', 'Key', 'Value']]




# Write df_all to the Excel file
with pd.ExcelWriter(xlsx_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_all.to_excel(writer, sheet_name=output_sheet_name, index=False)

    # Get the openpyxl workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets[output_sheet_name]

    # Set the width of the columns
    worksheet.column_dimensions['A'].width = 10
    worksheet.column_dimensions['B'].width = 30
    worksheet.column_dimensions['C'].width = 100

    # Set word wrap for specific columns
    for row in worksheet['B:C']:
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)


