"""
This module processes an Excel file and writes the output to a new sheet.
"""
import os
import pandas as pd
# pylint: disable=invalid-name

input_sheet_name = 'Input'
# rest of your code

# path to your Excel file.
xlsx_path = os.path.abspath('./2024-03-25 Explainers.xlsx')
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

# Convert the list of dictionaries to a DataFrame for the output.
output_df = pd.DataFrame(output_data_list)

# Write the output DataFrame to a new sheet in the Excel workbook.
output_sheet_name = 'Output'

with pd.ExcelWriter(xlsx_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    output_df.to_excel(writer, sheet_name='Sheet1')

print("Data has been processed and written to the output sheet.")
