"""
This module copies and transposes data from one excel sheet to another.
It copies answers from the CS_Answers sheet to the Output sheet based on the 
Explainer Reference number and Conversation Starter number.
"""
import os
import pandas as pd
import numpy as np  
from openpyxl.styles import Alignment
# pylint: disable=invalid-name

# Load the Excel file.  path to your Excel file.
xlsx_path = os.path.abspath('./2024-03-25b Explainers.xlsx')
workbook = pd.ExcelFile(xlsx_path)
output_sheet_name = 'Output'
cs_answers_sheet_name = 'CS_Answers'
output_with_answers_sheet_name = 'Output with Answers'

# Read the sheets into DataFrames
output_df = pd.read_excel(xlsx_path, sheet_name=output_sheet_name)
cs_answers_df = pd.read_excel(xlsx_path, sheet_name=cs_answers_sheet_name)

# Initialize the maximum reference number seen so far
max_reference_number = 0

# Dictionary to store CS answers keyed by Reference number and Conversation Starter
cs_answers_dict = {}

# Iterate through the CS_Answers sheet to collect answers
for _, row in cs_answers_df.iterrows():
    # Update the maximum reference number seen so far
    reference_number = max(max_reference_number, row['Reference'])
    max_reference_number = reference_number
  
    # Check if the row contains a Conversation Starter heading
    if pd.notnull(row['Answers']) and "Conversation Starter" in row['Answers']:
        # Identify the number of the Conversation Starter (1, 2, or 3)
        # Conversation Starter headings are assumed to be in the format "Conversation Starter X: ..."
        # Extract the Conversation starter number from the string by selecting the number berfore the colon
        cs_number = row['Answers'].split(":")[0].split()[-1]  # The last word contains the number
        
        # Collect the next five rows as answers
        answers = cs_answers_df.iloc[_+2 : _+7, 2].tolist()  # Assuming answers are in Column C
        
        # Store these answers in the dictionary with a tuple key of (reference_number, cs_number)
        cs_answers_dict[(reference_number, cs_number)] = answers

# print("Answers have been collected from the CS_Answers sheet.", cs_answers_dict)

# Insert answers into the Output DataFrame
max_reference_number = 0 # Reset the maximum reference number
for index, row in output_df.iterrows():

    reference_number = max(max_reference_number, row['Reference'])
    max_reference_number = reference_number
    # Identify the Conversation Starter by checking if it contains the string "Conversation Starter"
    if pd.notnull(row['Key']) and  "Conversation Starter" in row['Key']:
        # Extract the number of the Conversation Starter from the string
        cs_number = row['Key'].split()[-1]  

        # Use the  reference number to find the correct answers
        answers = cs_answers_dict.get((reference_number, cs_number), [""] * 5)
        # print('E C A',reference_number, cs_number, answers)

        
# Create a new DataFrame to hold the modified data
new_output_df = pd.DataFrame()
reference_number = 0
max_reference_number = 0

for index, row in output_df.iterrows():
    reference_number = max(max_reference_number, row['Reference'])
    max_reference_number = reference_number

    # check if the row contains the word 'Title' in the 'Key' column and if so replace it with the word 'Explainer'
    if pd.notnull(row['Key']) and "Title" in row['Key']:
        row['Key'] = row['Key'].replace('Title', 'Explainer')   

    # Add the current row to the new DataFrame
    new_output_df = pd.concat([new_output_df, pd.DataFrame(row).transpose()])


    if pd.notnull(row['Key']) and  "Conversation Starter" in row['Key']:
        cs_number = row['Key'].split()[-1]  
        answers = cs_answers_dict.get((reference_number, cs_number), [""] * 5)
        # print('answers dict',reference_number, cs_number, answers)

        # Create a new DataFrame for each answer and append it to the new DataFrame
        for answer in answers:
            answer_data = pd.Series([None]*len(new_output_df.columns), index=new_output_df.columns)
            answer_data['Answer'] = answer
            answer_df = pd.DataFrame([answer_data])
            # print('answer data frame',reference_number, cs_number, answer_df)
            new_output_df = pd.concat([new_output_df, answer_df])


# Reset the index of the new DataFrame
new_output_df.reset_index(drop=True, inplace=True)



# Save the modified DataFrame back to the Excel file
with pd.ExcelWriter(xlsx_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    new_output_df.to_excel(writer, sheet_name=output_with_answers_sheet_name, index=False)

    # Get the openpyxl workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets[output_with_answers_sheet_name]

    # Set the width of the columns
    worksheet.column_dimensions['A'].width = 10
    worksheet.column_dimensions['B'].width = 30
    worksheet.column_dimensions['C'].width = 100
    worksheet.column_dimensions['D'].width = 75

    # Set formatting for specific columns
    for row in worksheet['A:D']:
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='center')
        if row[1].value == 'Trustie Bot says':
            worksheet.row_dimensions[row[0].row].height = 100
        elif row[3].value is not None:
            worksheet.row_dimensions[row[0].row].height = 35
        else:
            worksheet.row_dimensions[row[0].row].height = 18

           

print("Answers have been populated for each Conversation Starter in the Output sheet.")







