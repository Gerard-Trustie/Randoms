import pandas as pd

# path to your Excel file.
xlsx_path = './2024-03-25 Explainers.xlsx'
workbook = pd.ExcelFile(xlsx_path)

# Read the input sheet.
input_sheet_name = 'Input'
input_df = workbook.parse(input_sheet_name, header=None)

# Initialize a counter for conversation starters to handle multiple occurrences.
cs_counter = 1

# Initialize the dictionary that will be used to create the DataFrame for the output sheet.
output_data = {
    'Title': '',
    'Trustie Bot says': '',
    'Conversation Starter 1': '',
    'Conversation Starter 2': '',
    'Conversation Starter 3': '',
    'Call to Action': ''
}

# Iterate over the rows of the input DataFrame and extract data based on labels in column B.
for _, row in input_df.iterrows():
    label = row[1]  # Label in column B
    text = row[2]  # Corresponding text in column C

    if pd.notna(label):
        if label == 'Title':
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

# Convert the dictionary to a DataFrame for the output.
output_df = pd.DataFrame([output_data])

# Write the output DataFrame to a new sheet in the Excel workbook.
# Replace 'Output' with the actual name of your output sheet.
output_sheet_name = 'Output'
with pd.ExcelWriter(xlsx_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    output_df.to_excel(writer, sheet_name=output_sheet_name, index=False)

print("Data has been processed and written to the output sheet.")
