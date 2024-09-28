
import pandas as pd
import re

FILEPATH = r'C:\Users\test\Downloads\All Project Selections Final.xlsx'
OUTPUT_FILE = 'output/Project_Selections_Cleaned.xlsx'
PROJECT_COLUMNS_START = 4
MAX_SHEET_NAME_LENGTH = 30

dataframes = pd.read_excel(FILEPATH, sheet_name=None)
project_selections_df = dataframes['All_Project_Selections']
project_columns = project_selections_df.columns[PROJECT_COLUMNS_START:]

project_data = {col: [] for col in project_columns}

for _, row in project_selections_df.iterrows():
    for col in project_columns:
        if pd.notna(row[col]):
            project_data[col].append({
                'Academic Year': row['Academic Year'],
                'Email': row['Email'],
                'Student ID Number': row['Student ID Number'],
                'Name': row['Name'],
                'Choice': row[col]
            })

def sanitize_sheet_name(name):
    return re.sub(r'[\[\]\*\?\/\\:]', '', name)[:MAX_SHEET_NAME_LENGTH]

with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
    for project, data in project_data.items():
        project_df = pd.DataFrame(data)
        project_df.sort_values(by='Choice', inplace=True)
        sanitized_name = sanitize_sheet_name(project)
        project_df.to_excel(writer, sheet_name=sanitized_name, index=False)

print(f"Data has been successfully written to '{OUTPUT_FILE}'")
