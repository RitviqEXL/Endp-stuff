import pandas as pd
from datetime import datetime

# Load the Excel file
file_path = r"C:\Users\wi_ritviqs\Downloads\AD13june2025.xlsx"
df = pd.read_excel
(file_path)

# Remove unwanted columns (if present)
columns_to_remove = ['objectClass', 'whenCreated', 'whenChanged', 'operatingSystemVersion']
df.drop(columns=[col for col in columns_to_remove if col in df.columns], inplace=True)

# Rename the column if it exists
if 'lastLogonTimestamp.1' in df.columns:
    df.rename(columns={'lastLogonTimestamp.1': 'lastLogonTimestamp'}, inplace=True)

# Move column F (index 5) to G (index 6), clear F
if df.shape[1] >= 7:
    df.iloc[:, 6] = df.iloc[:, 5].astype(str)  # Copy values from F to G as string
    df.iloc[:, 5] = ''  # Clear F

# Convert column G (index 6) to datetime
df.iloc[:, 6] = pd.to_datetime(df.iloc[:, 6], errors='coerce')

# Calculate ageing from today
today = pd.to_datetime(datetime.now().date())
df['Ageing_Days'] = (today - df.iloc[:, 6]).apply(lambda x: x.days if pd.notnull(x) else None)

# Filter for ageing between 0 and 15 days
filtered_df = df[(df['Ageing_Days'] >= 0) & (df['Ageing_Days'] <= 15)]

# Export to Excel
output_path = r"C:\Users\wi_ritviqs\Downloads\filtered_modified_file.xlsx"
filtered_df.to_excel(output_path, index=False)

print("✅ Ageing calculated and filtered Excel exported successfully.")
