import pandas as pd
from datetime import datetime

# File paths
cs_path = r"C:\Users\wi_ritviqs\Downloads\CSdump13june2025.csv"
ad_path = r"C:\Users\wi_ritviqs\Downloads\AD13june2025.xlsx"
output_path = r"C:\Users\wi_ritviqs\Downloads\CS15DaysProcessed.xlsx"

# Load files
cs_df = pd.read_csv(cs_path)
ad_df = pd.read_excel(ad_path)

# Normalize column names
cs_df.columns = cs_df.columns.str.strip()
ad_df.columns = ad_df.columns.str.strip()

# Step 1: Filter CS to keep only Hostname and Last Seen
cs_df = cs_df[['Hostname', 'Last Seen']].copy()

# Step 2: Rename for processing
cs_df.rename(columns={'Hostname': 'Hostname_temp'}, inplace=True)
cs_df.insert(0, 'Hostname', cs_df['Hostname_temp'])
cs_df.drop(columns=['Hostname_temp'], inplace=True)

# Step 3: Prepare AD data (assuming 'cn' column in AD holds hostnames)
lookup_fields = ['DN', 'userAccountControl', 'pwdLastSet', 'operatingSystem']
ad_lookup = ad_df[['cn'] + lookup_fields].copy()  # Use 'cn' as Hostname
ad_lookup.rename(columns={'cn': 'Hostname'}, inplace=True)

# Step 4: Merge CS with AD on Hostname
merged_df = pd.merge(cs_df, ad_lookup, on='Hostname', how='left')

# Step 5: Reorder and rename columns (A to G)
merged_df = merged_df[[
    'Hostname',                     # A
    'DN',           # B (DN)
    'userAccountControl',          # C
    'pwdLastSet',                  # D
    'operatingSystem',             # E
    'Last Seen'                    # F (will move to G)
]]

# Step 6: Rename and move 'Last Seen' to G
merged_df.rename(columns={'Last Seen': 'Last_seen'}, inplace=True)
merged_df.insert(5, 'Filler', '')  # Insert empty F column
merged_df.insert(6, 'Last_seen', merged_df.pop('Last_seen'))  # Now 'Last_seen' is at G

# Step 7: Convert 'Last_seen' to datetime (remove timezone)
merged_df['Last_seen'] = pd.to_datetime(merged_df['Last_seen'], errors='coerce').dt.tz_localize(None)

# Step 8: Remove duplicates based on Hostname, keep latest
merged_df.sort_values(by='Last_seen', ascending=False, inplace=True)
merged_df.drop_duplicates(subset='Hostname', keep='first', inplace=True)

# Step 9: Calculate Ageing in column H
today = pd.Timestamp.today().normalize()
merged_df.insert(7, 'Ageing', (today - merged_df['Last_seen']).dt.days)

# Step 10: Filter Ageing between 0 and 15 days
filtered_df = merged_df[(merged_df['Ageing'] >= 0) & (merged_df['Ageing'] <= 15)]

# Step 11: Export result
filtered_df.to_excel(output_path, index=False)
print(f"✅ Success! Final file saved at: {output_path}")
