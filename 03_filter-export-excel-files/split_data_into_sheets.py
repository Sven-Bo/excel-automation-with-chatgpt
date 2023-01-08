import pandas as pd
import os

# Load the Data sheet in the Financial_Data.xlsx workbook
df = pd.read_excel('Financial_Data.xlsx', sheet_name='Data')

# Filter the data for the year 2021
df = df[df['Year'] == 2021]

# Get a list of unique countries
countries = df['Country'].unique()

# Create the Attachments folder if it doesn't already exist
if not os.path.exists('Attachments'):
    os.makedirs('Attachments')

# Iterate through each country
for country in countries:
    # Extract the financial data for this country
    country_df = df[df['Country'] == country]
    
    # Select the columns from A to P
    country_df = country_df.iloc[:, :16]
    
    # Save the data for this country to an Excel file
    country_df.to_excel(f'Attachments/{country}.xlsx', index=False)
