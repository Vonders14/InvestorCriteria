import pandas as pd
import re

# Step 1: Read the Excel file
file_path = r'C:\Users\dhoffmann\Documents\MULTIFAMILY_RE Criteria.xlsx'
df = pd.read_excel(file_path)

# Step 2: Search for state abbreviations in all columns
abbreviations = ['AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY',
                 'LA', 'ME', 'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 'NM', 'NY', 'NC', 'ND',
                 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY']

abbreviations_pattern = r"(?<!\S)(?:{})\b".format("|".join(abbreviations))

df['Abbreviations'] = df.apply(lambda row: ', '.join(set(re.findall(abbreviations_pattern, str(row.values), re.IGNORECASE))), axis=1)

# Step 3: Search for state names in all columns
state_names = ['Alabama', 'Alaska', 'Arizona', 'Arkansas', 'California', 'Colorado', 'Connecticut', 'Delaware',
               'Florida', 'Georgia', 'Hawaii', 'Idaho', 'Illinois', 'Indiana', 'Iowa', 'Kansas', 'Kentucky',
               'Louisiana', 'Maine', 'Maryland', 'Massachusetts', 'Michigan', 'Minnesota', 'Mississippi',
               'Missouri', 'Montana', 'Nebraska', 'Nevada', 'New Hampshire', 'New Jersey', 'New Mexico', 'New York',
               'North Carolina', 'North Dakota', 'Ohio', 'Oklahoma', 'Oregon', 'Pennsylvania', 'Rhode Island',
               'South Carolina', 'South Dakota', 'Tennessee', 'Texas', 'Utah', 'Vermont', 'Virginia', 'Washington',
               'West Virginia', 'Wisconsin', 'Wyoming']

state_names_pattern = r"\b(?:{})\b".format("|".join(state_names), re.IGNORECASE)

df['State Names'] = df.apply(lambda row: ', '.join(set(re.findall(state_names_pattern, str(row.values), re.IGNORECASE))), axis=1)

# Step 4: Save and close the Excel file
df.to_excel(file_path, index=False)

