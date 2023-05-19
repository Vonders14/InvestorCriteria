import pandas as pd
import re

# Step 1: Read the Excel file
file_path = r'C:\Users\dhoffmann\Downloads\Apollo RE Investor Criteria_MASTER - COPY_bad text removed (1).xlsx'
df = pd.read_excel(file_path)

# Step 2: Search for state names in all columns
state_names = ['Alabama', 'Alaska', 'Arizona', 'Arkansas', 'California', 'Colorado', 'Connecticut', 'Delaware',
               'Florida', 'Georgia', 'Hawaii', 'Idaho', 'Illinois', 'Indiana', 'Iowa', 'Kansas', 'Kentucky',
               'Louisiana', 'Maine', 'Maryland', 'Massachusetts', 'Michigan', 'Minnesota', 'Mississippi',
               'Missouri', 'Montana', 'Nebraska', 'Nevada', 'New Hampshire', 'New Jersey', 'New Mexico', 'New York',
               'North Carolina', 'North Dakota', 'Ohio', 'Oklahoma', 'Oregon', 'Pennsylvania', 'Rhode Island',
               'South Carolina', 'South Dakota', 'Tennessee', 'Texas', 'Utah', 'Vermont', 'Virginia', 'Washington',
               'West Virginia', 'Wisconsin', 'Wyoming']

state_names_pattern = r'(' + '|'.join(state_names) + r')'

# Define a function to find states in a string
def find_states(s):
    matches = re.findall(state_names_pattern, s, re.IGNORECASE)
    matches = [match for match in matches]
    return ', '.join(matches)

# Apply the function to each row in the DataFrame
df['State Names'] = df.apply(lambda row: find_states(str(row.values)), axis=1)

# Step 3: Save the modified DataFrame to a new Excel file
new_file_path = r'C:\Users\dhoffmann\Documents\Apollo RE Investor_Criteria with States.xlsx'
df.to_excel(new_file_path, index=False)

