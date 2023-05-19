import pandas as pd

# Step 1: Read the Excel file
file_path = r'C:\Users\dhoffmann\Documents\MULTIFAMILY_RE Criteria.xlsx'
df = pd.read_excel(file_path)

# Step 2: Search for US states and their abbreviations
state_mapping = {
    "AL": "Alabama",
    "AK": "Alaska",
    "AZ": "Arizona",
    "AR": "Arkansas",
    "CA": "California",
    "CO": "Colorado",
    "CT": "Connecticut",
    "DE": "Delaware",
    "FL": "Florida",
    "GA": "Georgia",
    "HI": "Hawaii",
    "ID": "Idaho",
    "IL": "Illinois",
    "IN": "Indiana",
    "IA": "Iowa",
    "KS": "Kansas",
    "KY": "Kentucky",
    "LA": "Louisiana",
    "ME": "Maine",
    "MD": "Maryland",
    "MA": "Massachusetts",
    "MI": "Michigan",
    "MN": "Minnesota",
    "MS": "Mississippi",
    "MO": "Missouri",
    "MT": "Montana",
    "NE": "Nebraska",
    "NV": "Nevada",
    "NH": "New Hampshire",
    "NJ": "New Jersey",
    "NM": "New Mexico",
    "NY": "New York",
    "NC": "North Carolina",
    "ND": "North Dakota",
    "OH": "Ohio",
    "OK": "Oklahoma",
    "OR": "Oregon",
    "PA": "Pennsylvania",
    "RI": "Rhode Island",
    "SC": "South Carolina",
    "SD": "South Dakota",
    "TN": "Tennessee",
    "TX": "Texas",
    "UT": "Utah",
    "VT": "Vermont",
    "VA": "Virginia",
    "WA": "Washington",
    "WV": "West Virginia",
    "WI": "Wisconsin",
    "WY": "Wyoming"
}


# Create a mapping for reverse lookup (state name to abbreviation)
reverse_state_mapping = {v: k for k, v in state_mapping.items()}

# Find the column positions for columns A to AT
start_column = 0  # Column A is at index 0
end_column = 45  # Column AT is at index 45

# Iterate over the columns A to AT and search for state abbreviations and names
for column in range(start_column, end_column + 1):
    df.iloc[:, column] = df.iloc[:, column].astype(str).apply(lambda x: x.strip())  # Convert to string and remove leading/trailing spaces
    df.iloc[:, column] = df.iloc[:, column].apply(lambda x: state_mapping.get(x.upper(), x))  # Look for exact abbreviation match
    df.iloc[:, column] = df.iloc[:, column].apply(lambda x: state_mapping.get(x.title(), x))  # Look for exact state name match

# Step 3: Continue searching for other matches even if one state is found for a record
# No additional code needed for this step since the previous code handles it

# Step 4: Add the matches to column AX
df['AX'] = df.iloc[:, start_column:end_column + 1].apply(lambda row: ', '.join(row), axis=1)  # Combine all columns A to AT

# Step 5: Remove duplicates
df['AX'] = df['AX'].apply(lambda x: ', '.join(set(x.split(', '))))

# Step 6: Format the outputs in column AX as "abbreviation - state name"
df['AX'] = df['AX'].apply(lambda x: ', '.join([f"{reverse_state_mapping.get(state.strip(), state.strip())} - {state.strip()}" for state in x.split(', ')]))

# Save the modified DataFrame back to the Excel file
df.to_excel(file_path, index=False)