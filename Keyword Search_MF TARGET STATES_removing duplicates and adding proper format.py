import pandas as pd

# Read the Excel file
df = pd.read_excel(r'C:\Users\dhoffmann\Documents\MULTIFAMILY_RE Criteria.xlsx', sheet_name='Multifamily_Master')

# Create a dictionary mapping state abbreviations to state names
state_dict = {
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

## Iterate over each row in the DataFrame
for index, row in df.iterrows():
    cell_value_aw = str(row['Target States'])  # Replace 'Target States' with the actual column name to search for state names and abbreviations
    output_set = set()  # Create a set to store the unique outputs for each row
    for abbreviation, state_name in state_dict.items():
        if abbreviation in cell_value_aw or state_name in cell_value_aw:
            output_set.add(f"{abbreviation} - {state_name}")

    # Convert the output set to a string
    output_string = ", ".join(output_set)

    # Assign the output string to the 'State Names' column
    df.at[index, 'State Names'] = output_string  # Replace 'State Names' with the actual column name to store the output

# Save the updated DataFrame to the Excel file
df.to_excel(r'C:\Users\dhoffmann\Documents\MULTIFAMILY_RE Criteria.xlsx', sheet_name='Multifamily_Master', index=False)