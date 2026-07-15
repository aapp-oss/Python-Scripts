import pandas as pd

# Load the excel file
input_file = "2026 Fact Book Data.xlsx"
df = pd.read_excel(input_file, engine='openpyxl')

# Convert the columns to lowercase
# Make sure the colums name match 
df["Preferred City"] = df["Preferred City"].str.lower().str.title()

# Save the updaeted file
output_file = "2026 Fact Book Data Updated.xlsx"
df.to_excel(output_file, index=False)

print(f"Updated file saved as '{output_file}' with 'Preferred City' column converted to lowercase.")