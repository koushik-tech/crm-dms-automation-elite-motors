import pandas as pd

# Read Excel file
df = pd.read_excel("Data Bank June 02062025.xlsx")  # Reads entire sheet by default

# Select only two columns (replace with your actual column names)
df = df[['Chassis No', 'Registration No']]

# Loop over the DataFrame and print one column's value
for index, row in df.iterrows():
    chasis_no = row["Chassis No"]
    print(chasis_no)
    print(type(chasis_no))
