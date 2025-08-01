import pandas as pd

# Load Excel file
df = pd.read_excel("CRM_AR_List_UK_Large.xlsx")

# Strip any extra whitespace from column names
df.columns = df.columns.str.strip()

# Check column names (optional)
print(df.columns)

# Filter one test row
overdue_customers = df[df["Status"] == "Overdue"].head(1)
print(overdue_customers)
