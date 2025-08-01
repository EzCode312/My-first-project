import pandas as pd
import win32com.client as win32

# Load Excel and clean column names
df = pd.read_excel(r"C:\Users\kudjoel\Downloads\CRM_AR_List_UK_Large.xlsx")
df.columns = df.columns.str.strip()

# Filter one overdue customer
overdue_customers = df[df["Status"] == "Overdue"].head(1)

# Outlook instance
outlook = win32.Dispatch('outlook.application')

# Get the first row
row = overdue_customers.iloc[0]

mail = outlook.CreateItem(0)
mail.To = row["Email"]
mail.Subject = f"Reminder: Overdue Invoice {row['Invoice #']}"

mail.HTMLBody = f"""
<p>Dear {row['Customer Name']},</p>
<p>This is a reminder that your invoice <strong>{row['Invoice #']}</strong> is overdue.</p>
<p><strong>Amount Due:</strong> £{row['Balance Due']:.2f}<br>
<strong>Due Date:</strong> {row['Due Date'].strftime('%Y-%m-%d')}</p>
<p>Please arrange payment at your earliest convenience to avoid disruption.</p>
<p>Best regards,<br><em>Accounts Receivable Team</em></p>
"""

# For testing, preview email instead of sending
mail.Display()
