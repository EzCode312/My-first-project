import pandas as pd
import win32com.client as win32

# Load data
df = pd.read_excel(r"C:\Users\kudjoel\Downloads\CRM_AR_List_UK_Large.xlsx")
df.columns = df.columns.str.strip()

overdue_customers = df[df["Status"] == "Overdue"]

outlook = win32.Dispatch('outlook.application')

# Your email for testing
test_email = "lemuel.kudjoe@cannock.nl"

# Send test email only for the first overdue customer but to YOUR email
row = overdue_customers.iloc[0]
mail = outlook.CreateItem(0)
mail.To = test_email  # override recipient with your email
mail.Subject = f"TEST Reminder: Overdue Invoice {row['Invoice #']}"
mail.HTMLBody = f"""
<p>Dear {row['Customer Name']},</p>
<p>This is a TEST reminder that your invoice <strong>{row['Invoice #']}</strong> is overdue.</p>
<p><strong>Amount Due:</strong> £{row['Balance Due']:.2f}<br>
<strong>Due Date:</strong> {row['Due Date'].strftime('%Y-%m-%d')}</p>
<p>Please arrange payment at your earliest convenience.</p>
<p>Best regards,<br><em>Accounts Receivable Team</em></p>
"""
mail.Send()
print("Test email sent to yourself.")
