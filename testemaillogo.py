import pandas as pd
import win32com.client as win32

# Your email for testing
test_email = "lemuel.kudjoe@cannock.nl"

# Load Excel
df = pd.read_excel(r"C:\Users\kudjoel\Downloads\CRM_AR_List_UK_Large.xlsx")
overdue_customers = df[df["Status"] == "Overdue"]

outlook = win32.Dispatch('outlook.application')

# Pick the first overdue customer row (or any one)
row = overdue_customers.iloc[0]

mail = outlook.CreateItem(0)
mail.To = test_email  # Send to yourself
mail.Subject = f"Test Payment Reminder: Invoice {row['Invoice #']}"

signature_html = """
<br><br>
Best regards,<br>
<strong>Your Company Name</strong><br>
<img src="cid:company_logo" alt="Company Logo" style="width:120px;"><br>
Working for your success.<br>
"""

body_html = f"""
<p>Dear {row['Customer Name']},</p>
<p>This is a friendly reminder that invoice <strong>{row['Invoice #']}</strong> with a balance of £{row['Balance Due']} was due on {row['Due Date'].strftime('%d %b %Y')}.</p>
<p>Please arrange payment at your earliest convenience.</p>
{signature_html}
"""

mail.HTMLBody = body_html

logo_path = r"C:\Users\kudjoel\Downloads\company_logo.png"
attachment = mail.Attachments.Add(logo_path)
attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "company_logo")

mail.Send()

print("Test email sent to yourself.")
