import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)

mail.to = 'l.kudjoe@amsterdam.nl'
mail.subject = 'This is Lems test mail from Python'
mail.Body = 'Hello,\n\nThis is a plain test email using Outlook and Python.\n\nBest Regards,\nPython Script'

mail.Send()

print("Email send sucessfully!")
