import tkinter as tk
from tkinter import messagebox
import pandas as pd
import win32com.client as win32

def send_emails():
    try:
        df = pd.read_excel(r"C:\Users\kudjoel\Downloads\CRM_AR_List_UK_Large.xlsx")
        overdue_customers = df[df["Status"] == "Overdue"]
        
        outlook = win32.Dispatch('outlook.application')
        for index, row in overdue_customers.iterrows():
            mail = outlook.CreateItem(0)
            mail.To = row['Email']
            mail.Subject = f"Payment Reminder: Invoice {row['Invoice #']}"
            mail.Body = f"""Dear {row['Customer Name']},

This is a friendly reminder that invoice {row['Invoice #']} with a balance of £{row['Balance Due']} was due on {row['Due Date'].strftime('%d %b %Y')}. Please arrange payment at your earliest convenience.

Thank you,
Your Company Name
"""
            mail.Send()  # Sends the email directly
        messagebox.showinfo("Success", "Emails sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# GUI setup
root = tk.Tk()
root.title("Send Overdue Emails")
root.geometry("300x100")

btn = tk.Button(root, text="Send Overdue Emails", command=send_emails, padx=10, pady=10)
btn.pack(expand=True)

root.mainloop()
