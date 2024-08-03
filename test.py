import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os


def send_email(smtp_server, smtp_port, smtp_username, smtp_password, from_email, to_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            print(f"Email sent from {from_email} to {to_email}")
    except Exception as e:
        print(f"Failed to send email: {e}")


def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        file_path_var.set(filename)


def start_sending():
    file_path = file_path_var.get()
    if not os.path.isfile(file_path):
        messagebox.showerror("Error", "File not found!")
        return

    df = pd.read_excel(file_path, engine='openpyxl')
    smtp_username = smtp_username_var.get()
    smtp_password = smtp_password_var.get()
    from_email = from_email_var.get()
    subject = subject_var.get()
    body = body_var.get()

    if 'email' in df.columns:
        for email in df['email']:
            if pd.notna(email):
                send_email(smtp_server, smtp_port, smtp_username, smtp_password, from_email, email, subject, body)
        messagebox.showinfo("Success", "Emails sent successfully!")
    else:
        messagebox.showerror("Error", "The 'email' column does not exist in the Excel file.")


smtp_server = 'smtp.gmail.com'
smtp_port = 587

# GUI setup
root = tk.Tk()
root.title("Email Sender")

file_path_var = tk.StringVar()
smtp_username_var = tk.StringVar()
smtp_password_var = tk.StringVar()
from_email_var = tk.StringVar()
subject_var = tk.StringVar()
body_var = tk.StringVar()

tk.Label(root, text="SMTP Username:").pack()
tk.Entry(root, textvariable=smtp_username_var).pack()

tk.Label(root, text="SMTP Password:").pack()
tk.Entry(root, textvariable=smtp_password_var, show='*').pack()

tk.Label(root, text="From Email:").pack()
tk.Entry(root, textvariable=from_email_var).pack()

tk.Label(root, text="Subject:").pack()
tk.Entry(root, textvariable=subject_var).pack()

tk.Label(root, text="Body:").pack()
tk.Entry(root, textvariable=body_var).pack()

tk.Label(root, text="Excel File:").pack()
tk.Entry(root, textvariable=file_path_var).pack()
tk.Button(root, text="Browse", command=browse_file).pack()

tk.Button(root, text="Send Emails", command=start_sending).pack()

root.mainloop()
