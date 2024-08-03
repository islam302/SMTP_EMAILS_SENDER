# -*- coding: utf-8 -*-
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def read_contacts(file_path):
    df = pd.read_excel(file_path)
    return df['name'], df['email']

def send_email(smtp_username, smtp_password, from_email, to_email, subject, body):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
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

def main():
    # Excel file path and selected sheet name
    file_path = 'data.xlsx'  # Update with your Excel file path
    # Email details
    sender_email = 'islambadran39@gmail.com'  # Update with your email address
    sender_password = 'yrwp xhds hdtp cuqe'  # Update with your email password
    from_email = 'tcc@una-oic.org'
    subject = 'GMAIL'
    message = 'hi stubid'

    # Read contacts from Excel
    names, emails = read_contacts(file_path)

    # Send emails
    for name, email in zip(names, emails):
        recipient_email = email

        send_email(sender_email, sender_password, from_email, recipient_email, subject, message)

if __name__ == "__main__":
    main()
