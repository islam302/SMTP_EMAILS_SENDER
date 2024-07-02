import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def read_contacts(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df['name'], df['email']

def send_email(sender_email, sender_password, recipient_email, subject, message):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        msg.attach(MIMEText(message, 'plain'))

        server = smtplib.SMTP('smtp.gmail.com', 587)  # Use your SMTP server details
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        print(f"Email sent successfully to {recipient_email}")
    except Exception as e:
        print(f"Error sending email to {recipient_email}: {str(e)}")

def main():
    # Excel file path and selected sheet name
    file_path = 'contacts.xlsx'  # Update with your Excel file path
    sheet_name = 'Sheet1'  # Update with your sheet name

    # Email details
    sender_email = 'youremail@gmail.com'  # Update with your email address
    sender_password = "smtp password"  # Update with your email password
    subject = 'test'
    message = 'hi stubid'

    # Read contacts from Excel
    names, emails = read_contacts(file_path, sheet_name)

    # Send emails
    for name, email in zip(names, emails):
        recipient_email = email
        personalized_message = f"Dear {name},\n\n{message}"

        send_email(sender_email, sender_password, recipient_email, subject, personalized_message)

if __name__ == "__main__":
    main()

