# -*- coding: utf-8 -*-
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import re
from docx import Document
import os

df = pd.read_excel('data.xlsx')
smtp_server = 'smtp-mail.outlook.com'
smtp_port = 587
smtp_username = 'islambadran39@gmail.com'  # ?????? ?????? ?????? ?????????? ????? ?? Outlook
smtp_password = 'jilifwjilhytlubb'

def read_word_template(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

message_template = read_word_template('message.docx')

server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(smtp_username, smtp_password)


def replace_variables(template, row):
    pattern = re.compile(r'�(.*?)�')
    matches = pattern.findall(template)
    for match in matches:
        if match in row:
            value = row[match]
            if pd.isna(value):
                value = ""
            template = template.replace(f'�{match}�', str(value))
        else:
            print(f"Warning: Column '{match}' not found in the Excel file")
    return template


for index, row in df.iterrows():
    try:
        email = row['email']
        attachments = row['attachments']  # Assuming 'attachments' column contains filenames separated by commas

        message_body = replace_variables(message_template, row)

        msg = MIMEMultipart()
        msg['From'] = smtp_username
        msg['To'] = email
        msg['Subject'] = 'OUT LOOK'

        msg.attach(MIMEText(message_body, 'plain', 'utf-8'))

        if attachments:
            attachment_files = attachments.split(',')
            for file in attachment_files:
                file_path = os.path.join('attachments', file.strip())
                if os.path.exists(file_path):
                    with open(file_path, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename= {file}')
                        msg.attach(part)
                else:
                    print(f"Attachment file '{file}' not found.")

        server.send_message(msg)
        print(f'Email sent to {email}')

    except KeyError as e:
        print(f"Error: Missing column in the Excel file - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

server.quit()
print('All emails sent successfully!')


