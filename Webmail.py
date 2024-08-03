import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from docx import Document
import re

class EmailSenderApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Email Sender App")
        self.geometry("500x300")

        self.create_widgets()

    def create_widgets(self):
        # Subject Entry
        self.subject_label = ttk.Label(self, text="Subject:")
        self.subject_label.pack(pady=5)
        self.subject_entry = ttk.Entry(self, width=50)
        self.subject_entry.pack(pady=5)

        # Word Template Path
        self.word_path_button = ttk.Button(self, text="Select Word Template", command=self.select_word_template)
        self.word_path_button.pack(pady=5)
        self.word_path_label = ttk.Label(self, text="No file selected")
        self.word_path_label.pack(pady=5)

        # Excel File Path
        self.excel_path_button = ttk.Button(self, text="Select Excel File", command=self.select_excel_file)
        self.excel_path_button.pack(pady=5)
        self.excel_path_label = ttk.Label(self, text="No file selected")
        self.excel_path_label.pack(pady=5)

        # Send Emails Button
        self.send_button = ttk.Button(self, text="Send Emails", command=self.send_emails)
        self.send_button.pack(pady=20)

    def select_word_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if file_path:
            self.word_path_label.config(text=file_path)
            self.word_path = file_path

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.excel_path_label.config(text=file_path)
            self.excel_path = file_path

    def send_emails(self):
        try:
            subject = self.subject_entry.get()
            word_path = getattr(self, 'word_path', None)
            excel_path = getattr(self, 'excel_path', None)

            # Validate input
            if not subject:
                messagebox.showerror("Error", "Please enter a subject")
                return

            if not word_path or not os.path.exists(word_path):
                messagebox.showerror("Error", "Please choose a Word template file")
                return

            if not excel_path or not os.path.exists(excel_path):
                messagebox.showerror("Error", "Please choose an Excel file")
                return

            # Read Excel file
            df = pd.read_excel(excel_path)

            # SMTP settings
            smtp_server = 'mail.una-oic.org'
            smtp_port = 465  # Use port 465 for SSL/TLS
            smtp_username = 'messages@una-oic.org'
            smtp_password = '}E~8NLAZ5Ki3'

            # Read message template from Word file
            def read_word_template(word_path):
                doc = Document(word_path)
                full_text = [para.text for para in doc.paragraphs]
                return '\n'.join(full_text)

            # Function to replace variables in the template with values from the Excel row
            def replace_variables(template, row):
                pattern = re.compile(r'«(.*?)»')
                matches = pattern.findall(template)
                for match in matches:
                    if match in row:
                        value = row[match]
                        template = template.replace(f'«{match}»', str(value) if not pd.isna(value) else "")
                    else:
                        print(f"Warning: Column '{match}' not found in the Excel file")
                return template

            # Send emails
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(smtp_username, smtp_password)
                for index, row in df.iterrows():
                    try:
                        email = row.get('email')
                        if not email:
                            print(f"Skipping row {index + 1}: No email specified")
                            continue

                        attachments = row.get('attachments', '')
                        message_template = read_word_template(word_path)
                        message_body = replace_variables(message_template, row)

                        # Create email message
                        msg = MIMEMultipart()
                        msg['From'] = smtp_username
                        msg['To'] = email
                        msg['Subject'] = subject

                        msg.attach(MIMEText(message_body, 'plain', 'utf-8'))

                        # Attach selected files
                        if attachments:
                            attachment_files = attachments.split(',')
                            for file in attachment_files:
                                file_path = os.path.join('attachments', file.strip())
                                if os.path.exists(file_path):
                                    with open(file_path, 'rb') as f:
                                        part = MIMEBase('application', 'octet-stream')
                                        part.set_payload(f.read())
                                        encoders.encode_base64(part)
                                        part.add_header('Content-Disposition', f'attachment; filename={file}')
                                        msg.attach(part)
                                else:
                                    print(f"Attachment file '{file}' not found.")

                        # Send the email
                        server.send_message(msg)
                        print(f'Email sent to {email}')

                    except KeyError as e:
                        print(f"Error: Missing column in the Excel file - {e}")
                    except Exception as e:
                        print(f"An error occurred: {e}")

            messagebox.showinfo("Emails Sent", "All emails sent successfully!")

        except FileNotFoundError:
            messagebox.showerror("File Error", "File not found.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    app = EmailSenderApp()
    app.mainloop()
