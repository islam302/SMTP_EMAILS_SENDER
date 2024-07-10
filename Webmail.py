# -*- coding: utf-8 -*-
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
from PIL import Image, ImageTk
import glob
import sys

class EmailSenderApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Email Sender")
        self.geometry("800x700")
        self.configure(bg="white")
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure('TButton', background='#0078D7', foreground='white', font=('Segoe UI', 10, 'bold'))
        self.style.configure('TLabel', background='white', foreground='#0078D7', font=('Segoe UI', 10))
        self.style.configure('TEntry', font=('Segoe UI', 10))
        self.style.configure('TFrame', background='white')

        self.word_template_path = None
        self.excel_file_path = None

        self.create_widgets()

    def find_logo_path(self):
        current_dir = os.path.dirname(sys.argv[0])  # For PyInstaller compatibility
        logo_files = glob.glob(os.path.join(current_dir, "logo.*"))
        if logo_files:
            return logo_files[0]
        else:
            return None

    def create_widgets(self):
        top_frame = ttk.Frame(self, padding=20, relief='raised')
        top_frame.pack(fill='both', expand=True)

        # Logo
        logo_path = self.find_logo_path()
        if logo_path and os.path.exists(logo_path):
            try:
                logo_image = Image.open(logo_path)
                logo_image = logo_image.resize((250, 250), Image.LANCZOS)
                logo_photo = ImageTk.PhotoImage(logo_image)
                logo_label = ttk.Label(top_frame, image=logo_photo, background='white')
                logo_label.image = logo_photo
                logo_label.grid(row=0, column=0, columnspan=3, pady=10)
            except Exception as e:
                print(f"Error loading image: {e}")
                messagebox.showerror("Error", "Failed to load logo image")

        label_subject = ttk.Label(top_frame, text="Subject:", font=("Segoe UI", 16, "bold"))
        label_subject.grid(row=1, column=0, pady=20, padx=20, sticky='w')

        self.entry_subject = ttk.Entry(top_frame, width=70, font=("Segoe UI", 12))
        self.entry_subject.grid(row=1, column=1, pady=20, padx=10, sticky='w')

        btn_browse_word = ttk.Button(top_frame, text="Choose Word Template", command=self.browse_word_template,
                                     style='TButton')
        btn_browse_word.grid(row=2, column=1, pady=10, padx=10, sticky='w')

        btn_browse_excel = ttk.Button(top_frame, text="Choose Excel File", command=self.browse_excel_file,
                                      style='TButton')
        btn_browse_excel.grid(row=3, column=1, pady=10, padx=10, sticky='w')

        btn_send_emails = ttk.Button(top_frame, text="Send Emails", command=self.send_emails, style='TButton')
        btn_send_emails.grid(row=4, column=1, pady=20, padx=10, sticky='w')

        btn_exit = ttk.Button(top_frame, text="Exit", command=self.destroy, style='TButton')
        btn_exit.grid(row=5, column=1, pady=20, padx=10, sticky='w')

    def browse_word_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if filename:
            self.word_template_path = filename
            messagebox.showinfo("Word Template Selected", f"Selected template: {filename}")

    def browse_excel_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.excel_file_path = filename
            messagebox.showinfo("Excel File Selected", f"Selected file: {filename}")

    def send_emails(self):
        try:
            subject = self.entry_subject.get()
            word_path = self.word_template_path
            excel_path = self.excel_file_path

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
                full_text = []
                for para in doc.paragraphs:
                    full_text.append(para.text)
                return '\n'.join(full_text)

            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
            server.login(smtp_username, smtp_password)

            # Function to replace variables in the template with values from the Excel row
            def replace_variables(template, row):
                pattern = re.compile(r'«(.*?)»')
                matches = pattern.findall(template)
                for match in matches:
                    if match in row:
                        value = row[match]
                        if pd.isna(value):
                            value = ""
                        template = template.replace(f'«{match}»', str(value))
                    else:
                        print(f"Warning: Column '{match}' not found in the Excel file")
                return template

            # Sending emails
            for index, row in df.iterrows():
                try:
                    email = row.get('email')
                    if not email:
                        print(f"Skipping row {index + 1}: No email specified")
                        continue

                    attachments = row.get('attachments')
                    if pd.isna(attachments):
                        attachments = ''

                    # Customize message body using variables from Excel
                    message_template = read_word_template(word_path)
                    message_body = replace_variables(message_template, row)

                    # Create email message
                    msg = MIMEMultipart()
                    msg['From'] = smtp_username
                    msg['To'] = email
                    msg['Subject'] = subject if subject else "No Subject"  # Default subject if not provided

                    # Attach Arabic text with proper encoding
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
                                    part.add_header('Content-Disposition', f'attachment; filename= {file}')
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

            # Closing SMTP connection
            server.quit()

            messagebox.showinfo("Emails Sent", "All emails sent successfully!")

        except FileNotFoundError:
            messagebox.showerror("File Error", "File not found.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    app = EmailSenderApp()
    app.mainloop()
