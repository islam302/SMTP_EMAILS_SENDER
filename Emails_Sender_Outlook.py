import tkinter as tk
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from docx import Document
import re
import os
import sys
import glob
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

class EmailSenderApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Email Sender")
        self.geometry("800x700")
        self.configure(bg="white")

        # Create a Canvas for the entire window background
        self.canvas_bg = Canvas(self, bg='white', width=800, height=700)
        self.canvas_bg.pack(fill='both', expand=True)

        # Load and display background image
        self.load_background_image()

        # Style configuration
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure('TButton', foreground='white', background='#0078D7', font=('Segoe UI', 10, 'bold'))
        self.style.configure('TLabel', background='white', foreground='#0078D7', font=('Segoe UI', 10))
        self.style.configure('TEntry', font=('Segoe UI', 10))

        self.word_template_path = None
        self.excel_file_path = None
        self.email_provider = StringVar()

        self.create_widgets()

    def create_widgets(self):
        # Subject label and entry
        label_subject = ttk.Label(self, text="Subject:", font=("Segoe UI", 16, "bold"))
        self.canvas_bg.create_window(20, 320, anchor='w', window=label_subject)

        self.entry_subject = ttk.Entry(self, width=70, font=("Segoe UI", 12))
        self.canvas_bg.create_window(150, 320, anchor='w', window=self.entry_subject)

        # Buttons
        btn_browse_word = ttk.Button(self, text="Choose Word Template", command=self.browse_word_template)
        self.canvas_bg.create_window(150, 360, anchor='w', window=btn_browse_word)

        btn_browse_excel = ttk.Button(self, text="Choose Excel File", command=self.browse_excel_file)
        self.canvas_bg.create_window(150, 400, anchor='w', window=btn_browse_excel)

        label_provider = ttk.Label(self, text="Choose Email Provider:", font=("Segoe UI", 16, "bold"))
        self.canvas_bg.create_window(20, 440, anchor='w', window=label_provider)

        provider_combobox = ttk.Combobox(self, textvariable=self.email_provider, state='readonly',
                                         font=("Segoe UI", 12))
        provider_combobox['values'] = ("Gmail", "Outlook", "Webmail")
        provider_combobox.current(0)
        self.canvas_bg.create_window(150, 440, anchor='w', window=provider_combobox)

        btn_send_emails = ttk.Button(self, text="Send Emails", command=self.send_emails)
        self.canvas_bg.create_window(150, 480, anchor='w', window=btn_send_emails)

        btn_exit = ttk.Button(self, text="Exit", command=self.destroy)
        self.canvas_bg.create_window(150, 520, anchor='w', window=btn_exit)

    def load_background_image(self):
        # Find and load the background image
        current_dir = os.path.dirname(sys.argv[0])  # For PyInstaller compatibility
        background_files = glob.glob(os.path.join(current_dir, "logo.*"))
        if background_files:
            background_image_path = background_files[0]
            if os.path.exists(background_image_path):
                try:
                    background_image = Image.open(background_image_path)
                    background_image = background_image.resize((800, 700), Image.LANCZOS)
                    self.background_photo = ImageTk.PhotoImage(background_image)
                    self.canvas_bg.create_image(0, 0, anchor=NW, image=self.background_photo)
                except Exception as e:
                    print(f"Error loading background image: {e}")
                    messagebox.showerror("Error", "Failed to load background image")
        else:
            print("Background image not found")

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
            provider = self.email_provider.get()

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

            # Read message template from Word file
            def read_word_template(word_path):
                doc = Document(word_path)
                full_text = []
                for para in doc.paragraphs:
                    full_text.append(para.text)
                return '\n'.join(full_text)

            # SMTP settings

            if provider == "Gmail":
                smtp_server = 'smtp.gmail.com'
                smtp_port = 465
                smtp_username = 'islambadran39@gmail.com'  # Update with your email address
                smtp_password = 'yrwp xhds hdtp cuqe'
            elif provider == "Outlook":
                smtp_server = 'smtp-mail.outlook.com'
                smtp_port = 587
                smtp_username = 'islambadran39@gmail.com'
                smtp_password = 'jilifwjilhytlubb'
            elif provider == "Webmail":
                smtp_server = 'mail.una-oic.org'
                smtp_port = 465
                smtp_username = 'messages@una-oic.org'
                smtp_password = '}E~8NLAZ5Ki3'

            if provider == "Outlook":
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
            else:
                server = smtplib.SMTP_SSL(smtp_server, smtp_port)

            server.login(smtp_username, smtp_password)

            def replace_variables(template, row):
                pattern = re.compile(r'«(.*?)»')  # تعديل النمط ليتناسب مع << >> بدلاً من ''
                matches = pattern.findall(template)
                for match in matches:
                    if match in row:
                        value = row[match]
                        if pd.isna(value):
                            value = ""
                        template = template.replace(f'«{match}»', str(value))
                    else:
                        print(f"Warning: Variable '{match}' not found in the Excel file")
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

                    # Attach HTML content
                    msg.attach(MIMEText(message_body, 'html', 'utf-8'))

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
            messagebox.showerror("File Error", "File not found, please check the file path")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    app = EmailSenderApp()
    app.mainloop()




