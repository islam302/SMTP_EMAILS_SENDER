#
#
#
#
# # import pandas as pd
# # import smtplib
# # from email.mime.text import MIMEText
# # from email.mime.multipart import MIMEMultipart
# # from email.mime.base import MIMEBase
# # from email import encoders
# # import re
# # from docx import Document
# # import os
# #
# # df = pd.read_excel('contacts.xlsx')
# # smtp_server = 'smtp-mail.outlook.com'
# # smtp_port = 587
# # smtp_username = 'islambadran39@gmail.com'  # ?????? ?????? ?????? ?????????? ????? ?? Outlook
# # smtp_password = 'jilifwjilhytlubb'
# # def read_word_template(file_path):
# #     doc = Document(file_path)
# #     full_text = []
# #     for para in doc.paragraphs:
# #         full_text.append(para.text)
# #     return '\n'.join(full_text)
# #
# # message_template = read_word_template('messages-ar.docx')
# #
# # server = smtplib.SMTP(smtp_server, smtp_port)
# # server.starttls()
# # server.login(smtp_username, smtp_password)
# #
# #
# # def replace_variables(template, row):
# #     pattern = re.compile(r'�(.*?)�')
# #     matches = pattern.findall(template)
# #     for match in matches:
# #         if match in row:
# #             value = row[match]
# #             if pd.isna(value):
# #                 value = ""
# #             template = template.replace(f'�{match}�', str(value))
# #         else:
# #             print(f"Warning: Column '{match}' not found in the Excel file")
# #     return template
# #
# #
# # for index, row in df.iterrows():
# #     try:
# #         email = row['email']
# #         attachments = row['attachments']  # Assuming 'attachments' column contains filenames separated by commas
# #
# #         # ????? ?? ??????? ???????? ????????? ?? ??? Excel
# #         message_body = replace_variables(message_template, row)
# #
# #         # ????? ????? ?????? ??????????
# #         msg = MIMEMultipart()
# #         msg['From'] = smtp_username
# #         msg['To'] = email
# #         msg['Subject'] = 'OUT LOOK'
# #
# #         # ????? ???? ?????? ???????? ??????
# #         msg.attach(MIMEText(message_body, 'plain', 'utf-8'))
# #
# #         # ????? ??????? ???????
# #         if attachments:
# #             attachment_files = attachments.split(',')
# #             for file in attachment_files:
# #                 file_path = os.path.join('attachments', file.strip())
# #                 if os.path.exists(file_path):
# #                     with open(file_path, 'rb') as f:
# #                         part = MIMEBase('application', 'octet-stream')
# #                         part.set_payload(f.read())
# #                         encoders.encode_base64(part)
# #                         part.add_header('Content-Disposition', f'attachment; filename= {file}')
# #                         msg.attach(part)
# #                 else:
# #                     print(f"Attachment file '{file}' not found.")
# #
# #         # ????? ??????? ???????????
# #         server.send_message(msg)
# #         print(f'Email sent to {email}')
# #
# #     except KeyError as e:
# #         print(f"Error: Missing column in the Excel file - {e}")
# #     except Exception as e:
# #         print(f"An error occurred: {e}")
# #
# # # ????? ????? SMTP
# # server.quit()
# #
# # print('All emails sent successfully!')
# #
# #
# #
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
# # with html frame
# # import tkinter as tk
# # from tkinter import ttk, filedialog, messagebox
# # import pandas as pd
# # import os
# # import smtplib
# # from email.mime.multipart import MIMEMultipart
# # from email.mime.text import MIMEText
# # from email.mime.base import MIMEBase
# # from email import encoders
# # from docx import Document
# # import re
# # from PIL import Image, ImageTk
# # import glob
# #
# # class EmailSenderApp(tk.Tk):
# #
# #     def __init__(self):
# #         super().__init__()
# #         self.title("Email Sender")
# #         self.geometry("800x700")
# #         self.configure(bg="white")
# #         self.style = ttk.Style()
# #         self.style.theme_use("clam")
# #         self.style.configure('TButton', background='#0078D7', foreground='white', font=('Segoe UI', 10, 'bold'))
# #         self.style.configure('TLabel', background='white', foreground='#0078D7', font=('Segoe UI', 10))
# #         self.style.configure('TEntry', font=('Segoe UI', 10))
# #         self.style.configure('TFrame', background='white')
# #
# #         self.word_template_path = None
# #         self.excel_file_path = None
# #         self.message_format = tk.StringVar(value='plain')
# #
# #         self.create_widgets()
# #
# #     def find_logo_path(self):
# #         current_dir = os.path.dirname(os.path.abspath(__file__))
# #         logo_files = glob.glob(os.path.join(current_dir, "logo.*"))
# #         if logo_files:
# #             return logo_files[0]
# #         else:
# #             return None
# #
# #     def create_widgets(self):
# #         top_frame = ttk.Frame(self, padding=20, relief='raised')
# #         top_frame.pack(fill='both', expand=True)
# #
# #         # Logo
# #         logo_path = self.find_logo_path()
# #         if logo_path and os.path.exists(logo_path):
# #             logo_files = glob.glob(logo_path)
# #             if logo_files:
# #                 logo_file = logo_files[0]
# #                 try:
# #                     logo_image = Image.open(logo_file)
# #                     logo_image = logo_image.resize((250, 250), Image.LANCZOS)
# #                     logo_photo = ImageTk.PhotoImage(logo_image)
# #                     logo_label = ttk.Label(top_frame, image=logo_photo, background='white')
# #                     logo_label.image = logo_photo
# #                     logo_label.grid(row=0, column=0, columnspan=3, pady=10)
# #                 except Exception as e:
# #                     print(f"Error loading image: {e}")
# #                     messagebox.showerror("Error", "Failed to load logo image")
# #             else:
# #                 print("Logo image not found")
# #
# #         label_subject = ttk.Label(top_frame, text="Subject:", font=("Segoe UI", 16, "bold"))
# #         label_subject.grid(row=1, column=0, pady=20, padx=20, sticky='w')
# #
# #         self.entry_subject = ttk.Entry(top_frame, width=70, font=("Segoe UI", 12))
# #         self.entry_subject.grid(row=1, column=1, pady=20, padx=10, sticky='w')
# #
# #         btn_browse_word = ttk.Button(top_frame, text="Choose Word Template", command=self.browse_word_template,
# #                                      style='TButton')
# #         btn_browse_word.grid(row=2, column=1, pady=10, padx=10, sticky='w')
# #
# #         btn_browse_excel = ttk.Button(top_frame, text="Choose Excel File", command=self.browse_excel_file,
# #                                       style='TButton')
# #         btn_browse_excel.grid(row=3, column=1, pady=10, padx=10, sticky='w')
# #
# #         btn_send_emails = ttk.Button(top_frame, text="Send Emails", command=self.send_emails, style='TButton')
# #         btn_send_emails.grid(row=5, column=1, pady=20, padx=10, sticky='w')
# #
# #         btn_exit = ttk.Button(top_frame, text="Exit", command=self.destroy, style='TButton')
# #         btn_exit.grid(row=6, column=1, pady=20, padx=10, sticky='w')
# #
# #         # Message format option
# #         label_format = ttk.Label(top_frame, text="Message Format:", font=("Segoe UI", 16, "bold"))
# #         label_format.grid(row=4, column=0, pady=20, padx=20, sticky='w')
# #
# #         self.radio_plain = ttk.Radiobutton(top_frame, text='Plain Text', variable=self.message_format, value='plain')
# #         self.radio_plain.grid(row=4, column=1, padx=10, sticky='w')
# #
# #         self.radio_html = ttk.Radiobutton(top_frame, text='HTML', variable=self.message_format, value='html')
# #         self.radio_html.grid(row=4, column=1, padx=120, sticky='w')
# #
# #     def browse_word_template(self):
# #         filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
# #         if filename:
# #             self.word_template_path = filename
# #             messagebox.showinfo("Word Template Selected", f"Selected template: {filename}")
# #
# #     def browse_excel_file(self):
# #         filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
# #         if filename:
# #             self.excel_file_path = filename
# #             messagebox.showinfo("Excel File Selected", f"Selected file: {filename}")
# #
# #     def send_emails(self):
# #         try:
# #             subject_template = self.entry_subject.get()
# #             word_path = self.word_template_path
# #             excel_path = self.excel_file_path
# #
# #             # Validate input
# #             if not subject_template:
# #                 messagebox.showerror("Error", "Please enter a subject")
# #                 return
# #
# #             if not word_path or not os.path.exists(word_path):
# #                 messagebox.showerror("Error", "Please choose a Word template file")
# #                 return
# #
# #             if not excel_path or not os.path.exists(excel_path):
# #                 messagebox.showerror("Error", "Please choose an Excel file")
# #                 return
# #
# #             # Read Excel file
# #             df = pd.read_excel(excel_path)
# #
# #             # SMTP settings
# #             smtp_server = 'mail.una-oic.org'
# #             smtp_port = 465  # Use port 465 for SSL/TLS
# #             smtp_username = 'messages@una-oic.org'
# #             smtp_password = '}E~8NLAZ5Ki3'
# #
# #             # Read message template from Word file
# #             def read_word_template(word_path):
# #                 doc = Document(word_path)
# #                 full_text = []
# #                 for para in doc.paragraphs:
# #                     full_text.append(para.text)
# #                 return '\n'.join(full_text)
# #
# #             server = smtplib.SMTP_SSL(smtp_server, smtp_port)
# #             server.login(smtp_username, smtp_password)
# #
# #             # Function to replace variables in the template with values from the Excel row
# #             def replace_variables(template, row):
# #                 pattern = re.compile(r'«(.*?)»')
# #                 matches = pattern.findall(template)
# #                 for match in matches:
# #                     if match in row:
# #                         value = row[match]
# #                         if pd.isna(value):
# #                             value = ""
# #                         template = template.replace(f'«{match}»', str(value))
# #                     else:
# #                         print(f"Warning: Column '{match}' not found in the Excel file")
# #                 return template
# #
# #             # Sending emails
# #             for index, row in df.iterrows():
# #                 try:
# #                     email = row.get('email')
# #                     if not email:
# #                         print(f"Skipping row {index + 1}: No email specified")
# #                         continue
# #
# #                     attachments = row.get('attachments')
# #                     if pd.isna(attachments):
# #                         attachments = ''
# #
# #                     # Customize message body and subject using variables from Excel
# #                     message_template = read_word_template(word_path)
# #                     message_body = replace_variables(message_template, row)
# #                     subject = replace_variables(subject_template, row)
# #                     print(subject)
# #
# #                     # Create email message
# #                     msg = MIMEMultipart()
# #                     msg['From'] = smtp_username
# #                     msg['To'] = email
# #                     msg['Subject'] = subject if subject else "No Subject"  # Default subject if not provided
# #
# #                     # Attach the message body based on the selected format
# #                     if self.message_format.get() == 'html':
# #                         msg.attach(MIMEText(message_body, 'html', 'utf-8'))
# #                     else:
# #                         msg.attach(MIMEText(message_body, 'plain', 'utf-8'))
# #
# #                     # Attach selected files
# #                     if attachments:
# #                         attachment_files = attachments.split(',')
# #                         for file in attachment_files:
# #                             file_path = os.path.join('attachments', file.strip())
# #                             if os.path.exists(file_path):
# #                                 with open(file_path, 'rb') as f:
# #                                     part = MIMEBase('application', 'octet-stream')
# #                                     part.set_payload(f.read())
# #                                     encoders.encode_base64(part)
# #                                     part.add_header('Content-Disposition', f'attachment; filename= {file}')
# #                                     msg.attach(part)
# #                             else:
# #                                 print(f"Attachment file '{file}' not found.")
# #
# #                     # Send the email
# #                     server.send_message(msg)
# #                     print(f'Email sent to {email}')
# #
# #                 except KeyError as e:
# #                     print(f"Error: Missing column in the Excel file - {e}")
# #                 except Exception as e:
# #                     print(f"An error occurred: {e}")
# #
# #             # Closing SMTP connection
# #             server.quit()
# #
# #             messagebox.showinfo("Emails Sent", "All emails sent successfully!")
# #
# #         except FileNotFoundError:
# #             messagebox.showerror("File Not Found Error", "The specified file could not be found.")
# #         except smtplib.SMTPException as e:
# #             messagebox.showerror("SMTP Error", f"An SMTP error occurred: {e}")
# #         except Exception as e:
# #             messagebox.showerror("Error", f"An unexpected error occurred: {e}")
# #
# # if __name__ == "__main__":
# #     app = EmailSenderApp()
# #     app.mainloop()


# -*- coding: utf-8 -*-


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

    def load_background_image(self):
        # Find and load the background image
        current_dir = os.path.dirname(sys.argv[0])  # For PyInstaller compatibility
        background_files = glob.glob(os.path.join(current_dir, "background.*"))
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

    # def create_widgets(self):
    #     top_frame = ttk.Frame(self, padding=20, relief='raised')
    #     top_frame.pack(fill='both', expand=True)
    #
    #     # Logo
    #     logo_path = self.find_background_path()
    #     if logo_path and os.path.exists(logo_path):
    #         try:
    #             logo_image = Image.open(logo_path)
    #             logo_image = logo_image.resize((250, 250), Image.LANCZOS)
    #             logo_photo = ImageTk.PhotoImage(logo_image)
    #             logo_label = ttk.Label(top_frame, image=logo_photo, background='white')
    #             logo_label.image = logo_photo
    #             logo_label.grid(row=0, column=0, columnspan=3, pady=10)
    #         except Exception as e:
    #             print(f"Error loading image: {e}")
    #             messagebox.showerror("Error", "Failed to load logo image")
    #
    #     label_subject = ttk.Label(top_frame, text="Subject:", font=("Segoe UI", 16, "bold"))
    #     label_subject.grid(row=1, column=0, pady=20, padx=20, sticky='w')
    #
    #     self.entry_subject = ttk.Entry(top_frame, width=70, font=("Segoe UI", 12))
    #     self.entry_subject.grid(row=1, column=1, pady=20, padx=10, sticky='w')
    #
    #     btn_browse_word = ttk.Button(top_frame, text="Choose Word Template", command=self.browse_word_template,
    #                                  style='TButton')
    #     btn_browse_word.grid(row=2, column=1, pady=10, padx=10, sticky='w')
    #
    #     btn_browse_excel = ttk.Button(top_frame, text="Choose Excel File", command=self.browse_excel_file,
    #                                   style='TButton')
    #     btn_browse_excel.grid(row=3, column=1, pady=10, padx=10, sticky='w')
    #
    #     label_provider = ttk.Label(top_frame, text="Choose Email Provider:", font=("Segoe UI", 16, "bold"))
    #     label_provider.grid(row=4, column=0, pady=20, padx=20, sticky='w')
    #
    #     provider_combobox = ttk.Combobox(top_frame, textvariable=self.email_provider, state='readonly',
    #                                      font=("Segoe UI", 12))
    #     provider_combobox['values'] = ("Gmail", "Outlook", "Webmail")
    #     provider_combobox.grid(row=4, column=1, pady=20, padx=10, sticky='w')
    #     provider_combobox.current(0)
    #
    #     btn_send_emails = ttk.Button(top_frame, text="Send Emails", command=self.send_emails, style='TButton')
    #     btn_send_emails.grid(row=5, column=1, pady=20, padx=10, sticky='w')
    #
    #     btn_exit = ttk.Button(top_frame, text="Exit", command=self.destroy, style='TButton')
    #     btn_exit.grid(row=6, column=1, pady=20, padx=10, sticky='w')

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

    # def send_emails(self):
    #     try:
    #         subject = self.entry_subject.get()
    #         word_path = self.word_template_path
    #         excel_path = self.excel_file_path
    #         provider = self.email_provider.get()
    #
    #         # Validate input
    #         if not subject:
    #             messagebox.showerror("Error", "Please enter a subject")
    #             return
    #
    #         if not word_path or not os.path.exists(word_path):
    #             messagebox.showerror("Error", "Please choose a Word template file")
    #             return
    #
    #         if not excel_path or not os.path.exists(excel_path):
    #             messagebox.showerror("Error", "Please choose an Excel file")
    #             return
    #
    #         # Read Excel file
    #         df = pd.read_excel(excel_path)
    #
    #         # Read message template from Word file
    #         def read_word_template(word_path):
    #             doc = Document(word_path)
    #             full_text = []
    #             for para in doc.paragraphs:
    #                 full_text.append(para.text)
    #             return '\n'.join(full_text)
    #
    #         # SMTP settings
    #         if provider == "Gmail":
    #             smtp_server = 'smtp.gmail.com'
    #             smtp_port = 465
    #             smtp_username = 'islambadran39@gmail.com'  # Update with your email address
    #             smtp_password = 'yrwp xhds hdtp cuqe'
    #         elif provider == "Outlook":
    #             smtp_server = 'smtp-mail.outlook.com'
    #             smtp_port = 587
    #             smtp_username = 'islambadran39@gmail.com'  # ?????? ?????? ?????? ?????????? ????? ?? Outlook
    #             smtp_password = 'jilifwjilhytlubb'
    #         elif provider == "Webmail":
    #             smtp_server = 'mail.una-oic.org'
    #             smtp_port = 465
    #             smtp_username = 'messages@una-oic.org'
    #             smtp_password = '}E~8NLAZ5Ki3'
    #
    #         if provider == "Outlook":
    #             server = smtplib.SMTP(smtp_server, smtp_port)
    #             server.starttls()
    #         else:
    #             server = smtplib.SMTP_SSL(smtp_server, smtp_port)
    #
    #         server.login(smtp_username, smtp_password)
    #
    #         # Function to replace variables in the template with values from the Excel row
    #         def replace_variables(template, row):
    #             pattern = re.compile(r'�(.*?)�')
    #             matches = pattern.findall(template)
    #             for match in matches:
    #                 if match in row:
    #                     value = row[match]
    #                     if pd.isna(value):
    #                         value = ""
    #                     template = template.replace(f'�{match}�', str(value))
    #                 else:
    #                     print(f"Warning: Column '{match}' not found in the Excel file")
    #             return template
    #
    #         # Sending emails
    #         for index, row in df.iterrows():
    #             try:
    #                 email = row.get('email')
    #                 if not email:
    #                     print(f"Skipping row {index + 1}: No email specified")
    #                     continue
    #
    #                 attachments = row.get('attachments')
    #                 if pd.isna(attachments):
    #                     attachments = ''
    #
    #                 # Customize message body using variables from Excel
    #                 message_template = read_word_template(word_path)
    #                 message_body = replace_variables(message_template, row)
    #
    #                 # Create email message
    #                 msg = MIMEMultipart()
    #                 msg['From'] = smtp_username
    #                 msg['To'] = email
    #                 msg['Subject'] = subject if subject else "No Subject"  # Default subject if not provided
    #
    #                 # Attach Arabic text with proper encoding
    #                 msg.attach(MIMEText(message_body, 'plain', 'utf-8'))
    #
    #                 # Attach selected files
    #                 if attachments:
    #                     attachment_files = attachments.split(',')
    #                     for file in attachment_files:
    #                         file_path = os.path.join('attachments', file.strip())
    #                         if os.path.exists(file_path):
    #                             with open(file_path, 'rb') as f:
    #                                 part = MIMEBase('application', 'octet-stream')
    #                                 part.set_payload(f.read())
    #                                 encoders.encode_base64(part)
    #                                 part.add_header('Content-Disposition', f'attachment; filename= {file}')
    #                                 msg.attach(part)
    #                         else:
    #                             print(f"Attachment file '{file}' not found.")
    #
    #                 # Send the email
    #                 server.send_message(msg)
    #                 print(f'Email sent to {email}')
    #
    #             except KeyError as e:
    #                 print(f"Error: Missing column in the Excel file - {e}")
    #             except Exception as e:
    #                 print(f"An error occurred: {e}")
    #
    #         # Closing SMTP connection
    #         server.quit()
    #
    #         messagebox.showinfo("Emails Sent", "All emails sent successfully!")
    #
    #     except FileNotFoundError:
    #         messagebox.showerror("File Error", "File not found, please check the file path")
    #     except Exception as e:
    #         messagebox.showerror("Error", f"An error occurred: {e}")

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
