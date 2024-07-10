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
    sender_email = 'islambadran39@gmail.com'  # Update with your email address
    sender_password = "tbaz szvk wyrv rshk"  # Update with your email password
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



# import pandas as pd
# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from email.mime.base import MIMEBase
# from email import encoders
# import re
# from docx import Document
# import os
#
# # ????? ??? Excel
# df = pd.read_excel('contacts.xlsx')
#
# # ??????? ???? SMTP
# smtp_server = 'smtp-mail.outlook.com'
# smtp_port = 587
# smtp_username = 'islambadran39@gmail.com'  # ?????? ?????? ?????? ?????????? ????? ?? Outlook
# smtp_password = 'jilifwjilhytlubb'
# '}E~8NLAZ5Ki3'# ?????? ????? ?????? ?????? ????? Outlook
#
#
# # ????? ???? ??????? ?? ??? Word
# def read_word_template(file_path):
#     doc = Document(file_path)
#     full_text = []
#     for para in doc.paragraphs:
#         full_text.append(para.text)
#     return '\n'.join(full_text)
#
#
# message_template = read_word_template('messages-ar.docx')
#
# # ????? ????? SMTP
# server = smtplib.SMTP(smtp_server, smtp_port)
# server.starttls()
# server.login(smtp_username, smtp_password)
#
#
# # ???? ???????? ????????? ?? ???? ?????? ?? ???? ??????
# def replace_variables(template, row):
#     pattern = re.compile(r'«(.*?)»')
#     matches = pattern.findall(template)
#     for match in matches:
#         if match in row:
#             value = row[match]
#             if pd.isna(value):
#                 value = ""
#             template = template.replace(f'«{match}»', str(value))
#         else:
#             print(f"Warning: Column '{match}' not found in the Excel file")
#     return template
#
#
# # ????? ??????? ???????????
# for index, row in df.iterrows():
#     try:
#         email = row['email']
#         attachments = row['attachments']  # Assuming 'attachments' column contains filenames separated by commas
#
#         # ????? ?? ??????? ???????? ????????? ?? ??? Excel
#         message_body = replace_variables(message_template, row)
#
#         # ????? ????? ?????? ??????????
#         msg = MIMEMultipart()
#         msg['From'] = smtp_username
#         msg['To'] = email
#         msg['Subject'] = 'islam elgamed'
#
#         # ????? ???? ?????? ???????? ??????
#         msg.attach(MIMEText(message_body, 'plain', 'utf-8'))
#
#         # ????? ??????? ???????
#         if attachments:
#             attachment_files = attachments.split(',')
#             for file in attachment_files:
#                 file_path = os.path.join('attachments', file.strip())
#                 if os.path.exists(file_path):
#                     with open(file_path, 'rb') as f:
#                         part = MIMEBase('application', 'octet-stream')
#                         part.set_payload(f.read())
#                         encoders.encode_base64(part)
#                         part.add_header('Content-Disposition', f'attachment; filename= {file}')
#                         msg.attach(part)
#                 else:
#                     print(f"Attachment file '{file}' not found.")
#
#         # ????? ??????? ???????????
#         server.send_message(msg)
#         print(f'Email sent to {email}')
#
#     except KeyError as e:
#         print(f"Error: Missing column in the Excel file - {e}")
#     except Exception as e:
#         print(f"An error occurred: {e}")
#
# # ????? ????? SMTP
# server.quit()
#
# print('All emails sent successfully!')

# import pandas as pd
# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from email.mime.base import MIMEBase
# from email import encoders
# import re
# from docx import Document
# import os
# import tkinter as tk
# from tkinter import filedialog
#
# # Function to send emails
# def send_emails(subject):
#     # ????? ??? Excel
#     df = pd.read_excel('contacts.xlsx')
#
#     # ??????? ???? SMTP
#     smtp_server = 'mail.una-oic.org'
#     smtp_port = 465  # Use port 465 for SSL/TLS
#     smtp_username = 'messages@una-oic.org'
#     smtp_password = '}E~8NLAZ5Ki3'
#
#     # ????? ???? ??????? ?? ??? Word
#     def read_word_template(file_path):
#         doc = Document(file_path)
#         full_text = []
#         for para in doc.paragraphs:
#             full_text.append(para.text)
#         return '\n'.join(full_text)
#
#     message_template = read_word_template('messages-ar.docx')
#
#     # ????? ????? SMTP ???????? SSL
#     server = smtplib.SMTP_SSL(smtp_server, smtp_port)
#     server.login(smtp_username, smtp_password)
#
#     # ???? ???????? ????????? ?? ???? ?????? ?? ???? ??????
#     def replace_variables(template, row):
#         pattern = re.compile(r'«(.*?)»')
#         matches = pattern.findall(template)
#         for match in matches:
#             if match in row:
#                 value = row[match]
#                 if pd.isna(value):
#                     value = ""
#                 template = template.replace(f'«{match}»', str(value))
#             else:
#                 print(f"Warning: Column '{match}' not found in the Excel file")
#         return template
#
#     # ????? ??????? ???????????
#     for index, row in df.iterrows():
#         try:
#             email = row['email']
#             attachments = row['attachments']  # Assuming 'attachments' column contains filenames separated by commas
#
#             # ????? ?? ??????? ???????? ????????? ?? ??? Excel
#             message_body = replace_variables(message_template, row)
#
#             # ????? ????? ?????? ??????????
#             msg = MIMEMultipart()
#             msg['From'] = smtp_username
#             msg['To'] = email
#             msg['Subject'] = subject  # Use the subject input by the user
#
#             # ????? ???? ?????? ???????? ??????
#             msg.attach(MIMEText(message_body, 'plain', 'utf-8'))
#
#             # ????? ??????? ???????
#             if attachments:
#                 attachment_files = attachments.split(',')
#                 for file in attachment_files:
#                     file_path = os.path.join('attachments', file.strip())
#                     if os.path.exists(file_path):
#                         with open(file_path, 'rb') as f:
#                             part = MIMEBase('application', 'octet-stream')
#                             part.set_payload(f.read())
#                             encoders.encode_base64(part)
#                             part.add_header('Content-Disposition', f'attachment; filename= {file}')
#                             msg.attach(part)
#                     else:
#                         print(f"Attachment file '{file}' not found.")
#
#             # ????? ??????? ???????????
#             server.send_message(msg)
#             print(f'Email sent to {email}')
#
#         except KeyError as e:
#             print(f"Error: Missing column in the Excel file - {e}")
#         except Exception as e:
#             print(f"An error occurred: {e}")
#
#     # ????? ????? SMTP
#     server.quit()
#
#     print('All emails sent successfully!')
#
# # Function to handle GUI and send emails
# def send_emails_gui():
#     def browse_file():
#         filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
#         if filename:
#             entry_filepath.delete(0, tk.END)
#             entry_filepath.insert(0, filename)
#
#     def send_emails_from_gui():
#         subject = entry_subject.get()
#         send_emails(subject)
#
#     # Create GUI window
#     root = tk.Tk()
#     root.title("Email Sender")
#     root.geometry("400x200")
#
#     # Subject input
#     label_subject = tk.Label(root, text="Subject:")
#     label_subject.pack()
#     entry_subject = tk.Entry(root)
#     entry_subject.pack()
#
#     # Filepath input
#     label_filepath = tk.Label(root, text="Excel File:")
#     label_filepath.pack()
#     entry_filepath = tk.Entry(root)
#     entry_filepath.pack()
#     btn_browse = tk.Button(root, text="Browse", command=browse_file)
#     btn_browse.pack()
#
#     # Send button
#     btn_send = tk.Button(root, text="Send Emails", command=send_emails_from_gui)
#     btn_send.pack()
#
#     root.mainloop()
#
# # Run GUI
# send_emails_gui()


# import tkinter as tk
# from tkinter import ttk, filedialog, messagebox
# import pandas as pd
# import os
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
# from docx import Document
# import re
# from PIL import Image, ImageTk
# import glob
# import sys
#
# class EmailSenderApp(tk.Tk):
#     def __init__(self):
#         super().__init__()
#         self.title("Email Sender")
#         self.geometry("550x600")
#         self.configure(bg="#282828")
#         self.style = ttk.Style()
#         self.style.theme_use("clam")
#         self.style.configure('TButton', background='blue', foreground='white')
#         self.create_widgets()
#
#
#     def create_widgets(self):
#         label = tk.Label(self, text="Welcome to Email Sender", font=("Arial", 16), bg="#282828", fg="white")
#         label.pack(pady=15)
#
#         current_dir = os.path.dirname(sys.argv[0])
#         logo_files = glob.glob(os.path.join(current_dir, "logo.*"))
#         if logo_files:
#             logo_path = logo_files[0]
#             try:
#                 logo_image = Image.open(logo_path)
#                 logo_image = logo_image.resize((260, 260), Image.LANCZOS)
#                 logo_photo = ImageTk.PhotoImage(logo_image)
#                 logo_label = tk.Label(self, image=logo_photo, bg="#282828")
#                 logo_label.image = logo_photo
#                 logo_label.pack()
#             except Exception as e:
#                 print(f"Error loading image: {e}")
#                 messagebox.showerror("Error", "Failed to load logo image")
#         else:
#             print("Logo image not found")
#
#         # Button to browse for Excel file
#         btn_browse = tk.Button(self, text="Browse Excel File", command=self.browse_excel_file, bg="#006400", fg="white", font=("Arial", 14))
#         btn_browse.pack(pady=15)
#
#         # Button to send emails
#         btn_send_emails = tk.Button(self, text="Send Emails", command=self.send_emails, bg="blue", fg="white", font=("Arial", 14))
#         btn_send_emails.pack(pady=15)
#
#         # Button to exit application
#         btn_exit = tk.Button(self, text="Exit", command=self.destroy, bg="red", fg="white", font=("Arial", 14))
#         btn_exit.pack(pady=15)
#
#     def browse_excel_file(self):
#         filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
#         if filename:
#             messagebox.showinfo("Excel File Selected", f"Selected file: {filename}")
#
#     def send_emails(self):
#         try:
#             # ????? ??? Excel
#             df = pd.read_excel('contacts.xlsx')
#
#             # ??????? ???? SMTP
#             smtp_server = 'mail.una-oic.org'
#             smtp_port = 465  # Use port 465 for SSL/TLS
#             smtp_username = 'messages@una-oic.org'
#             smtp_password = '}E~8NLAZ5Ki3'
#
#             # ????? ???? ??????? ?? ??? Word
#             def read_word_template(file_path):
#                 doc = Document(file_path)
#                 full_text = []
#                 for para in doc.paragraphs:
#                     full_text.append(para.text)
#                 return '\n'.join(full_text)
#
#             message_template = read_word_template('messages-ar.docx')
#
#             # ????? ????? SMTP ???????? SSL
#             server = smtplib.SMTP_SSL(smtp_server, smtp_port)
#             server.login(smtp_username, smtp_password)
#
#             # ???? ???????? ????????? ?? ???? ?????? ?? ???? ??????
#             def replace_variables(template, row):
#                 pattern = re.compile(r'«(.*?)»')
#                 matches = pattern.findall(template)
#                 for match in matches:
#                     if match in row:
#                         value = row[match]
#                         if pd.isna(value):
#                             value = ""
#                         template = template.replace(f'«{match}»', str(value))
#                     else:
#                         print(f"Warning: Column '{match}' not found in the Excel file")
#                 return template
#
#             # ????? ??????? ???????????
#             for index, row in df.iterrows():
#                 try:
#                     email = row['email']
#                     attachments = row['attachments']  # Assuming 'attachments' column contains filenames separated by commas
#
#                     # ????? ?? ??????? ???????? ????????? ?? ??? Excel
#                     message_body = replace_variables(message_template, row)
#
#                     # ????? ????? ?????? ??????????
#                     msg = MIMEMultipart()
#                     msg['From'] = smtp_username
#                     msg['To'] = email
#                     msg['Subject'] = "Subject of Email"  # Default subject (can be modified in GUI)
#
#                     # ????? ???? ?????? ???????? ??????
#                     msg.attach(MIMEText(message_body, 'plain', 'utf-8'))
#
#                     # ????? ??????? ???????
#                     if attachments:
#                         attachment_files = attachments.split(',')
#                         for file in attachment_files:
#                             file_path = os.path.join('attachments', file.strip())
#                             if os.path.exists(file_path):
#                                 with open(file_path, 'rb') as f:
#                                     part = MIMEBase('application', 'octet-stream')
#                                     part.set_payload(f.read())
#                                     encoders.encode_base64(part)
#                                     part.add_header('Content-Disposition', f'attachment; filename= {file}')
#                                     msg.attach(part)
#                             else:
#                                 print(f"Attachment file '{file}' not found.")
#
#                     # ????? ??????? ???????????
#                     server.send_message(msg)
#                     print(f'Email sent to {email}')
#
#                 except KeyError as e:
#                     print(f"Error: Missing column in the Excel file - {e}")
#                 except Exception as e:
#                     print(f"An error occurred: {e}")
#
#             # ????? ????? SMTP
#             server.quit()
#
#             messagebox.showinfo("Emails Sent", "All emails sent successfully!")
#
#         except FileNotFoundError:
#             messagebox.showerror("File Error", "Excel file not found.")
#         except Exception as e:
#             messagebox.showerror("Error", f"An error occurred: {e}")
#
# # Run the application
# if __name__ == "__main__":
#     app = EmailSenderApp()
#
#     app.mainloop()


# import tkinter as tk
# from tkinter import ttk, filedialog, messagebox
# import pandas as pd
# import os
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
# from docx import Document
# import re
# from PIL import Image, ImageTk
# import glob
# import sys
#
# class EmailSenderApp(tk.Tk):
#
#     def __init__(self):
#         super().__init__()
#         self.title("Email Sender")
#         self.geometry("800x700")
#         self.configure(bg="white")
#         self.style = ttk.Style()
#         self.style.theme_use("clam")
#         self.style.configure('TButton', background='#0078D7', foreground='white', font=('Segoe UI', 10, 'bold'))
#         self.style.configure('TLabel', background='white', foreground='#0078D7', font=('Segoe UI', 10))
#         self.style.configure('TEntry', font=('Segoe UI', 10))
#         self.style.configure('TFrame', background='white')
#
#         self.word_template_path = None
#         self.excel_file_path = None
#
#         self.create_widgets()
#
#     def find_logo_path(self):
#         current_dir = os.path.dirname(os.path.abspath(__file__))
#         logo_files = glob.glob(os.path.join(current_dir, "logo.*"))
#         if logo_files:
#             return logo_files[0]
#         else:
#             return None
#
#     def create_widgets(self):
#         top_frame = ttk.Frame(self, padding=20, relief='raised')
#         top_frame.pack(fill='both', expand=True)
#
#         # Logo
#         logo_path = self.find_logo_path()
#         if os.path.exists(logo_path):
#             logo_files = glob.glob(logo_path)
#             if logo_files:
#                 logo_file = logo_files[0]
#                 try:
#                     logo_image = Image.open(logo_file)
#                     logo_image = logo_image.resize((250, 250), Image.LANCZOS)
#                     logo_photo = ImageTk.PhotoImage(logo_image)
#                     logo_label = ttk.Label(top_frame, image=logo_photo, background='white')
#                     logo_label.image = logo_photo
#                     logo_label.grid(row=0, column=0, columnspan=3, pady=10)
#                 except Exception as e:
#                     print(f"Error loading image: {e}")
#                     messagebox.showerror("Error", "Failed to load logo image")
#             else:
#                 print("Logo image not found")
#
#         label_subject = ttk.Label(top_frame, text="Subject:", font=("Segoe UI", 16, "bold"))
#         label_subject.grid(row=1, column=0, pady=20, padx=20, sticky='w')
#
#         self.entry_subject = ttk.Entry(top_frame, width=70, font=("Segoe UI", 12))
#         self.entry_subject.grid(row=1, column=1, pady=20, padx=10, sticky='w')
#
#         btn_browse_word = ttk.Button(top_frame, text="Choose Word Template", command=self.browse_word_template,
#                                      style='TButton')
#         btn_browse_word.grid(row=2, column=1, pady=10, padx=10, sticky='w')
#
#         btn_browse_excel = ttk.Button(top_frame, text="Choose Excel File", command=self.browse_excel_file,
#                                       style='TButton')
#         btn_browse_excel.grid(row=3, column=1, pady=10, padx=10, sticky='w')
#
#         btn_send_emails = ttk.Button(top_frame, text="Send Emails", command=self.send_emails, style='TButton')
#         btn_send_emails.grid(row=4, column=1, pady=20, padx=10, sticky='w')
#
#         btn_exit = ttk.Button(top_frame, text="Exit", command=self.destroy, style='TButton')
#         btn_exit.grid(row=5, column=1, pady=20, padx=10, sticky='w')
#
#     def browse_word_template(self):
#         filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
#         if filename:
#             self.word_template_path = filename
#             messagebox.showinfo("Word Template Selected", f"Selected template: {filename}")
#
#     def browse_excel_file(self):
#         filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
#         if filename:
#             self.excel_file_path = filename
#             messagebox.showinfo("Excel File Selected", f"Selected file: {filename}")
#
#     # def send_emails(self):
#     #     try:
#     #         subject = self.entry_subject.get()
#     #         word_path = self.word_template_path
#     #         excel_path = self.excel_file_path
#     #
#     #         # Validate input
#     #         if not subject:
#     #             messagebox.showerror("Error", "Please enter a subject")
#     #             return
#     #
#     #         if not word_path or not os.path.exists(word_path):
#     #             messagebox.showerror("Error", "Please choose a Word template file")
#     #             return
#     #
#     #         if not excel_path or not os.path.exists(excel_path):
#     #             messagebox.showerror("Error", "Please choose an Excel file")
#     #             return
#     #
#     #         # Read Excel file
#     #         df = pd.read_excel(excel_path)
#     #
#     #         # SMTP settings
#     #         smtp_server = 'mail.una-oic.org'
#     #         smtp_port = 465  # Use port 465 for SSL/TLS
#     #         smtp_username = 'messages@una-oic.org'
#     #         smtp_password = '}E~8NLAZ5Ki3'
#     #
#     #         # Read message template from Word file
#     #         def read_word_template(word_path):
#     #             doc = Document(word_path)
#     #             full_text = []
#     #             for para in doc.paragraphs:
#     #                 full_text.append(para.text)
#     #             return '\n'.join(full_text)
#     #
#     #         server = smtplib.SMTP_SSL(smtp_server, smtp_port)
#     #         server.login(smtp_username, smtp_password)
#     #
#     #         # Function to replace variables in the template with values from the Excel row
#     #         def replace_variables(template, row):
#     #             pattern = re.compile(r'«(.*?)»')
#     #             matches = pattern.findall(template)
#     #             for match in matches:
#     #                 if match in row:
#     #                     value = row[match]
#     #                     if pd.isna(value):
#     #                         value = ""
#     #                     template = template.replace(f'«{match}»', str(value))
#     #                 else:
#     #                     print(f"Warning: Column '{match}' not found in the Excel file")
#     #             return template
#     #
#     #         # Sending emails
#     #         for index, row in df.iterrows():
#     #             try:
#     #                 email = row['email']
#     #                 attachments = row['attachments']  # Assuming 'attachments' column contains filenames separated by commas
#     #
#     #                 # Customize message body using variables from Excel
#     #                 message_template = read_word_template(word_path)
#     #                 message_body = replace_variables(message_template, row)
#     #
#     #                 # Create email message
#     #                 msg = MIMEMultipart()
#     #                 msg['From'] = smtp_username
#     #                 msg['To'] = email
#     #                 msg['Subject'] = subject if subject else "No Subject"  # Default subject if not provided
#     #
#     #                 # Attach Arabic text with proper encoding
#     #                 msg.attach(MIMEText(message_body, 'plain', 'utf-8'))
#     #
#     #                 # Attach selected files
#     #                 if attachments:
#     #                     attachment_files = attachments.split(',')
#     #                     for file in attachment_files:
#     #                         file_path = os.path.join('attachments', file.strip())
#     #                         if os.path.exists(file_path):
#     #                             with open(file_path, 'rb') as f:
#     #                                 part = MIMEBase('application', 'octet-stream')
#     #                                 part.set_payload(f.read())
#     #                                 encoders.encode_base64(part)
#     #                                 part.add_header('Content-Disposition', f'attachment; filename= {file}')
#     #                                 msg.attach(part)
#     #                         else:
#     #                             print(f"Attachment file '{file}' not found.")
#     #
#     #                 # Send the email
#     #                 server.send_message(msg)
#     #                 print(f'Email sent to {email}')
#     #
#     #             except KeyError as e:
#     #                 print(f"Error: Missing column in the Excel file - {e}")
#     #             except Exception as e:
#     #                 print(f"An error occurred: {e}")
#     #
#     #         # Closing SMTP connection
#     #         server.quit()
#     #
#     #         messagebox.showinfo("Emails Sent", "All emails sent successfully!")
#     #
#     #     except FileNotFoundError:
#     #         messagebox.showerror("File Error", "File not found.")
#     #     except Exception as e:
#     #         messagebox.showerror("Error", f"An error occurred: {e}")
#     def send_emails(self):
#         try:
#             subject = self.entry_subject.get()
#             word_path = self.word_template_path
#             excel_path = self.excel_file_path
#
#             # Validate input
#             if not subject:
#                 messagebox.showerror("Error", "Please enter a subject")
#                 return
#
#             if not word_path or not os.path.exists(word_path):
#                 messagebox.showerror("Error", "Please choose a Word template file")
#                 return
#
#             if not excel_path or not os.path.exists(excel_path):
#                 messagebox.showerror("Error", "Please choose an Excel file")
#                 return
#
#             # Read Excel file
#             df = pd.read_excel(excel_path)
#
#             # SMTP settings
#             smtp_server = 'mail.una-oic.org'
#             smtp_port = 465  # Use port 465 for SSL/TLS
#             smtp_username = 'messages@una-oic.org'
#             smtp_password = '}E~8NLAZ5Ki3'
#
#             # Read message template from Word file
#             def read_word_template(word_path):
#                 doc = Document(word_path)
#                 full_text = []
#                 for para in doc.paragraphs:
#                     full_text.append(para.text)
#                 return '\n'.join(full_text)
#
#             server = smtplib.SMTP_SSL(smtp_server, smtp_port)
#             server.login(smtp_username, smtp_password)
#
#             # Function to replace variables in the template with values from the Excel row
#             def replace_variables(template, row):
#                 pattern = re.compile(r'«(.*?)»')
#                 matches = pattern.findall(template)
#                 for match in matches:
#                     if match in row:
#                         value = row[match]
#                         if pd.isna(value):
#                             value = ""
#                         template = template.replace(f'«{match}»', str(value))
#                     else:
#                         print(f"Warning: Column '{match}' not found in the Excel file")
#                 return template
#
#             # Sending emails
#             for index, row in df.iterrows():
#                 try:
#                     email = row['email']
#                     attachments = row[
#                         'attachments']  # Assuming 'attachments' column contains filenames separated by commas
#
#                     # Skip rows without attachments
#                     if pd.isna(attachments) or not attachments.strip():
#                         print(f"Skipping row {index}: No attachments specified")
#                         continue
#
#                     # Customize message body using variables from Excel
#                     message_template = read_word_template(word_path)
#                     message_body = replace_variables(message_template, row)
#
#                     # Create email message
#                     msg = MIMEMultipart()
#                     msg['From'] = smtp_username
#                     msg['To'] = email
#                     msg['Subject'] = subject if subject else "No Subject"  # Default subject if not provided
#
#                     # Attach Arabic text with proper encoding
#                     msg.attach(MIMEText(message_body, 'plain', 'utf-8'))
#
#                     # Attach selected files
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
#                     # Send the email
#                     server.send_message(msg)
#                     print(f'Email sent to {email}')
#
#                 except KeyError as e:
#                     print(f"Error: Missing column in the Excel file - {e}")
#                 except Exception as e:
#                     print(f"An error occurred: {e}")
#
#             # Closing SMTP connection
#             server.quit()
#
#             messagebox.showinfo("Emails Sent", "All emails sent successfully!")
#
#         except FileNotFoundError:
#             messagebox.showerror("File Error", "File not found.")
#         except Exception as e:
#             messagebox.showerror("Error", f"An error occurred: {e}")
#
#
# # Run the application
# if __name__ == "__main__":
#     app = EmailSenderApp()
#     app.mainloop()