import tkinter as tk
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_COLOR_INDEX
import os
import time
from tkinter import messagebox
import re
import sys
import glob
from tkinter import *
from tkinter import ttk, filedialog, messagebox, StringVar
from PIL import Image, ImageTk
import base64
import smtplib
from datetime import datetime

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
        self.html_template_path = None
        self.excel_file_path = None
        self.attachment_folder_path = None
        self.email_provider = StringVar()

        self.email_credentials = self.load_email_credentials()
        self.selected_email = StringVar()

        self.create_widgets()

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

    def create_widgets(self):
        label_subject = ttk.Label(self, text="Subject:", font=("Segoe UI", 16, "bold"))
        self.canvas_bg.create_window(20, 320, anchor='w', window=label_subject)

        self.entry_subject = ttk.Entry(self, width=70, font=("Segoe UI", 12))
        self.canvas_bg.create_window(150, 320, anchor='w', window=self.entry_subject)

        btn_browse_word = ttk.Button(self, text="Choose Word Template", command=self.browse_word_template)
        self.canvas_bg.create_window(150, 360, anchor='w', window=btn_browse_word)

        btn_browse_html = ttk.Button(self, text="Choose HTML Template", command=self.browse_html_template)
        self.canvas_bg.create_window(150, 400, anchor='w', window=btn_browse_html)

        btn_browse_excel = ttk.Button(self, text="Choose Excel File", command=self.browse_excel_file)
        self.canvas_bg.create_window(150, 440, anchor='w', window=btn_browse_excel)

        # ✅ زر اختيار مجلد المرفقات بعد زر الإكسل مباشرة
        btn_browse_attachments = ttk.Button(self, text="Choose Attachments Folder",
                                            command=self.browse_attachment_folder)
        self.canvas_bg.create_window(150, 480, anchor='w', window=btn_browse_attachments)

        # 📧 اختيار الإيميل
        label_email = ttk.Label(self, text="Choose Email Account:", font=("Segoe UI", 16, "bold"))
        self.canvas_bg.create_window(20, 520, anchor='w', window=label_email)

        email_combobox = ttk.Combobox(self, textvariable=self.selected_email, state='readonly',
                                      font=("Segoe UI", 12))
        email_combobox['values'] = list(self.email_credentials.keys())
        email_combobox.current(0)
        self.canvas_bg.create_window(150, 520, anchor='w', window=email_combobox)

        # ✅ زر الإرسال
        btn_send_emails = ttk.Button(self, text="Send Emails", command=self.send_emails)
        self.canvas_bg.create_window(150, 600, anchor='w', window=btn_send_emails)

        # ❌ زر الخروج
        btn_exit = ttk.Button(self, text="Exit", command=self.destroy)
        self.canvas_bg.create_window(150, 640, anchor='w', window=btn_exit)

    def load_email_credentials(self):
        credentials = {}
        try:
            with open("credentials.txt", "r") as file:
                content = file.read()
                blocks = content.split("\n\n")
                for block in blocks:
                    lines = block.split("\n")
                    if len(lines) == 3:  # Assuming each block has email, password, and provider
                        encoded_email = lines[0]
                        encoded_password = lines[1]
                        encoded_provider = lines[2]

                        email = base64.b64decode(encoded_email).decode('utf-8')
                        password = base64.b64decode(encoded_password).decode('utf-8')
                        provider = base64.b64decode(encoded_provider).decode('utf-8')

                        credentials[email] = (password, provider)
        except Exception as e:
            print(f"Error loading email credentials: {e}")
            messagebox.showerror("Error", "Failed to load email credentials")
        return credentials

    def browse_word_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if filename:
            self.word_template_path = filename
            messagebox.showinfo("Word Template Selected", f"Selected template: {filename}")

    def browse_html_template(self):
        filename = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
        if filename:
            self.html_template_path = filename
            messagebox.showinfo("HTML Template Selected", f"Selected template: {filename}")

    def browse_excel_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.excel_file_path = filename
            messagebox.showinfo("Excel File Selected", f"Selected file: {filename}")

    def browse_attachment_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.attachment_folder_path = folder
            messagebox.showinfo("Attachments Folder Selected", f"Selected folder: {folder}")

    def send_emails(self):
        sent_emails = []
        max_retries = 3
        retry_delay = 10
        retries = 0

        while retries < max_retries:
            try:
                word_path = self.word_template_path
                html_path = self.html_template_path
                excel_path = self.excel_file_path
                email = self.selected_email.get()
                password, provider = self.email_credentials.get(email, ("", ""))

                if not word_path or not os.path.exists(word_path):
                    messagebox.showerror("Error", "Please choose a Word template file")
                    return
                if not excel_path or not os.path.exists(excel_path):
                    messagebox.showerror("Error", "Please choose an Excel file")
                    return
                if not email or not password:
                    messagebox.showerror("Error", "Please choose a valid email account")
                    return

                if provider == "Gmail":
                    smtp_server = 'smtp.gmail.com'
                    smtp_port = 465
                elif provider == "Outlook":
                    smtp_server = 'smtp-mail.outlook.com'
                    smtp_port = 587
                elif provider == "Webmail":
                    smtp_server = 'mail.una-oic.org'
                    smtp_port = 465
                else:
                    messagebox.showerror("Error", f"Unknown email provider: {provider}")
                    return

                if provider == "Outlook":
                    server = smtplib.SMTP(smtp_server, smtp_port)
                    server.starttls()
                else:
                    server = smtplib.SMTP_SSL(smtp_server, smtp_port)

                server.login(email, password)

                def read_word_template(file_path):
                    try:
                        doc = Document(file_path)
                        html = ""
                        for paragraph in doc.paragraphs:
                            html += '<p>'
                            for run in paragraph.runs:
                                if run.bold: html += '<b>'
                                if run.italic: html += '<i>'
                                if run.underline: html += '<u>'
                                if run.font.color.rgb:
                                    color = run.font.color.rgb
                                    html += f'<span style="color:#{color};">'
                                html += run.text
                                if run.font.color.rgb: html += '</span>'
                                if run.underline: html += '</u>'
                                if run.italic: html += '</i>'
                                if run.bold: html += '</b>'
                            html += '</p>'
                        return html
                    except Exception as e:
                        print(f"Error reading Word template: {e}")
                        return ""

                def replace_variables(template, row):
                    try:
                        pattern = re.compile(r'<<([^<>]+)>>')
                        matches = pattern.findall(template)
                        for match in matches:
                            value = row.get(match, "")
                            if pd.isna(value): value = ""
                            template = template.replace(f'<<{match}>>', str(value))
                        return template
                    except Exception as e:
                        print(f"Error replacing variables: {e}")
                        return template

                def read_html_template(file_path, row):
                    try:
                        with open(file_path, 'r', encoding='utf-8') as file:
                            html_content = file.read()
                            pattern = re.compile(r'<<([^<>]+)>>')
                            matches = pattern.findall(html_content)
                            for match in matches:
                                value = row.get(match, "")
                                if pd.isna(value): value = ""
                                html_content = html_content.replace(f'<<{match}>>', str(value))
                            return html_content
                    except Exception as e:
                        print(f"Error reading HTML template: {e}")
                        return ""

                df = pd.read_excel(excel_path)
                email_counter = 0

                # تفريغ ملف الإيميلات المرسلة
                open("sent_live.txt", "w").close()

                for index, row in df.iterrows():
                    if email_counter >= 50:
                        email_counter = 0
                        server.quit()
                        time.sleep(retry_delay)
                        if provider == "Outlook":
                            server = smtplib.SMTP(smtp_server, smtp_port)
                            server.starttls()
                        else:
                            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
                        server.login(email, password)

                    recipient_emails_raw = row.get('email')
                    if not recipient_emails_raw:
                        print(f"Skipping row {index + 1}: No email specified")
                        continue

                    recipient_emails = [e.strip() for e in str(recipient_emails_raw).split(',') if e.strip()]
                    if not recipient_emails:
                        continue

                    subject_template = self.entry_subject.get()
                    customized_subject = replace_variables(subject_template, row)
                    word_message = replace_variables(read_word_template(word_path), row)
                    html_message = replace_variables(read_html_template(html_path, row), row)
                    html_final = html_message.replace("[[CUSTOM_CONTENT]]", word_message)

                    msg = MIMEMultipart()
                    msg['From'] = email
                    msg['To'] = ", ".join(recipient_emails)
                    msg['Subject'] = customized_subject
                    msg.attach(MIMEText(html_final, 'html'))

                    # ✅ إرفاق المرفقات فقط لو العمود موجود وله قيمة
                    if 'attachments' in df.columns:
                        attachments = row.get('attachments', '')
                        if isinstance(attachments, str) and attachments.strip():
                            for file in attachments.split(','):
                                file = file.strip()
                                if not file:
                                    continue
                                if not self.attachment_folder_path:
                                    messagebox.showerror("Error", "Please choose an attachments folder.")
                                    return

                                path = os.path.join(self.attachment_folder_path, file)
                                if os.path.exists(path):
                                    with open(path, 'rb') as f:
                                        part = MIMEBase('application', 'octet-stream')
                                        part.set_payload(f.read())
                                        encoders.encode_base64(part)
                                        part.add_header(
                                            'Content-Disposition',
                                            f'attachment; filename={os.path.basename(path)}'
                                        )
                                        msg.attach(part)
                                        print(f"📎 Attached: {file}")
                                else:
                                    print(f"❌ Attachment not found: {file}")

                    text = msg.as_string()
                    server.sendmail(email, recipient_emails, text)

                    for r_email in recipient_emails:
                        sent_emails.append({'email': r_email, 'status': 'Sent'})
                        with open("sent_live.txt", "a", encoding="utf-8") as f:
                            f.write(f"{r_email}\n")

                    print(f"✅ Email sent to: {', '.join(recipient_emails)}")
                    email_counter += 1

                server.quit()
                pd.DataFrame(sent_emails).to_excel('sent_emails.xlsx', index=False)
                messagebox.showinfo("Success", "All Emails sent successfully")
                return

            except Exception as e:
                print(f"❌ Error: {e}")
                pd.DataFrame(sent_emails).to_excel('sent_emails.xlsx', index=False)
                messagebox.showerror("Error", "Some Emails Failed to send")
                retries += 1
                time.sleep(retry_delay)

        if retries >= max_retries:
            messagebox.showerror("Error", "Failed to send some or all emails after multiple attempts")


if __name__ == "__main__":
    app = EmailSenderApp()
    app.mainloop()

# import tkinter as tk
# import pandas as pd
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
# from docx import Document
# from docx.shared import RGBColor, Pt
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
# from docx.enum.text import WD_COLOR_INDEX
# import os
# import time
# from tkinter import ttk, filedialog, messagebox, StringVar
# from tkinter import messagebox, filedialog, StringVar
# from PIL import Image, ImageTk
# import base64
# import re
# import sys
# import glob
# from datetime import datetime
#
# class EmailSenderApp(tk.Tk):
#
#     def __init__(self):
#         super().__init__()
#         self.title("Email Sender")
#         self.geometry("800x700")
#         self.configure(bg="white")
#
#         self.canvas_bg = tk.Canvas(self, bg='white', width=800, height=700)
#         self.canvas_bg.pack(fill='both', expand=True)
#
#         self.load_background_image()
#
#         self.style = ttk.Style()
#         self.style.theme_use("clam")
#         self.style.configure('TButton', foreground='white', background='#0078D7', font=('Segoe UI', 10, 'bold'))
#         self.style.configure('TLabel', background='white', foreground='#0078D7', font=('Segoe UI', 10))
#         self.style.configure('TEntry', font=('Segoe UI', 10))
#
#         self.word_template_path = None
#         self.html_template_path = None
#         self.excel_file_path = None
#         self.email_provider = StringVar()
#
#         self.email_credentials = self.load_email_credentials()
#         self.selected_email = StringVar()
#
#         self.create_widgets()
#
#     def load_background_image(self):
#         current_dir = os.path.dirname(sys.argv[0])
#         background_files = glob.glob(os.path.join(current_dir, "logo.*"))
#         if background_files:
#             background_image_path = background_files[0]
#             if os.path.exists(background_image_path):
#                 try:
#                     background_image = Image.open(background_image_path)
#                     background_image = background_image.resize((800, 700), Image.LANCZOS)
#                     self.background_photo = ImageTk.PhotoImage(background_image)
#                     self.canvas_bg.create_image(0, 0, anchor=tk.NW, image=self.background_photo)
#                 except Exception as e:
#                     print(f"Error loading background image: {e}")
#                     messagebox.showerror("Error", "Failed to load background image")
#
#     def create_widgets(self):
#         label_subject = ttk.Label(self, text="Subject:", font=("Segoe UI", 16, "bold"))
#         self.canvas_bg.create_window(20, 320, anchor='w', window=label_subject)
#
#         self.entry_subject = ttk.Entry(self, width=70, font=("Segoe UI", 12))
#         self.canvas_bg.create_window(150, 320, anchor='w', window=self.entry_subject)
#
#         btn_browse_word = ttk.Button(self, text="Choose Word Template", command=self.browse_word_template)
#         self.canvas_bg.create_window(150, 360, anchor='w', window=btn_browse_word)
#
#         btn_browse_html = ttk.Button(self, text="Choose HTML Template", command=self.browse_html_template)
#         self.canvas_bg.create_window(150, 400, anchor='w', window=btn_browse_html)
#
#         btn_browse_excel = ttk.Button(self, text="Choose Excel File", command=self.browse_excel_file)
#         self.canvas_bg.create_window(150, 440, anchor='w', window=btn_browse_excel)
#
#         label_from_email = ttk.Label(self, text="From Email:", font=("Segoe UI", 16, "bold"))
#         self.canvas_bg.create_window(20, 480, anchor='w', window=label_from_email)
#
#         self.entry_from_email = ttk.Entry(self, width=30, font=("Segoe UI", 12))
#         self.canvas_bg.create_window(150, 480, anchor='w', window=self.entry_from_email)
#
#         label_provider = ttk.Label(self, text="Choose Email Provider:", font=("Segoe UI", 16, "bold"))
#         self.canvas_bg.create_window(20, 520, anchor='w', window=label_provider)
#
#         provider_combobox = ttk.Combobox(self, textvariable=self.email_provider, state='readonly', font=("Segoe UI", 12))
#         provider_combobox['values'] = ("Gmail", "Outlook", "Webmail")
#         provider_combobox.current(0)
#         self.canvas_bg.create_window(150, 520, anchor='w', window=provider_combobox)
#
#         btn_send_emails = ttk.Button(self, text="Send Emails", command=self.send_emails)
#         self.canvas_bg.create_window(150, 600, anchor='w', window=btn_send_emails)
#
#         btn_exit = ttk.Button(self, text="Exit", command=self.destroy)
#         self.canvas_bg.create_window(150, 640, anchor='w', window=btn_exit)
#
#     def load_email_credentials(self):
#         credentials = {}
#         try:
#             with open("credentials.txt", "r") as file:
#                 content = file.read()
#                 blocks = content.split("\n\n")
#                 for block in blocks:
#                     lines = block.split("\n")
#                     if len(lines) == 3:
#                         encoded_email = lines[0]
#                         encoded_password = lines[1]
#                         encoded_provider = lines[2]
#
#                         email = base64.b64decode(encoded_email).decode('utf-8')
#                         password = base64.b64decode(encoded_password).decode('utf-8')
#                         provider = base64.b64decode(encoded_provider).decode('utf-8')
#
#                         credentials[email] = (password, provider)
#         except Exception as e:
#             print(f"Error loading email credentials: {e}")
#             messagebox.showerror("Error", "Failed to load email credentials")
#         return credentials
#
#     def browse_word_template(self):
#         filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
#         if filename:
#             self.word_template_path = filename
#             messagebox.showinfo("Word Template Selected", f"Selected template: {filename}")
#
#     def browse_html_template(self):
#         filename = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
#         if filename:
#             self.html_template_path = filename
#             messagebox.showinfo("HTML Template Selected", f"Selected template: {filename}")
#
#     def browse_excel_file(self):
#         filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
#         if filename:
#             self.excel_file_path = filename
#             messagebox.showinfo("Excel File Selected", f"Selected file: {filename}")
#
#     def Gmail_send(self, smtp_username, smtp_password, from_email, to_email, subject, body):
#         smtp_server = 'smtp.gmail.com'
#         smtp_port = 587
#         msg = MIMEMultipart()
#         msg['From'] = from_email
#         msg['To'] = to_email
#         msg['Subject'] = subject
#         msg.attach(MIMEText(body, 'plain'))
#
#         try:
#             with smtplib.SMTP(smtp_server, smtp_port) as server:
#                 server.starttls()
#                 server.login(smtp_username, smtp_password)
#                 server.send_message(msg)
#                 print(f"Email sent from {from_email} to {to_email}")
#         except Exception as e:
#             print(f"Failed to send email to {to_email}: {e}")
#
#     def send_emails(self):
#         try:
#             excel_path = self.excel_file_path
#             email = self.selected_email.get()
#             password, provider = self.email_credentials.get(email, ("", ""))
#
#             if not excel_path or not os.path.exists(excel_path):
#                 messagebox.showerror("Error", "Please choose an Excel file")
#                 return
#
#             df = pd.read_excel(excel_path)
#
#             for index, row in df.iterrows():
#                 recipient_email = row.get('email')
#                 if not recipient_email:
#                     print(f"Skipping row {index + 1}: No email specified")
#                     continue
#
#                 subject_template = self.entry_subject.get()
#                 subject = subject_template.replace("<<email>>", recipient_email)
#
#                 from_email = self.entry_from_email.get()
#                 body = "This is a test email."
#
#                 self.Gmail_send(email, password, from_email, recipient_email, subject, body)
#
#             messagebox.showinfo("Success", "All Emails sent successfully")
#
#         except Exception as e:
#             print(f"Error: {e}")
#             messagebox.showerror("Error", "Some Emails Failed to send")
#
# if __name__ == "__main__":
#     app = EmailSenderApp()
#     app.mainloop()
