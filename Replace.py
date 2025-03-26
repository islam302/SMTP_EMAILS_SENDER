# import os
# import shutil
# import pandas as pd
# from docx import Document
# from docx2pdf import convert
# from tkinter import *
# from tkinter import ttk, messagebox, filedialog
# from PIL import Image, ImageTk
# import glob
# import sys
#
# class Replace(Tk):
#
#     def __init__(self):
#         super().__init__()
#         self.title("Create Pdf files")
#         self.geometry("800x600")
#         self.configure(bg="#F0F0F0")
#
#         self.load_background_image()
#
#         self.style = ttk.Style()
#         self.style.theme_use("clam")
#         self.style.configure('TButton', foreground='white', background='#0078D7', font=('Segoe UI', 12, 'bold'))
#         self.style.map('TButton', background=[('active', '#005A9E')])
#         self.style.configure('TLabel', background='#F0F0F0', foreground='#0078D7', font=('Segoe UI', 12))
#         self.style.configure('TEntry', font=('Segoe UI', 12))
#         self.style.configure('TCheckbutton', background='#F0F0F0', foreground='#0078D7', font=('Segoe UI', 12))
#         self.style.map('TCheckbutton', background=[('active', '#005A9E')], foreground=[('active', 'white')])
#
#         self.style.configure('Red.TButton', foreground='white', background='red', font=('Segoe UI', 12, 'bold'))
#         self.style.map('Red.TButton', background=[('active', '#8B0000')])
#
#         self.word_template_path = None
#         self.excel_file_path = None
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
#                     background_image = background_image.resize((800, 600), Image.LANCZOS)
#                     self.background_photo = ImageTk.PhotoImage(background_image)
#                     self.canvas_bg = Canvas(self, width=800, height=600)
#                     self.canvas_bg.create_image(0, 0, anchor='nw', image=self.background_photo)
#                     self.canvas_bg.pack(fill='both', expand=True)
#                 except Exception:
#                     messagebox.showerror("Error", "Failed to load background image")
#         else:
#             print("Background image not found")
#
#     def create_widgets(self):
#         btn_browse_word = ttk.Button(self, text="WORD FILE", command=self.browse_word_template)
#         self.canvas_bg.create_window(400, 160, anchor='center', window=btn_browse_word)
#
#         btn_browse_excel = ttk.Button(self, text="EXCEL FILE", command=self.browse_excel_file)
#         self.canvas_bg.create_window(400, 220, anchor='center', window=btn_browse_excel)
#
#         self.word_var = BooleanVar()
#         self.pdf_var = BooleanVar()
#
#         chk_word = ttk.Checkbutton(self, text="Create Word File", variable=self.word_var)
#         self.canvas_bg.create_window(340, 280, anchor='center', window=chk_word)
#
#         chk_pdf = ttk.Checkbutton(self, text="Create PDF File", variable=self.pdf_var)
#         self.canvas_bg.create_window(480, 280, anchor='center', window=chk_pdf)
#
#         btn_generate_files = ttk.Button(self, text="Create Files", command=self.generate_files)
#         self.canvas_bg.create_window(400, 340, anchor='center', window=btn_generate_files)
#
#         btn_exit = ttk.Button(self, text="EXIT", style='Red.TButton', command=self.destroy)
#         self.canvas_bg.create_window(400, 400, anchor='center', window=btn_exit)
#
#     def browse_word_template(self):
#         self.word_template_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
#         if self.word_template_path:
#             messagebox.showinfo("Selected File", f"Word Template: {self.word_template_path}")
#
#     def browse_excel_file(self):
#         self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
#         if self.excel_file_path:
#             messagebox.showinfo("Selected File", f"Excel File: {self.excel_file_path}")
#
#     def generate_files(self):
#         if not self.word_template_path or not self.excel_file_path:
#             messagebox.showerror("Error", "Please select both a Word template and an Excel file.")
#             return
#
#         if not (self.word_var.get() or self.pdf_var.get()):
#             messagebox.showerror("Error", "Please select at least one file type to create.")
#             return
#
#         # Create results directory if it doesn't exist
#         results_dir = "results"
#         if not os.path.exists(results_dir):
#             os.makedirs(results_dir)
#
#         # Load the Excel file
#         df = pd.read_excel(self.excel_file_path)
#
#         # Iterate over rows in the Excel file
#         for _, row in df.iterrows():
#             file_name = row.get('file name')
#             if not file_name:
#                 continue
#
#             # Copy the Word template to the results directory with the specified filename
#             file_path = os.path.join(results_dir, f"{file_name}.docx")
#             shutil.copy(self.word_template_path, file_path)
#
#             # Open the copied Word file and replace variables
#             doc = Document(file_path)
#             for para in doc.paragraphs:
#                 # Process text runs in paragraphs to replace placeholders without changing formatting
#                 for run in para.runs:
#                     for key, value in row.items():
#                         if pd.notna(value):
#                             placeholder = f'<<{key}>>'
#                             if placeholder in run.text:
#                                 run.text = run.text.replace(placeholder, str(value))
#
#             # Save the modified document
#             doc.save(file_path)
#
#             # Convert to PDF if the option is selected
#             if self.pdf_var.get():
#                 pdf_path = os.path.splitext(file_path)[0] + '.pdf'
#                 convert(file_path, pdf_path)
#
#         messagebox.showinfo("Success", "Files generated successfully")
#
# if __name__ == "__main__":
#     app = Replace()
#     app.mainloop()



import os
import shutil
import pandas as pd
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import glob
import sys

class Replace(Tk):

    def __init__(self):
        super().__init__()
        self.title("Create Pdf files")
        self.geometry("800x600")
        self.configure(bg="#F0F0F0")

        self.load_background_image()

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure('TButton', foreground='white', background='#0078D7', font=('Segoe UI', 12, 'bold'))
        self.style.map('TButton', background=[('active', '#005A9E')])
        self.style.configure('TLabel', background='#F0F0F0', foreground='#0078D7', font=('Segoe UI', 12))
        self.style.configure('TEntry', font=('Segoe UI', 12))
        self.style.configure('TCheckbutton', background='#F0F0F0', foreground='#0078D7', font=('Segoe UI', 12))
        self.style.map('TCheckbutton', background=[('active', '#005A9E')], foreground=[('active', 'white')])

        self.style.configure('Red.TButton', foreground='white', background='red', font=('Segoe UI', 12, 'bold'))
        self.style.map('Red.TButton', background=[('active', '#8B0000')])

        self.word_template_path = None
        self.excel_file_path = None

        self.create_widgets()

    def load_background_image(self):
        current_dir = os.path.dirname(sys.argv[0])
        background_files = glob.glob(os.path.join(current_dir, "logo.*"))
        if background_files:
            background_image_path = background_files[0]
            if os.path.exists(background_image_path):
                try:
                    background_image = Image.open(background_image_path)
                    background_image = background_image.resize((800, 600), Image.LANCZOS)
                    self.background_photo = ImageTk.PhotoImage(background_image)
                    self.canvas_bg = Canvas(self, width=800, height=600)
                    self.canvas_bg.create_image(0, 0, anchor='nw', image=self.background_photo)
                    self.canvas_bg.pack(fill='both', expand=True)
                except Exception:
                    messagebox.showerror("Error", "Failed to load background image")
        else:
            print("Background image not found")

    def create_widgets(self):
        y_offset = 30  # ???? ??????? ????????

        btn_browse_word = ttk.Button(self, text="WORD FILE", command=self.browse_word_template)
        self.canvas_bg.create_window(400, 180 + y_offset, anchor='center', window=btn_browse_word)

        btn_browse_excel = ttk.Button(self, text="EXCEL FILE", command=self.browse_excel_file)
        self.canvas_bg.create_window(400, 240 + y_offset, anchor='center', window=btn_browse_excel)

        self.word_var = BooleanVar()
        self.pdf_var = BooleanVar()

        chk_word = ttk.Checkbutton(self, text="Create Word File", variable=self.word_var)
        self.canvas_bg.create_window(340, 300 + y_offset, anchor='center', window=chk_word)

        chk_pdf = ttk.Checkbutton(self, text="Create PDF File", variable=self.pdf_var)
        self.canvas_bg.create_window(480, 300 + y_offset, anchor='center', window=chk_pdf)

        lbl_folder_name = ttk.Label(self, text="Results Folder Name:")
        self.canvas_bg.create_window(400, 360 + y_offset, anchor='center', window=lbl_folder_name)

        self.folder_name_entry = ttk.Entry(self)
        self.canvas_bg.create_window(400, 390 + y_offset, anchor='center', window=self.folder_name_entry)

        btn_generate_files = ttk.Button(self, text="Create Files", command=self.generate_files)
        self.canvas_bg.create_window(400, 450 + y_offset, anchor='center', window=btn_generate_files)

        btn_exit = ttk.Button(self, text="EXIT", style='Red.TButton', command=self.destroy)
        self.canvas_bg.create_window(400, 510 + y_offset, anchor='center', window=btn_exit)

    def browse_word_template(self):
        self.word_template_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if self.word_template_path:
            messagebox.showinfo("Selected File", f"Word Template: {self.word_template_path}")

    def browse_excel_file(self):
        self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.excel_file_path:
            messagebox.showinfo("Selected File", f"Excel File: {self.excel_file_path}")

    def generate_files(self):
        if not self.word_template_path or not self.excel_file_path:
            messagebox.showerror("Error", "Please select both a Word template and an Excel file.")
            return

        if not (self.word_var.get() or self.pdf_var.get()):
            messagebox.showerror("Error", "Please select at least one file type to create.")
            return

        results_dir = self.folder_name_entry.get().strip()
        if not results_dir:
            messagebox.showerror("Error", "Please enter a results folder name.")
            return

        if not os.path.exists(results_dir):
            os.makedirs(results_dir)

        df = pd.read_excel(self.excel_file_path)

        for _, row in df.iterrows():
            file_name = row.get('file name')
            if not file_name:
                continue

            file_path = os.path.join(results_dir, f"{file_name}.docx")
            shutil.copy(self.word_template_path, file_path)

            doc = Document(file_path)
            for para in doc.paragraphs:
                for run in para.runs:
                    for key, value in row.items():
                        if pd.notna(value):
                            placeholder = f'<<{key}>>'
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, str(value))

            doc.save(file_path)

            if self.pdf_var.get():
                pdf_path = os.path.splitext(file_path)[0] + '.pdf'
                convert(file_path, pdf_path)

        messagebox.showinfo("Success", "Files generated successfully")

if __name__ == "__main__":
    app = Replace()
    app.mainloop()
