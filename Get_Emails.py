import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import base64
import smtplib

def credentials_exist(email, app_password, provider):
    try:
        filename = "credentials.txt"
        with open(filename, 'r') as file:
            lines = file.readlines()
            for i in range(0, len(lines), 4):  # Check every block of credentials
                if (base64.b64encode(email.encode()).decode('utf-8') == lines[i].strip() and
                        base64.b64encode(app_password.encode()).decode('utf-8') == lines[i + 1].strip() and
                        base64.b64encode(provider.encode()).decode('utf-8') == lines[i + 2].strip()):
                    return True
        return False
    except FileNotFoundError:
        return False

def save_credentials(email, app_password, provider):
    try:
        if credentials_exist(email, app_password, provider):
            messagebox.showinfo("Info", "Credentials already exist in the file.")
            return

        encoded_email = base64.b64encode(email.encode()).decode('utf-8')
        encoded_app_password = base64.b64encode(app_password.encode()).decode('utf-8')
        encoded_provider = base64.b64encode(provider.encode()).decode('utf-8')

        filename = "credentials.txt"
        with open(filename, 'a') as file:
            file.write(encoded_email + '\n')
            file.write(encoded_app_password + '\n')
            file.write(encoded_provider + '\n')
            file.write('\n')

        messagebox.showinfo("Info", "Encoded credentials saved successfully.")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to save credentials. Error: {e}")

def validate_smtp_credentials(email, app_password, provider):
    try:
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
            messagebox.showerror("Error", "Invalid provider.")
            return False

        if provider == "Outlook":
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
        else:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)

        server.login(email, app_password)
        server.quit()

        return True

    except Exception as e:
        messagebox.showerror("Error", f"Failed to authenticate using SMTP. Error: {e}")
        return False

def submit_credentials():
    email = email_entry.get()
    app_password = app_password_entry.get()
    provider = provider_var.get()

    if validate_smtp_credentials(email, app_password, provider):
        save_credentials(email, app_password, provider)
    else:
        messagebox.showerror("Error", "Failed to validate credentials.")

def toggle_password_visibility():
    if app_password_entry.cget('show') == '*':
        app_password_entry.config(show='')
        toggle_button.config(image=eye_open_photo)
    else:
        app_password_entry.config(show='*')
        toggle_button.config(image=eye_closed_photo)

def create_gui():
    root = tk.Tk()
    root.title("SMTP Credentials Validator and Saver")
    root.geometry("550x600")
    root.configure(bg="#282828")

    style = ttk.Style()
    style.theme_use("clam")

    style.configure('Modern.TFrame', background='#282828')
    style.configure('Modern.TLabel', background='#282828', foreground='#f0f0f0', font=('Arial', 12))
    style.configure('Modern.TButton', background='#4CAF50', foreground='white', font=('Arial', 12))
    style.configure('Modern.TCombobox', width=27, font=('Arial', 12))

    # Load and display logo, resizing it
    logo_image = Image.open("una-bot.jpg")  # Replace with your logo file path
    logo_image = logo_image.resize((200, 200), Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = tk.Label(root, image=logo_photo, bg="#282828")
    logo_label.image = logo_photo
    logo_label.pack(pady=10)

    frame = ttk.Frame(root, padding="20", style='Modern.TFrame')
    frame.pack(fill=tk.BOTH, expand=True)

    email_label = ttk.Label(frame, text="Email:", style='Modern.TLabel')
    email_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
    global email_entry
    email_entry = ttk.Entry(frame, width=30, font=('Arial', 12))
    email_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

    app_password_label = ttk.Label(frame, text="App Password:", style='Modern.TLabel')
    app_password_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
    global app_password_entry
    app_password_entry = ttk.Entry(frame, show="*", width=30, font=('Arial', 12))
    app_password_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

    # Load eye icons
    global eye_open_photo, eye_closed_photo
    eye_open_image = Image.open("eye-open.png").resize((20, 20), Image.LANCZOS)  # Replace with your eye-open icon path
    eye_open_photo = ImageTk.PhotoImage(eye_open_image)
    eye_closed_image = Image.open("eye-closed.png").resize((20, 20), Image.LANCZOS)  # Replace with your eye-closed icon path
    eye_closed_photo = ImageTk.PhotoImage(eye_closed_image)

    global toggle_button
    toggle_button = tk.Button(frame, image=eye_closed_photo, command=toggle_password_visibility, bg="#282828", bd=0)
    toggle_button.grid(row=1, column=2, padx=5, pady=5)

    provider_label = ttk.Label(frame, text="Provider:", style='Modern.TLabel')
    provider_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
    global provider_var
    provider_var = tk.StringVar()
    provider_combo = ttk.Combobox(frame, textvariable=provider_var, values=["Gmail", "Outlook", "Webmail"],
                                  width=27, style='Modern.TCombobox', font=('Arial', 12))
    provider_combo.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
    provider_combo.current(0)

    submit_btn = ttk.Button(frame, text="Submit", command=submit_credentials, style='Modern.TButton')
    submit_btn.grid(row=3, column=0, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
