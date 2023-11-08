import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tqdm import tqdm  # Import tqdm

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("TicketSage: Automated Alert System for Invalid Tickets")
        self.pack()
        self.create_widgets()
        # Center the window on the screen
        window_width = self.master.winfo_reqwidth()
        window_height = self.master.winfo_reqheight()
        position_right = int(self.master.winfo_screenwidth() / 2 - window_width / 2)
        position_down = int(self.master.winfo_screenheight() / 2 - window_height / 2)
        self.master.geometry(f"+{position_right}+{position_down}")


    def create_widgets(self):
       
        # Excel file selection
        file_frame = tk.Frame(self)
        file_frame.pack(fill="x", padx=10, pady=5)

        self.file_label = tk.Label(file_frame, text="Select an Excel file:")
        self.file_label.pack(side="left")

        self.file_button = tk.Button(file_frame, text="Browse...", command=self.select_file)
        self.file_button.pack(side="left", padx=5)

        # Recipients input
        to_frame = tk.Frame(self)
        to_frame.pack(fill="x", padx=10, pady=5)

        self.to_label = tk.Label(to_frame, text="To:")
        self.to_label.pack(side="left")

        self.to_entry = tk.Entry(to_frame)
        self.to_entry.pack(side="left", padx=5)

        self.fs_button = tk.Button(to_frame, text="EUC", command=lambda: self.to_entry.insert(tk.END, "EUC@example.com;"))
        self.fs_button.pack(side="left", padx=5)

        self.fs_button = tk.Button(to_frame, text="FS", command=lambda: self.to_entry.insert(tk.END, "EUC@example.com;"))
        self.fs_button.pack(side="left", padx=5)

        self.fs_button = tk.Button(to_frame, text="EPS", command=lambda: self.to_entry.insert(tk.END, "EPS@example.com;"))
        self.fs_button.pack(side="left", padx=5)

        self.fs_button = tk.Button(to_frame, text="Voice", command=lambda: self.to_entry.insert(tk.END, "Voice@example.com;"))
        self.fs_button.pack(side="left", padx=5)

        self.fs_button = tk.Button(to_frame, text="UAM", command=lambda: self.to_entry.insert(tk.END, "UAM@example.com;"))
        self.fs_button.pack(side="left", padx=5)

        # Subject and sheet selection
        subject_sheet_frame = tk.Frame(self)
        subject_sheet_frame.pack(fill="x", padx=10, pady=5)

        self.subject_label = tk.Label(subject_sheet_frame, text="Subject:")
        self.subject_label.pack(side="left")

        self.subject_entry = tk.Entry(subject_sheet_frame)
        self.subject_entry.pack(side="left", padx=5)

        self.sheet_label = tk.Label(subject_sheet_frame, text="Select a sheet:")
        self.sheet_label.pack(side="left", padx=10)

        self.sheet_variable = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(subject_sheet_frame, textvariable=self.sheet_variable)
        self.sheet_dropdown.pack(side="left", padx=5)

        # Columns input
        #columns = ['A', 'E', 'K']

        # Send and Quit buttons
        button_frame = tk.Frame(self)
        button_frame.pack(fill="x", padx=10, pady=5)

        self.send_button = tk.Button(button_frame, text="Send", command=self.send_email)
        self.send_button.pack(side="left", padx=5)

        self.quit_button = tk.Button(button_frame, text="Quit", fg="red", command=self.master.destroy)
        self.quit_button.pack(side="left", padx=5)

        # Confirmation label
        self.confirm_label = tk.Label(self)
        self.confirm_label.pack(side="top")



    def select_file(self):
        # Show a file dialog to let the user select an Excel file
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        if not file_path:  # Check if the user canceled the file dialog
            return

        try:
            # Verify if the selected file is a valid Excel file
            sheets = pd.read_excel(file_path, sheet_name=None)
            if not sheets:
                raise ValueError("No sheets found in the selected Excel file.")
        except Exception as e:
            # Handle any exception that occurs during file reading (e.g., invalid file format)
            error_message = f"Error while reading the Excel file: {str(e)}"
            tk.messagebox.showerror("File Error", error_message)
            return

        # Update the file label with the selected file path
        self.file_label["text"] = f"Selected file: {file_path}"
        self.file_path = file_path

        # Get the sheets in the selected Excel file and update the dropdown
        self.sheet_dropdown["values"] = list(sheets.keys())

    def send_email(self):
        # Read the selected Excel file into a pandas DataFrame
        sheet_name = self.sheet_variable.get()
        df = pd.read_excel(self.file_path, sheet_name)

        # Select the specified columns (A, E, K)
        columns = ['Number','Reason For Invalid','Valid/Invalid','Assignment group','Assigned to']
        df = df[columns]

        # Convert the DataFrame to an HTML table
        table_html = df.to_html(index=False, na_rep="")

        # Get the total number of emails to be sent
        total_emails = len(self.to_entry.get().split(';'))

        # Create an Outlook object and log in
        outlook = win32.Dispatch("Outlook.Application")

        # Use tqdm to create a progress bar for sending emails
        with tqdm(total=total_emails, desc="Sending Emails", unit="email") as progress_bar:
            for email in self.to_entry.get().split(';'):
                if email.strip():  # Check for empty emails (due to extra semicolon in the entry)
                    mail = outlook.CreateItem(0)
                    mail.Display()

                    # Set the email recipients, subject, and body
                    mail.To = email
                    mail.Subject = self.subject_entry.get()
                    mail.HTMLBody = f'<p>Hi Team</p><p>Please find below list of invalids which need to be updated asap:</p>{table_html}<p>Kindly confirm back once updated. Fail to update in next 2 hours will result in moving the state to InProgress.</p>'

                    # Send the email
                    mail.Send()

                    progress_bar.update(1)  # Update the progress bar

        # Update the confirmation label
        self.confirm_label["text"] = "Emails sent successfully."

# Create the GUI
root = tk.Tk()
app = Application(master=root)
app.mainloop()
