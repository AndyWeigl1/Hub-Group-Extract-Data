import win32com.client
import win32gui
import pythoncom
import os
import ctypes
import tkinter as tk
from tkinter import messagebox
import PyPDF2
import re
from datetime import datetime
from io import BytesIO
import shutil

def check_outlook_open():
    hwnd = win32gui.FindWindow(None, "Microsoft Outlook")
    return hwnd != 0

def extract_text_from_pdf(pdf_file_bytes):
    pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_file_bytes))
    text = ''
    for page_num in range(len(pdf_reader.pages)):
        text += pdf_reader.pages[page_num].extract_text()
    return text

def find_due_date(text):
    pattern = r'INVOICE DATE\s+(\d{2}/\d{2}/\d{4})'
    match = re.search(pattern, text)

    if match:
        return match.group(1)
    else:
        return None

def save_attachments():
    if not check_outlook_open():
        error_message = "Outlook is not open. Please open Outlook and try again."
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", error_message, parent=root)
        return

    outlook = win32com.client.Dispatch("Outlook.Application")

    try:
        explorer = outlook.ActiveExplorer()
        selection = explorer.Selection
    except AttributeError:
        error_message = "Unable to retrieve selected emails."
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", error_message, parent=root)
        return

    if selection.Count == 0:
        error_message = "No emails selected."
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Info", error_message, parent=root)
        return

    base_save_folder = r"C:\Users\Andy Weigl\Kodiak Cakes\Kodiak Cakes Team Site - Public\Vendors\Hub Group\Bills"  # Update with your desired folder path

    saved_files = []

    additional_save_folder = r"C:\Users\Andy Weigl\OneDrive - Kodiak Cakes\Hub Group\Invoices"

    for email in selection:
        attachments = email.Attachments
        if attachments.Count == 0:
            print(f"No attachments found in email: {email.Subject}")
            continue

        for attachment in attachments:
            # Save the attachment to a temporary file
            temp_file_path = os.path.join(base_save_folder, attachment.FileName)
            attachment.SaveAsFile(temp_file_path)

            # Read the temporary file as bytes
            with open(temp_file_path, 'rb') as file:
                pdf_file_bytes = file.read()

            # Extract the date from the PDF
            pdf_text = extract_text_from_pdf(pdf_file_bytes)
            due_date = find_due_date(pdf_text)

            if due_date:
                due_date = datetime.strptime(due_date, "%m/%d/%Y")
                year_folder = os.path.join(base_save_folder, str(due_date.year))
                month_folder = os.path.join(year_folder, f"{due_date.strftime('%m')}.{due_date.year}")

                # Create year and month folders if they don't exist
                if not os.path.exists(year_folder):
                    os.makedirs(year_folder)
                if not os.path.exists(month_folder):
                    os.makedirs(month_folder)

                # Move the temporary file to the correct folder
                final_file_path = os.path.join(month_folder, attachment.FileName)
                shutil.move(temp_file_path, final_file_path)
                print(f"Attachment saved from email '{email.Subject}': {final_file_path}")
                saved_files.append(final_file_path)

                # Save the file to the additional folder
                additional_file_path = os.path.join(additional_save_folder, attachment.FileName)
                shutil.copy(final_file_path, additional_file_path)
                print(f"Attachment saved to additional folder from email '{email.Subject}': {additional_file_path}")
            else:
                print(f"Due date not found in email '{email.Subject}', saved to temporary folder: {temp_file_path}")
                saved_files.append(temp_file_path)

    # Call the display_summary function to show pop-ups for each month
    display_summary(saved_files)

def display_summary(saved_files):
    summary = {}
    for file_path in saved_files:
        folder, file_name = os.path.split(file_path)
        month_year = os.path.basename(folder)

        invoice_number = file_name.split('.')[0]  # Assuming invoice number is the part before the file extension

        if month_year not in summary:
            summary[month_year] = []

        summary[month_year].append(invoice_number)

    for month_year, invoice_numbers in summary.items():
        message = f"Invoices saved in {month_year}:\n\n" + "\n".join(invoice_numbers)
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Attachments Saved", message, parent=root)

pythoncom.CoInitialize()
save_attachments()
pythoncom.CoUninitialize()