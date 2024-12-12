import os
import win32print
import win32api
import tkinter as tk
from tkinter import filedialog, messagebox
import win32con
import copy

def print_pdfs_in_folder(folder_path):
    """
    Prints all PDF files in the specified folder in grayscale and landscape orientation.

    :param folder_path: The path of the folder containing PDF files.
    """
    try:
        # Get all files in the folder
        files = os.listdir(folder_path)
        pdf_files = [file for file in files if file.lower().endswith('.pdf')]

        # Check if there are PDF files in the folder
        if not pdf_files:
            messagebox.showinfo("No PDFs", "No PDF files found in the selected folder.")
            return

        # Get the default printer
        printer_name = win32print.GetDefaultPrinter()
        printer_handle = win32print.OpenPrinter(printer_name)

        # Get current printer settings
        printer_info = win32print.GetPrinter(printer_handle, 2)
        original_devmode = copy.deepcopy(printer_info['pDevMode'])

        # Modify DEVMODE for grayscale and landscape
        devmode = printer_info['pDevMode']
        if devmode:
            # Set to grayscale
            devmode.Color = win32con.DMCOLOR_MONOCHROME

            # Set orientation to landscape
            devmode.Orientation = win32con.DMORIENT_LANDSCAPE

            # Update the printer settings
            printer_info['pDevMode'] = devmode
            win32print.SetPrinter(printer_handle, 2, printer_info, 0)
        else:
            messagebox.showerror("Error", "Failed to retrieve printer settings.")
            return

        # Print each PDF file
        for pdf in pdf_files:
            pdf_path = os.path.join(folder_path, pdf)
            print(f"Printing: {pdf}")

            # Use win32api to send the file to the default printer
            # The "print" verb sends the file to the printer without showing a dialog
            win32api.ShellExecute(
                0,
                "print",
                pdf_path,
                None,
                ".",
                0
            )

        messagebox.showinfo("Success", "All PDFs have been sent to the printer.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

    finally:
        try:
            # Restore the original printer settings
            if 'original_devmode' in locals() and original_devmode:
                printer_info['pDevMode'] = original_devmode
                win32print.SetPrinter(printer_handle, 2, printer_info, 0)
            win32print.ClosePrinter(printer_handle)
        except Exception as e:
            messagebox.showwarning("Warning", f"Failed to restore printer settings: {e}")

def select_folder():
    """
    Opens a folder selection dialog and starts the printing process.
    """
    folder_path = filedialog.askdirectory(title="Select Folder Containing PDFs")
    if folder_path:
        print_pdfs_in_folder(folder_path)

# Create the main window
root = tk.Tk()
root.title("PDF Printer")
root.geometry("300x150")

# Create a button to select the folder
select_button = tk.Button(root, text="Select Folder", command=select_folder, width=20, height=2)
select_button.pack(pady=20)

# Run the tkinter event loop
root.mainloop()