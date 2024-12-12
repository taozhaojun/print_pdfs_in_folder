import os
import win32print
import win32api
import tkinter as tk
from tkinter import filedialog, messagebox


def print_pdfs_in_folder(folder_path):
    """
    Prints all PDF files in the specified folder.

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

        # Print each PDF file
        for pdf in pdf_files:
            pdf_path = os.path.join(folder_path, pdf)
            print(f"Printing: {pdf}")

            # Use win32api to send the file to the default printer
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