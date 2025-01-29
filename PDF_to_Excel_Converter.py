import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime as dt
import tkinter as tk
from tkinter import filedialog, messagebox


def log_message(message, log_file="logfile.txt"):
    with open(log_file, 'a') as file:
        file.write(f"{dt.now()} - {message}\n")

def extract_information(pdf_path):
    pdf_obj = pdfplumber.open(pdf_path)
    return len(pdf_obj.pages), pdf_obj

def process_pdf(pdf_obj, page_count, column_names, regex_pattern, filter_index):
    extracted_data = []
    for i in range(page_count):
        page = pdf_obj.pages[i]
        table_data = page.extract_table()
        if table_data:
            extracted_data.extend(table_data)
    
    filtered_data = []
    for row in extracted_data:
        if row and len(row) > filter_index and re.match(regex_pattern, str(row[filter_index])):
            filtered_data.append(row)

    # Creating dynamic variables for columns
    column_data = {col: [] for col in column_names}
    for row in filtered_data:
        for idx, col in enumerate(column_names):
            column_data[col].append(row[idx] if idx < len(row) else None)

    return column_data

def save_to_excel(column_data, output_file):
    df = pd.DataFrame(column_data)
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)

def convert_pdfs_to_excel(input_folder, output_folder, column_names, regex_pattern, filter_index):
    pdf_files = [file for file in os.listdir(input_folder) if file.endswith('.pdf')]

    if not pdf_files:
        log_message("No PDF files found in the input folder.")
        return

    log_message("Processing started.")

    for pdf_file in pdf_files:
        try:
            pdf_path = os.path.join(input_folder, pdf_file)
            page_count, pdf_obj = extract_information(pdf_path)

            log_message(f"Processing {pdf_file} with {page_count} pages.")

            column_data = process_pdf(pdf_obj, page_count, column_names, regex_pattern, filter_index)

            output_file = os.path.join(output_folder, f"{os.path.splitext(pdf_file)[0]}.xlsx")
            save_to_excel(column_data, output_file)

            log_message(f"Successfully processed {pdf_file} and saved to {output_file}.")

            pdf_obj.close()

        except Exception as e:
            log_message(f"Failed to process {pdf_file}. Error: {str(e)}")
            continue

def browse_input_folder(entry):
    folder_selected = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_selected)

def browse_output_folder(entry):
    folder_selected = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_selected)

def start_conversion(input_entry, output_entry, columns_entry, regex_entry, index_entry):
    input_folder = input_entry.get()
    output_folder = output_entry.get()
    column_names = [col.strip() for col in columns_entry.get().split(',')]
    regex_pattern = regex_entry.get()
    try:
        filter_index = int(index_entry.get())
    except ValueError:
        messagebox.showerror("Error", "Filter index must be an integer.")
        return

    if not input_folder or not output_folder or not column_names or not regex_pattern:
        messagebox.showerror("Error", "All fields are required.")
        return

    try:
        convert_pdfs_to_excel(input_folder, output_folder, column_names, regex_pattern, filter_index)
        messagebox.showinfo("Success", "PDFs successfully converted to Excel.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def main():
    root = tk.Tk()
    root.title("PDF to Excel Converter")

    # Input Folder
    tk.Label(root, text="Input Folder:", anchor="w").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    input_entry = tk.Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_input_folder(input_entry)).grid(row=0, column=2, padx=10, pady=5)

    # Output Folder
    tk.Label(root, text="Output Folder:", anchor="w").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_output_folder(output_entry)).grid(row=1, column=2, padx=10, pady=5)

    # Column Names
    tk.Label(root, text="Column Names (Example - Name, Email):", anchor="w").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    columns_entry = tk.Entry(root, width=50)
    columns_entry.grid(row=2, column=1, padx=10, pady=5)

    # Regex for Filtering
    tk.Label(root, text="Regex for Identifying Table (Example - \w{3}-\w{3}):", anchor="w").grid(row=3, column=0, sticky="w", padx=10, pady=5)
    regex_entry = tk.Entry(root, width=50)
    regex_entry.grid(row=3, column=1, padx=10, pady=5)

    # Filter Index
    tk.Label(root, text="Index for Regex (0-based):", anchor="w").grid(row=4, column=0, sticky="w", padx=10, pady=5)
    index_entry = tk.Entry(root, width=50)
    index_entry.grid(row=4, column=1, padx=10, pady=5)

    # Convert Button
    tk.Button(root, text="Convert", command=lambda: start_conversion(input_entry, output_entry, columns_entry, regex_entry, index_entry)).grid(row=5, column=1, pady=20)
    root.mainloop()

if __name__ == "__main__":
    main()
