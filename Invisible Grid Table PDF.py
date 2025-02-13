import pdfplumber
import pandas as pd
import os
import re
import json
from datetime import datetime as dt
from tkinter import filedialog, messagebox, Tk, Label, Entry, Button, Frame

def log_message(message, log_file="logfile.txt"):
    with open(log_file, 'a') as file:
        file.write(f"{dt.now()} - {message}\n")

def extract_text_from_pdf(pdf_path):
    extracted_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                extracted_text.append(text)
    return " ".join(extracted_text)

def process_text_data(text, regex_pattern):
    matches = re.findall(regex_pattern, text, re.MULTILINE)
    return matches

def save_to_excel(data, column_names, output_file):
    df = pd.DataFrame(data, columns=column_names)
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)

def convert_pdfs_to_excel(input_folder, output_folder, column_names, regex_pattern):
    pdf_files = [file for file in os.listdir(input_folder) if file.endswith('.pdf')]
    if not pdf_files:
        log_message("No PDF files found in the input folder.")
        return

    log_message("Processing started.")

    for pdf_file in pdf_files:
        try:
            pdf_path = os.path.join(input_folder, pdf_file)
            pdf_name = os.path.splitext(pdf_file)[0]
            extracted_text = extract_text_from_pdf(pdf_path)
            extracted_data = process_text_data(extracted_text, regex_pattern)

            if extracted_data:
                output_file = os.path.join(output_folder, f"{pdf_name}.xlsx")
                save_to_excel(extracted_data, column_names, output_file)
                log_message(f"Successfully processed {pdf_file} and saved to {output_file}.")
            else:
                log_message(f"No matches found in {pdf_file}.")

        except Exception as e:
            log_message(f"Failed to process {pdf_file}. Error: {str(e)}")
            continue

def browse_folder(entry):
    folder_selected = filedialog.askdirectory()
    entry.delete(0, "end")
    entry.insert(0, folder_selected)

def save_config(input_folder, output_folder, column_names, regex_pattern):
    file_path = filedialog.asksaveasfilename(title="Save Configuration", defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if file_path:
        config = {
            'input_folder': input_folder,
            'output_folder': output_folder,
            'column_names': column_names,
            'regex_pattern': regex_pattern
        }
        with open(file_path, 'w') as file:
            json.dump(config, file)

def load_config(input_entry, output_entry, columns_entry, regex_entry):
    file_path = filedialog.askopenfilename(title="Select Configuration File", filetypes=[("JSON Files", "*.json")])
    if file_path:
        with open(file_path, 'r') as file:
            config = json.load(file)
        
        input_entry.delete(0, "end")
        input_entry.insert(0, config.get('input_folder', ''))

        output_entry.delete(0, "end")
        output_entry.insert(0, config.get('output_folder', ''))

        columns_entry.delete(0, "end")
        columns_entry.insert(0, ', '.join(config.get('column_names', [])))

        regex_entry.delete(0, "end")
        regex_entry.insert(0, config.get('regex_pattern', ''))

def start_conversion(input_entry, output_entry, columns_entry, regex_entry):
    input_folder = input_entry.get()
    output_folder = output_entry.get()
    column_names = [col.strip() for col in columns_entry.get().split(',')]
    regex_pattern = regex_entry.get().strip()

    if not input_folder or not output_folder or not column_names or not regex_pattern:
        messagebox.showerror("Error", "All fields are required.")
        return

    try:
        convert_pdfs_to_excel(input_folder, output_folder, column_names, regex_pattern)
        messagebox.showinfo("Success", "PDFs successfully converted to Excel.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def main():
    root = Tk()
    root.title("Invisible Grid Table PDF")

    Label(root, text="Input Folder:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    input_entry = Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=10, pady=5)
    Button(root, text="Browse", command=lambda: browse_folder(input_entry)).grid(row=0, column=2, padx=10, pady=5)

    Label(root, text="Output Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    output_entry = Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=5)
    Button(root, text="Browse", command=lambda: browse_folder(output_entry)).grid(row=1, column=2, padx=10, pady=5)

    Label(root, text="Column Names (comma-separated):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    columns_entry = Entry(root, width=50)
    columns_entry.grid(row=2, column=1, padx=10, pady=5)

    Label(root, text="Regex Pattern:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
    regex_entry = Entry(root, width=50)
    regex_entry.grid(row=3, column=1, padx=10, pady=5)

    button_frame = Frame(root)
    button_frame.grid(row=4, column=0, columnspan=3, pady=20)
    
    Button(button_frame, text="Save Config", command=lambda: save_config(input_entry.get(), output_entry.get(), columns_entry.get().split(','), regex_entry.get())).grid(row=0, column=0, padx=10)
    Button(button_frame, text="Load Config", command=lambda: load_config(input_entry, output_entry, columns_entry, regex_entry)).grid(row=0, column=1, padx=10)
    Button(button_frame, text="Convert", command=lambda: start_conversion(input_entry, output_entry, columns_entry, regex_entry)).grid(row=0, column=2, padx=10)

    root.mainloop()

if __name__ == "__main__":
    main()