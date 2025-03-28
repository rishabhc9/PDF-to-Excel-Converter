import pdfplumber
import pandas as pd
import os
import re
import json
import requests
from bs4 import BeautifulSoup
from datetime import datetime as dt
from tkinter import filedialog, messagebox, Tk, Label, Entry, Button, Frame, ttk, Menu
from tkinter import scrolledtext

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

def scrape_regex_data(search_query):
    url = f"https://www.regexlib.com/Search.aspx?k={search_query}"
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    tables = soup.find_all('table', class_='searchResultsTable')

    data = []
    for table in tables:
        table_html = str(table)
        title_match = re.search(r'<tr class="title".*?<a href="REDetails\.aspx\?regexp_id=\d+">(.*?)</a>', table_html, re.DOTALL)
        expression_match = re.search(r'<div class="expressionDiv">(.*?)</div>', table_html, re.DOTALL)
        description_match = re.search(r'<tr class="description".*?<div class="overflowFixDiv">(.*?)</div>', table_html, re.DOTALL)
        matches_match = re.search(r'<tr class="matches".*?<div class="overflowFixDiv">(.*?)</div>', table_html, re.DOTALL)
        non_matches_match = re.search(r'<tr class="nonmatches".*?<div class="overflowFixDiv">(.*?)</div>', table_html, re.DOTALL)

        title = title_match.group(1).strip() if title_match else "N/A"
        expression = expression_match.group(1).strip() if expression_match else "N/A"
        description = description_match.group(1).strip() if description_match else "N/A"
        matches = re.sub(r'<.*?>', '', matches_match.group(1)).strip() if matches_match else "N/A"
        non_matches = re.sub(r'<.*?>', '', non_matches_match.group(1)).strip() if non_matches_match else "N/A"

        data.append([title, expression, description, matches, non_matches])  # Removed author

    return data

def display_regex_results(search_entry, results_tree):
    search_query = search_entry.get().strip()
    if not search_query:
        messagebox.showerror("Error", "Please enter a search query.")
        return

    data = scrape_regex_data(search_query)
    for row in results_tree.get_children():
        results_tree.delete(row)
    for item in data:
        results_tree.insert("", "end", values=item)

def load_db_regex(results_tree):
    try:
        df = pd.read_excel("regex_database.xlsx")
        for row in results_tree.get_children():
            results_tree.delete(row)
        for _, row in df.iterrows():
            # Keep the "Title" column empty
            results_tree.insert("", "end", values=("", row['Expression'], row['Description'], row['Matches']))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load db regex: {e}")

def search_db_regex(search_entry, results_tree):
    search_query = search_entry.get().strip()
    if not search_query:
        messagebox.showerror("Error", "Please enter a search query.")
        return

    try:
        df = pd.read_excel("regex_database.xlsx")
        filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False), axis=1).any(axis=1)]
        for row in results_tree.get_children():
            results_tree.delete(row)
        for _, row in filtered_df.iterrows():
            # Use the "Matches" column data instead of the "Title" column
            results_tree.insert("", "end", values=("",row['Expression'], row['Description'], row['Matches']))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to search db regex: {e}")
        
def copy_to_clipboard(results_tree):
    selected_item = results_tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a row to copy.")
        return

    # Get the selected row's data
    item_data = results_tree.item(selected_item, "values")
    if not item_data:
        return

    # Extract the regex expression (second column, index 1)
    regex_expression = item_data[1]  # Index 1 corresponds to the "Expression" column

    # Copy the regex expression to the clipboard
    root.clipboard_clear()
    root.clipboard_append(regex_expression)
    root.update()  # Required to finalize the clipboard update
    messagebox.showinfo("Copied", "Regex expression copied to clipboard.")

def setup_context_menu(results_tree):
    # Create a context menu
    context_menu = Menu(root, tearoff=0)
    context_menu.add_command(label="Copy Regex", command=lambda: copy_to_clipboard(results_tree))

    # Bind the context menu to the Treeview
    def show_context_menu(event):
        item = results_tree.identify_row(event.y)
        if item:
            results_tree.selection_set(item)
            context_menu.post(event.x_root, event.y_root)

    results_tree.bind("<Button-3>", show_context_menu)  # Right-click binding

def main():
    global root
    root = Tk()
    root.title("Invisible Grid Table PDF to Excel Converter")

    # Input Folder
    Label(root, text="Input Folder:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    input_entry = Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=10, pady=5)
    Button(root, text="Browse", command=lambda: browse_folder(input_entry)).grid(row=0, column=2, padx=10, pady=5, sticky="w")

    # Output Folder
    Label(root, text="Output Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    output_entry = Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=5)
    Button(root, text="Browse", command=lambda: browse_folder(output_entry)).grid(row=1, column=2, padx=10, pady=5, sticky="w")

    # Column Names
    Label(root, text="Column Names (comma-separated):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    columns_entry = Entry(root, width=50)
    columns_entry.grid(row=2, column=1, padx=10, pady=5)

    # Regex Pattern
    Label(root, text="Regex Pattern:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
    regex_entry = Entry(root, width=50)
    regex_entry.grid(row=3, column=1, padx=10, pady=5)

    # Search Bar for Regex Query
    Label(root, text="Search Regex Query from Web (regexlib):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
    search_entry = Entry(root, width=50)
    search_entry.grid(row=4, column=1, padx=10, pady=5)
    Button(root, text="Find", command=lambda: display_regex_results(search_entry, results_tree)).grid(row=4, column=2, padx=10, pady=5, sticky="w")

    # Table to Display Results
    columns = ("Title", "Expression", "Description", "Matches", "Non-Matches")  # Removed "Author"
    results_tree = ttk.Treeview(root, columns=columns, show="headings")
    for col in columns:
        results_tree.heading(col, text=col)
    results_tree.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

    # Set up the context menu for copying regex expressions
    setup_context_menu(results_tree)

    # db Regex Patterns
    # Use a Frame to pack the elements tightly
    db_frame = Frame(root)
    db_frame.grid(row=6, column=0, columnspan=3, sticky="w", padx=10, pady=5)

    Label(db_frame, text="Search Regex from Custom DB:").pack(side="left", padx=(0, 5))
    db_search_entry = Entry(db_frame, width=30)
    db_search_entry.pack(side="left", padx=(0, 5))
    Button(db_frame, text="Search", command=lambda: search_db_regex(db_search_entry, results_tree)).pack(side="left", padx=(0, 5))
    Button(db_frame, text="Show All Regex from Custom DB", command=lambda: load_db_regex(results_tree)).pack(side="left", padx=(0, 10))

    # Button Frame
    button_frame = Frame(root)
    button_frame.grid(row=7, column=0, columnspan=3, pady=20)
    
    Button(button_frame, text="Save Config", command=lambda: save_config(input_entry.get(), output_entry.get(), columns_entry.get().split(','), regex_entry.get())).grid(row=0, column=0, padx=10)
    Button(button_frame, text="Load Config", command=lambda: load_config(input_entry, output_entry, columns_entry, regex_entry)).grid(row=0, column=1, padx=10)
    Button(button_frame, text="Convert", command=lambda: start_conversion(input_entry, output_entry, columns_entry, regex_entry)).grid(row=0, column=2, padx=10)

    root.mainloop()

if __name__ == "__main__":
    main()