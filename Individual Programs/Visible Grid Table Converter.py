import pdfplumber
import pandas as pd
import os
import re
import json
import requests
from bs4 import BeautifulSoup
from datetime import datetime as dt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Menu


# Logging function
def log_message(message, log_file="logfile.txt"):
    with open(log_file, 'a') as file:
        file.write(f"{dt.now()} - {message}\n")


# PDF processing functions
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


# Regex-related functions
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

        data.append([title, expression, description, matches, non_matches])

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
            results_tree.insert("", "end", values=("", row['Expression'], row['Description'], row['Matches']))
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


# Configuration functions
def load_config(input_entry, output_entry, columns_entry, regex_entry, index_entry):
    file_path = filedialog.askopenfilename(title="Select Configuration File", filetypes=[("JSON Files", "*.json")])
    if file_path:
        with open(file_path, 'r') as file:
            config = json.load(file)
        
        # Update the entry fields with loaded configuration
        input_entry.delete(0, tk.END)
        input_entry.insert(0, config.get('input_folder', ''))

        output_entry.delete(0, tk.END)
        output_entry.insert(0, config.get('output_folder', ''))

        columns_entry.delete(0, tk.END)
        columns_entry.insert(0, ', '.join(config.get('column_names', [])))

        regex_entry.delete(0, tk.END)
        regex_entry.insert(0, config.get('regex_pattern', ''))

        index_entry.delete(0, tk.END)
        index_entry.insert(0, str(config.get('filter_index', 0)))


def save_config(input_folder, output_folder, column_names, regex_pattern, filter_index):
    file_path = filedialog.asksaveasfilename(title="Save Configuration", defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if file_path:
        config = {
            'input_folder': input_folder,
            'output_folder': output_folder,
            'column_names': column_names,
            'regex_pattern': regex_pattern,
            'filter_index': filter_index
        }
        with open(file_path, 'w') as file:
            json.dump(config, file)
        return file_path
    return None


# Start conversion function
def start_conversion(input_entry, output_entry, columns_entry, regex_entry, index_entry, config_file_path):
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

    # Saving the configuration for future use
    if config_file_path:
        save_config(input_folder, output_folder, column_names, regex_pattern, filter_index)

    try:
        convert_pdfs_to_excel(input_folder, output_folder, column_names, regex_pattern, filter_index)
        messagebox.showinfo("Success", "PDFs successfully converted to Excel.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def browse_input_folder(entry):
    folder_selected = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_selected)

def browse_output_folder(entry):
    folder_selected = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_selected)
    
def main():
    global root
    root = tk.Tk()
    root.title("Grid-Based Table PDF to Excel Converter")

    # Input Folder
    tk.Label(root, text="Input Folder:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    input_entry = tk.Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_input_folder(input_entry)).grid(row=0, column=2, padx=10, pady=5, sticky="w")

    # Output Folder
    tk.Label(root, text="Output Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_output_folder(output_entry)).grid(row=1, column=2, padx=10, pady=5, sticky="w")

    # Column Names
    tk.Label(root, text="Column Names (comma-separated):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    columns_entry = tk.Entry(root, width=50)
    columns_entry.grid(row=2, column=1, padx=10, pady=5)

    # Regex Pattern
    tk.Label(root, text="Regex Pattern:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
    regex_entry = tk.Entry(root, width=50)
    regex_entry.grid(row=3, column=1, padx=10, pady=5)

    # Filter Index
    tk.Label(root, text="Index for Regex (0-based):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
    index_entry = tk.Entry(root, width=50)
    index_entry.grid(row=4, column=1, padx=10, pady=5)

    # Search Bar for Regex Query
    tk.Label(root, text="Search Regex Query:").grid(row=5, column=0, sticky="w", padx=10, pady=5)
    search_entry = tk.Entry(root, width=50)
    search_entry.grid(row=5, column=1, padx=10, pady=5)
    tk.Button(root, text="Find", command=lambda: display_regex_results(search_entry, results_tree)).grid(row=5, column=2, padx=10, pady=5, sticky="w")

    # Table to Display Results
    columns = ("Title", "Expression", "Description", "Matches", "Non-Matches")
    results_tree = ttk.Treeview(root, columns=columns, show="headings")
    for col in columns:
        results_tree.heading(col, text=col)
    results_tree.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

    # Set up the context menu for copying regex expressions
    setup_context_menu(results_tree)

    # db Regex Patterns
    # Use a Frame to pack the elements tightly
    db_frame = tk.Frame(root)
    db_frame.grid(row=7, column=0, columnspan=3, sticky="w", padx=10, pady=5)

    tk.Label(db_frame, text="Search Regex from Custom DB:").pack(side="left", padx=(0, 5))
    db_search_entry = tk.Entry(db_frame, width=30)
    db_search_entry.pack(side="left", padx=(0, 5))
    tk.Button(db_frame, text="Search", command=lambda: search_db_regex(db_search_entry, results_tree)).pack(side="left", padx=(0, 5))
    tk.Button(db_frame, text="Show All Regex from Custom DB", command=lambda: load_db_regex(results_tree)).pack(side="left", padx=(0, 10))

    # Button Frame
    button_frame = tk.Frame(root)
    button_frame.grid(row=8, column=0, columnspan=3, pady=20)
    
    tk.Button(button_frame, text="Save Config", command=lambda: save_config(input_entry.get(), output_entry.get(), columns_entry.get().split(','), regex_entry.get(), int(index_entry.get()))).grid(row=0, column=0, padx=10)
    tk.Button(button_frame, text="Load Config", command=lambda: load_config(input_entry, output_entry, columns_entry, regex_entry, index_entry)).grid(row=0, column=1, padx=10)
    tk.Button(button_frame, text="Convert", command=lambda: start_conversion(input_entry, output_entry, columns_entry, regex_entry, index_entry, '')).grid(row=0, column=2, padx=10)

    root.mainloop()


if __name__ == "__main__":
    main()