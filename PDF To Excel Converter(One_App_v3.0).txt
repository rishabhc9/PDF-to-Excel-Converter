import pdfplumber
import pandas as pd
import os
import re
import json
import requests
from bs4 import BeautifulSoup
from datetime import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu, scrolledtext
import shutil
from datetime import datetime
from tkcalendar import Calendar

# ==================== COMMON FUNCTIONS ====================
def log_message(message, log_file="logfile.txt"):
    with open(log_file, 'a') as file:
        file.write(f"{dt.now()} - {message}\n")

def browse_folder(entry):
    folder_selected = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_selected)

def setup_context_menu(results_tree):
    context_menu = Menu(root, tearoff=0)
    context_menu.add_command(label="Copy Regex", command=lambda: copy_to_clipboard(results_tree))
    
    def show_context_menu(event):
        item = results_tree.identify_row(event.y)
        if item:
            results_tree.selection_set(item)
            context_menu.post(event.x_root, event.y_root)
    
    results_tree.bind("<Button-3>", show_context_menu)

def copy_to_clipboard(results_tree):
    selected_item = results_tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a row to copy.")
        return
    
    item_data = results_tree.item(selected_item, "values")
    if not item_data:
        return
    
    regex_expression = item_data[1]
    root.clipboard_clear()
    root.clipboard_append(regex_expression)
    root.update()
    messagebox.showinfo("Copied", "Regex expression copied to clipboard.")

# ==================== PROGRAM 1: INVISIBLE GRID CONVERTER ====================
class InvisibleGridConverter:
    def __init__(self, tab):
        self.tab = tab
        self.setup_ui()
    
    def setup_ui(self):
        # Input Folder
        ttk.Label(self.tab, text="Input Folder:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.input_entry = ttk.Entry(self.tab, width=50)
        self.input_entry.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(self.tab, text="Browse", command=lambda: browse_folder(self.input_entry)).grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # Output Folder
        ttk.Label(self.tab, text="Output Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.output_entry = ttk.Entry(self.tab, width=50)
        self.output_entry.grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(self.tab, text="Browse", command=lambda: browse_folder(self.output_entry)).grid(row=1, column=2, padx=10, pady=5, sticky="w")

        # Column Names
        ttk.Label(self.tab, text="Column Names (comma-separated):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.columns_entry = ttk.Entry(self.tab, width=50)
        self.columns_entry.grid(row=2, column=1, padx=10, pady=5)

        # Regex Pattern
        ttk.Label(self.tab, text="Regex Pattern:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.regex_entry = ttk.Entry(self.tab, width=50)
        self.regex_entry.grid(row=3, column=1, padx=10, pady=5)

        # Search Bar for Regex Query
        ttk.Label(self.tab, text="Search Regex Query from Web (regexlib):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.search_entry = ttk.Entry(self.tab, width=50)
        self.search_entry.grid(row=4, column=1, padx=10, pady=5)
        ttk.Button(self.tab, text="Find", command=self.display_regex_results).grid(row=4, column=2, padx=10, pady=5, sticky="w")

        # Table to Display Results
        columns = ("Title", "Expression", "Description", "Matches", "Non-Matches")
        self.results_tree = ttk.Treeview(self.tab, columns=columns, show="headings")
        for col in columns:
            self.results_tree.heading(col, text=col)
        self.results_tree.grid(row=5, column=0, columnspan=3, padx=10, pady=10)
        setup_context_menu(self.results_tree)

        # db Regex Patterns
        db_frame = ttk.Frame(self.tab)
        db_frame.grid(row=6, column=0, columnspan=3, sticky="w", padx=10, pady=5)

        ttk.Label(db_frame, text="Search Regex from Custom DB:").pack(side="left", padx=(0, 5))
        self.db_search_entry = ttk.Entry(db_frame, width=30)
        self.db_search_entry.pack(side="left", padx=(0, 5))
        ttk.Button(db_frame, text="Search", command=self.search_db_regex).pack(side="left", padx=(0, 5))
        ttk.Button(db_frame, text="Show All", command=self.load_db_regex).pack(side="left", padx=(0, 10))

        # Button Frame
        button_frame = ttk.Frame(self.tab)
        button_frame.grid(row=7, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Save Config", command=self.save_config).grid(row=0, column=0, padx=10)
        ttk.Button(button_frame, text="Load Config", command=self.load_config).grid(row=0, column=1, padx=10)
        ttk.Button(button_frame, text="Convert", command=self.start_conversion).grid(row=0, column=2, padx=10)

    def extract_text_from_pdf(self, pdf_path):
        extracted_text = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted_text.append(text)
        return " ".join(extracted_text)

    def process_text_data(self, text, regex_pattern):
        return re.findall(regex_pattern, text, re.MULTILINE)

    def save_to_excel(self, data, column_names, output_file):
        df = pd.DataFrame(data, columns=column_names)
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        df.to_excel(output_file, index=False)

    def convert_pdfs_to_excel(self, input_folder, output_folder, column_names, regex_pattern):
        pdf_files = [file for file in os.listdir(input_folder) if file.endswith('.pdf')]
        if not pdf_files:
            log_message("No PDF files found in the input folder.")
            return

        log_message("Processing started.")

        for pdf_file in pdf_files:
            try:
                pdf_path = os.path.join(input_folder, pdf_file)
                pdf_name = os.path.splitext(pdf_file)[0]
                extracted_text = self.extract_text_from_pdf(pdf_path)
                extracted_data = self.process_text_data(extracted_text, regex_pattern)

                if extracted_data:
                    output_file = os.path.join(output_folder, f"{pdf_name}.xlsx")
                    self.save_to_excel(extracted_data, column_names, output_file)
                    log_message(f"Successfully processed {pdf_file} and saved to {output_file}.")
                else:
                    log_message(f"No matches found in {pdf_file}.")

            except Exception as e:
                log_message(f"Failed to process {pdf_file}. Error: {str(e)}")
                continue

    def scrape_regex_data(self, search_query):
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

    def display_regex_results(self):
        search_query = self.search_entry.get().strip()
        if not search_query:
            messagebox.showerror("Error", "Please enter a search query.")
            return

        data = self.scrape_regex_data(search_query)
        for row in self.results_tree.get_children():
            self.results_tree.delete(row)
        for item in data:
            self.results_tree.insert("", "end", values=item)

    def load_db_regex(self):
        try:
            df = pd.read_excel("regex_database.xlsx")
            for row in self.results_tree.get_children():
                self.results_tree.delete(row)
            for _, row in df.iterrows():
                self.results_tree.insert("", "end", values=("", row['Expression'], row['Description'], row['Matches']))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load db regex: {e}")

    def search_db_regex(self):
        search_query = self.db_search_entry.get().strip()
        if not search_query:
            messagebox.showerror("Error", "Please enter a search query.")
            return

        try:
            df = pd.read_excel("regex_database.xlsx")
            filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False), axis=1).any(axis=1)]
            for row in self.results_tree.get_children():
                self.results_tree.delete(row)
            for _, row in filtered_df.iterrows():
                self.results_tree.insert("", "end", values=("", row['Expression'], row['Description'], row['Matches']))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to search db regex: {e}")

    def save_config(self):
        input_folder = self.input_entry.get()
        output_folder = self.output_entry.get()
        column_names = [col.strip() for col in self.columns_entry.get().split(',')]
        regex_pattern = self.regex_entry.get().strip()

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

    def load_config(self):
        file_path = filedialog.askopenfilename(title="Select Configuration File", filetypes=[("JSON Files", "*.json")])
        if file_path:
            with open(file_path, 'r') as file:
                config = json.load(file)
            
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, config.get('input_folder', ''))

            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, config.get('output_folder', ''))

            self.columns_entry.delete(0, tk.END)
            self.columns_entry.insert(0, ', '.join(config.get('column_names', [])))

            self.regex_entry.delete(0, tk.END)
            self.regex_entry.insert(0, config.get('regex_pattern', ''))

    def start_conversion(self):
        input_folder = self.input_entry.get()
        output_folder = self.output_entry.get()
        column_names = [col.strip() for col in self.columns_entry.get().split(',')]
        regex_pattern = self.regex_entry.get().strip()

        if not input_folder or not output_folder or not column_names or not regex_pattern:
            messagebox.showerror("Error", "All fields are required.")
            return

        try:
            self.convert_pdfs_to_excel(input_folder, output_folder, column_names, regex_pattern)
            messagebox.showinfo("Success", "PDFs successfully converted to Excel.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# ==================== PROGRAM 2: GRID-BASED CONVERTER ====================
class GridBasedConverter:
    def __init__(self, tab):
        self.tab = tab
        self.setup_ui()
    
    def setup_ui(self):
        # Input Folder
        ttk.Label(self.tab, text="Input Folder:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.input_entry = ttk.Entry(self.tab, width=50)
        self.input_entry.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(self.tab, text="Browse", command=lambda: browse_folder(self.input_entry)).grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # Output Folder
        ttk.Label(self.tab, text="Output Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.output_entry = ttk.Entry(self.tab, width=50)
        self.output_entry.grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(self.tab, text="Browse", command=lambda: browse_folder(self.output_entry)).grid(row=1, column=2, padx=10, pady=5, sticky="w")

        # Column Names
        ttk.Label(self.tab, text="Column Names (comma-separated):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.columns_entry = ttk.Entry(self.tab, width=50)
        self.columns_entry.grid(row=2, column=1, padx=10, pady=5)

        # Regex Pattern
        ttk.Label(self.tab, text="Regex Pattern:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.regex_entry = ttk.Entry(self.tab, width=50)
        self.regex_entry.grid(row=3, column=1, padx=10, pady=5)

        # Filter Index
        ttk.Label(self.tab, text="Index for Regex (0-based):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.index_entry = ttk.Entry(self.tab, width=50)
        self.index_entry.grid(row=4, column=1, padx=10, pady=5)

        # Search Bar for Regex Query
        ttk.Label(self.tab, text="Search Regex Query from Web (regexlib):").grid(row=5, column=0, sticky="w", padx=10, pady=5)
        self.search_entry = ttk.Entry(self.tab, width=50)
        self.search_entry.grid(row=5, column=1, padx=10, pady=5)
        ttk.Button(self.tab, text="Find", command=self.display_regex_results).grid(row=5, column=2, padx=10, pady=5, sticky="w")

        # Table to Display Results
        columns = ("Title", "Expression", "Description", "Matches", "Non-Matches")
        self.results_tree = ttk.Treeview(self.tab, columns=columns, show="headings")
        for col in columns:
            self.results_tree.heading(col, text=col)
        self.results_tree.grid(row=6, column=0, columnspan=3, padx=10, pady=10)
        setup_context_menu(self.results_tree)

        # db Regex Patterns
        db_frame = ttk.Frame(self.tab)
        db_frame.grid(row=7, column=0, columnspan=3, sticky="w", padx=10, pady=5)

        ttk.Label(db_frame, text="Search Regex from Custom DB:").pack(side="left", padx=(0, 5))
        self.db_search_entry = ttk.Entry(db_frame, width=30)
        self.db_search_entry.pack(side="left", padx=(0, 5))
        ttk.Button(db_frame, text="Search", command=self.search_db_regex).pack(side="left", padx=(0, 5))
        ttk.Button(db_frame, text="Show All", command=self.load_db_regex).pack(side="left", padx=(0, 10))

        # Button Frame
        button_frame = ttk.Frame(self.tab)
        button_frame.grid(row=8, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Save Config", command=self.save_config).grid(row=0, column=0, padx=10)
        ttk.Button(button_frame, text="Load Config", command=self.load_config).grid(row=0, column=1, padx=10)
        ttk.Button(button_frame, text="Convert", command=self.start_conversion).grid(row=0, column=2, padx=10)

    def extract_information(self, pdf_path):
        pdf_obj = pdfplumber.open(pdf_path)
        return len(pdf_obj.pages), pdf_obj

    def process_pdf(self, pdf_obj, page_count, column_names, regex_pattern, filter_index):
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

        column_data = {col: [] for col in column_names}
        for row in filtered_data:
            for idx, col in enumerate(column_names):
                column_data[col].append(row[idx] if idx < len(row) else None)

        return column_data

    def save_to_excel(self, column_data, output_file):
        df = pd.DataFrame(column_data)
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        df.to_excel(output_file, index=False)

    def convert_pdfs_to_excel(self, input_folder, output_folder, column_names, regex_pattern, filter_index):
        pdf_files = [file for file in os.listdir(input_folder) if file.endswith('.pdf')]

        if not pdf_files:
            log_message("No PDF files found in the input folder.")
            return

        log_message("Processing started.")

        for pdf_file in pdf_files:
            try:
                pdf_path = os.path.join(input_folder, pdf_file)
                page_count, pdf_obj = self.extract_information(pdf_path)

                log_message(f"Processing {pdf_file} with {page_count} pages.")

                column_data = self.process_pdf(pdf_obj, page_count, column_names, regex_pattern, filter_index)

                output_file = os.path.join(output_folder, f"{os.path.splitext(pdf_file)[0]}.xlsx")
                self.save_to_excel(column_data, output_file)

                log_message(f"Successfully processed {pdf_file} and saved to {output_file}.")

                pdf_obj.close()

            except Exception as e:
                log_message(f"Failed to process {pdf_file}. Error: {str(e)}")
                continue

    def scrape_regex_data(self, search_query):
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

    def display_regex_results(self):
        search_query = self.search_entry.get().strip()
        if not search_query:
            messagebox.showerror("Error", "Please enter a search query.")
            return

        data = self.scrape_regex_data(search_query)
        for row in self.results_tree.get_children():
            self.results_tree.delete(row)
        for item in data:
            self.results_tree.insert("", "end", values=item)

    def load_db_regex(self):
        try:
            df = pd.read_excel("regex_database.xlsx")
            for row in self.results_tree.get_children():
                self.results_tree.delete(row)
            for _, row in df.iterrows():
                self.results_tree.insert("", "end", values=("", row['Expression'], row['Description'], row['Matches']))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load db regex: {e}")

    def search_db_regex(self):
        search_query = self.db_search_entry.get().strip()
        if not search_query:
            messagebox.showerror("Error", "Please enter a search query.")
            return

        try:
            df = pd.read_excel("regex_database.xlsx")
            filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False), axis=1).any(axis=1)]
            for row in self.results_tree.get_children():
                self.results_tree.delete(row)
            for _, row in filtered_df.iterrows():
                self.results_tree.insert("", "end", values=("", row['Expression'], row['Description'], row['Matches']))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to search db regex: {e}")

    def save_config(self):
        input_folder = self.input_entry.get()
        output_folder = self.output_entry.get()
        column_names = [col.strip() for col in self.columns_entry.get().split(',')]
        regex_pattern = self.regex_entry.get()
        try:
            filter_index = int(self.index_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Filter index must be an integer.")
            return

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

    def load_config(self):
        file_path = filedialog.askopenfilename(title="Select Configuration File", filetypes=[("JSON Files", "*.json")])
        if file_path:
            with open(file_path, 'r') as file:
                config = json.load(file)
            
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, config.get('input_folder', ''))

            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, config.get('output_folder', ''))

            self.columns_entry.delete(0, tk.END)
            self.columns_entry.insert(0, ', '.join(config.get('column_names', [])))

            self.regex_entry.delete(0, tk.END)
            self.regex_entry.insert(0, config.get('regex_pattern', ''))

            self.index_entry.delete(0, tk.END)
            self.index_entry.insert(0, str(config.get('filter_index', 0)))

    def start_conversion(self):
        input_folder = self.input_entry.get()
        output_folder = self.output_entry.get()
        column_names = [col.strip() for col in self.columns_entry.get().split(',')]
        regex_pattern = self.regex_entry.get()
        try:
            filter_index = int(self.index_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Filter index must be an integer.")
            return

        if not input_folder or not output_folder or not column_names or not regex_pattern:
            messagebox.showerror("Error", "All fields are required.")
            return

        try:
            self.convert_pdfs_to_excel(input_folder, output_folder, column_names, regex_pattern, filter_index)
            messagebox.showinfo("Success", "PDFs successfully converted to Excel.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# ==================== PROGRAM 3: FLATTEN FOLDER TOOL ====================
class FlattenFolderTool:
    def __init__(self, tab):
        self.tab = tab
        self.setup_ui()
    
    def setup_ui(self):
        # Configure grid weights for proper resizing
        self.tab.grid_columnconfigure(1, weight=1)
        self.tab.grid_rowconfigure(7, weight=1)
        
        # Input Folder
        ttk.Label(self.tab, text="Source Folder:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.input_entry = ttk.Entry(self.tab, width=50)
        self.input_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(self.tab, text="Browse", command=lambda: browse_folder(self.input_entry)).grid(row=0, column=2, padx=10, pady=5)
        
        # Output Folder
        ttk.Label(self.tab, text="Destination Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.output_entry = ttk.Entry(self.tab, width=50)
        self.output_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(self.tab, text="Browse", command=lambda: browse_folder(self.output_entry)).grid(row=1, column=2, padx=10, pady=5)
        
        # Operation Type (Move/Copy)
        ttk.Label(self.tab, text="Operation:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.operation_var = tk.StringVar(value="Move")
        operation_menu = ttk.Combobox(self.tab, textvariable=self.operation_var, values=["Move", "Copy"], state="readonly", width=47)
        operation_menu.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        
        # File Extensions
        ttk.Label(self.tab, text="File Extensions (comma-separated):").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.extensions_entry = ttk.Entry(self.tab, width=50)
        self.extensions_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        
        # All Extensions Checkbox
        self.all_extensions_var = tk.BooleanVar()
        all_extensions_cb = ttk.Checkbutton(self.tab, text="All Extensions", variable=self.all_extensions_var)
        all_extensions_cb.grid(row=3, column=2, padx=10, pady=5, sticky="w")
        
        # Duplicate Handling
        ttk.Label(self.tab, text="Handle Duplicates:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.duplicates_var = tk.StringVar(value="rename")
        duplicates_frame = ttk.Frame(self.tab)
        duplicates_frame.grid(row=4, column=1, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(duplicates_frame, text="Rename", variable=self.duplicates_var, value="rename").pack(side="left")
        ttk.Radiobutton(duplicates_frame, text="Overwrite", variable=self.duplicates_var, value="overwrite").pack(side="left", padx=10)
        ttk.Radiobutton(duplicates_frame, text="Skip", variable=self.duplicates_var, value="skip").pack(side="left")
        
        # Execute Button
        ttk.Button(self.tab, text="Extract Files", command=self.extract_files, width=20).grid(row=5, column=1, pady=10)

        # Log Area
        ttk.Label(self.tab, text="Operation Log:").grid(row=6, column=0, sticky="nw", padx=10, pady=5)
        
        # Create a frame for the log text with integrated scrollbar
        log_frame = ttk.Frame(self.tab)
        log_frame.grid(row=7, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
        # Create Text widget with integrated scrollbar
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky="nsew")
    
    def extract_files(self):
        source_folder = self.input_entry.get()
        destination_folder = self.output_entry.get()
        operation = self.operation_var.get()  # 'Move' or 'Copy'
        extensions = self.extensions_entry.get().strip()
        all_extensions = self.all_extensions_var.get()  # Boolean
        handle_duplicates = self.duplicates_var.get()  # 'keep', 'overwrite', or 'rename'
        
        # Validate inputs
        if not source_folder or not destination_folder:
            messagebox.showerror("Error", "Please select both source and destination folders")
            return
        
        if not os.path.exists(source_folder):
            messagebox.showerror("Error", "Source folder does not exist")
            return
        
        # Create destination folder if it doesn't exist
        os.makedirs(destination_folder, exist_ok=True)
        
        # Process extensions
        if all_extensions:
            extensions = None  # Process all extensions
        else:
            extensions = [ext.strip().lower() for ext in extensions.split(',') if ext.strip()]
            if not extensions:
                messagebox.showerror("Error", "Please specify file extensions or check 'All Extensions'")
                return
        
        # Clear the log
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, f"Starting {operation.lower()} operation...\n")
        self.log_text.insert(tk.END, f"Duplicate handling: {handle_duplicates}\n")
        self.log_text.see(tk.END)
        self.tab.update()
        
        try:
            processed_files = 0
            skipped_files = 0
            
            # Iterate through each secondary folder in the main folder
            for secondary_folder in os.listdir(source_folder):
                secondary_folder_path = os.path.join(source_folder, secondary_folder)
                
                # Check if the item is a directory
                if os.path.isdir(secondary_folder_path):
                    
                    # Iterate through the items within each secondary folder
                    for item in os.listdir(secondary_folder_path):
                        item_path = os.path.join(secondary_folder_path, item)
                        
                        # Check if the item is a directory or a file
                        if os.path.isdir(item_path):
                            # Iterate through the files in each subfolder
                            for file_name in os.listdir(item_path):
                                file_path = os.path.join(item_path, file_name)
                                
                                # Check extension if needed
                                if extensions is None or os.path.splitext(file_name)[1].lower() in extensions:
                                    dest_path = os.path.join(destination_folder, file_name)
                                    
                                    # Handle duplicates based on user selection
                                    if os.path.exists(dest_path):
                                        if handle_duplicates == 'skip':
                                            self.log_text.insert(tk.END, f"Skipped duplicate: {file_path}\n")
                                            skipped_files += 1
                                            continue
                                        elif handle_duplicates == 'overwrite':
                                            if operation == 'Move':
                                                os.remove(dest_path)  # Remove existing file before moving
                                            else:
                                                pass  # copy2 will overwrite by default
                                        elif handle_duplicates == 'rename':
                                            counter = 1
                                            name, ext = os.path.splitext(file_name)
                                            while os.path.exists(dest_path):
                                                dest_path = os.path.join(destination_folder, f"{name}_{counter}{ext}")
                                                counter += 1
                                    
                                    # Perform the operation
                                    try:
                                        if operation == 'Move':
                                            shutil.move(file_path, dest_path)
                                            self.log_text.insert(tk.END, f"Moved: {file_path} → {dest_path}\n")
                                        else:
                                            shutil.copy2(file_path, dest_path)
                                            self.log_text.insert(tk.END, f"Copied: {file_path} → {dest_path}\n")
                                        processed_files += 1
                                    except Exception as e:
                                        self.log_text.insert(tk.END, f"Error processing {file_path}: {str(e)}\n")
                                    
                                    self.log_text.see(tk.END)
                                    self.tab.update()
                        else:
                            # If it's a file, check extension if needed
                            if extensions is None or os.path.splitext(item)[1].lower() in extensions:
                                dest_path = os.path.join(destination_folder, item)
                                
                                # Handle duplicates based on user selection
                                if os.path.exists(dest_path):
                                    if handle_duplicates == 'skip':
                                        self.log_text.insert(tk.END, f"Skipped duplicate: {item_path}\n")
                                        skipped_files += 1
                                        continue
                                    elif handle_duplicates == 'overwrite':
                                        if operation == 'Move':
                                            os.remove(dest_path)  # Remove existing file before moving
                                        else:
                                            pass  # copy2 will overwrite by default
                                    elif handle_duplicates == 'rename':
                                        counter = 1
                                        name, ext = os.path.splitext(item)
                                        while os.path.exists(dest_path):
                                            dest_path = os.path.join(destination_folder, f"{name}_{counter}{ext}")
                                            counter += 1
                                
                                # Perform the operation
                                try:
                                    if operation == 'Move':
                                        shutil.move(item_path, dest_path)
                                        self.log_text.insert(tk.END, f"Moved: {item_path} → {dest_path}\n")
                                    else:
                                        shutil.copy2(item_path, dest_path)
                                        self.log_text.insert(tk.END, f"Copied: {item_path} → {dest_path}\n")
                                    processed_files += 1
                                except Exception as e:
                                    self.log_text.insert(tk.END, f"Error processing {item_path}: {str(e)}\n")
                                
                                self.log_text.see(tk.END)
                                self.tab.update()
            
            self.log_text.insert(tk.END, f"\nOperation completed!\n")
            self.log_text.insert(tk.END, f"Files processed: {processed_files}\n")
            self.log_text.insert(tk.END, f"Files skipped: {skipped_files}\n")
            messagebox.showinfo("Success", f"Operation completed!\nProcessed: {processed_files} files\nSkipped: {skipped_files} files")
        except Exception as e:
            self.log_text.insert(tk.END, f"\nError: {str(e)}\n")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            self.log_text.see(tk.END)

# ==================== PROGRAM 4: FILE ORGANIZER ====================
class FileOrganizerTool:
    def __init__(self, tab):
        self.tab = tab
        self.calendar_windows = []
        self.setup_ui()
    
    def setup_ui(self):
        # Configure grid weights
        self.tab.grid_columnconfigure(1, weight=1)
        self.tab.grid_rowconfigure(10, weight=1)
        
        # Notebook for different organization methods
        self.org_notebook = ttk.Notebook(self.tab)
        self.org_notebook.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=10, pady=5)

        # Create tabs
        self.extension_tab = ttk.Frame(self.org_notebook)
        self.size_tab = ttk.Frame(self.org_notebook)
        self.date_tab = ttk.Frame(self.org_notebook)
        self.name_tab = ttk.Frame(self.org_notebook)

        self.org_notebook.add(self.extension_tab, text="By Extension")
        self.org_notebook.add(self.size_tab, text="By Size")
        self.org_notebook.add(self.date_tab, text="By Date")
        self.org_notebook.add(self.name_tab, text="By Name")

        # Setup tab-specific UIs
        self.setup_extension_tab()
        self.setup_size_tab()
        self.setup_date_tab()
        self.setup_name_tab()
        
        # Common widgets for all tabs
        self.setup_common_widgets()
    
    def setup_common_widgets(self):
        # Input Folder
        ttk.Label(self.tab, text="Source Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.input_entry = ttk.Entry(self.tab, width=60)
        self.input_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(self.tab, text="Browse", command=lambda: self.browse_folder(self.input_entry)).grid(row=1, column=2, padx=10, pady=5)
        
        # Operation Type
        ttk.Label(self.tab, text="Operation:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.operation_var = tk.StringVar(value="Copy")
        ttk.Combobox(self.tab, textvariable=self.operation_var, 
                    values=["Copy", "Move"], state="readonly").grid(row=2, column=1, padx=10, pady=5, sticky="w")
        
        # Destination Folder
        ttk.Label(self.tab, text="Destination Folder:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.dest_entry = ttk.Entry(self.tab, width=60)
        self.dest_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(self.tab, text="Browse", command=lambda: self.browse_folder(self.dest_entry)).grid(row=3, column=2, padx=10, pady=5)
        
        # Duplicate Handling
        ttk.Label(self.tab, text="Handle Duplicates:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.duplicates_var = tk.StringVar(value="rename")
        duplicates_frame = ttk.Frame(self.tab)
        duplicates_frame.grid(row=4, column=1, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(duplicates_frame, text="Rename", variable=self.duplicates_var, value="rename").pack(side="left")
        ttk.Radiobutton(duplicates_frame, text="Overwrite", variable=self.duplicates_var, value="overwrite").pack(side="left", padx=10)
        ttk.Radiobutton(duplicates_frame, text="Skip", variable=self.duplicates_var, value="skip").pack(side="left")
        
        # Action Buttons
        btn_frame = ttk.Frame(self.tab)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=10)
        ttk.Button(btn_frame, text="Preview", command=self.preview_organization).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Execute", command=self.execute_organization).pack(side="left", padx=10)
        
        # Log Area
        ttk.Label(self.tab, text="Operation Log:").grid(row=6, column=0, sticky="nw", padx=10, pady=5)
        
        log_frame = ttk.Frame(self.tab)
        log_frame.grid(row=7, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        
        # Status Bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        ttk.Label(self.tab, textvariable=self.status_var, relief="sunken").grid(
            row=8, column=0, columnspan=3, sticky="ew", padx=10, pady=5)
    
    def setup_extension_tab(self):
        # Extensions to organize
        ttk.Label(self.extension_tab, text="Extensions to organize (comma-separated):").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.extensions_entry = ttk.Entry(self.extension_tab, width=50)
        self.extensions_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        
        # Miscellaneous folder option
        self.misc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.extension_tab, text="Create 'Miscellaneous' folder for other extensions", 
                        variable=self.misc_var).grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5)
    
    def setup_size_tab(self):
        # Size criteria
        ttk.Label(self.size_tab, text="File size:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        size_frame = ttk.Frame(self.size_tab)
        size_frame.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        self.size_operator = ttk.Combobox(size_frame, values=["<", "<=", "=", ">=", ">"], width=3, state="readonly")
        self.size_operator.current(0)
        self.size_operator.pack(side="left")
        
        self.size_value = ttk.Entry(size_frame, width=10)
        self.size_value.pack(side="left", padx=5)
        
        self.size_unit = ttk.Combobox(size_frame, values=["bytes", "KB", "MB", "GB"], width=5, state="readonly")
        self.size_unit.current(1)
        self.size_unit.pack(side="left")
        
        # Folder naming
        ttk.Label(self.size_tab, text="Folder name pattern:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.size_folder_pattern = ttk.Combobox(self.size_tab, 
                                              values=["{operator}{value}{unit}", "Files {operator} {value}{unit}", "Size {operator} {value}{unit}"],
                                              state="readonly")
        self.size_folder_pattern.current(0)
        self.size_folder_pattern.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
    
    def setup_date_tab(self):
        # Date criteria
        ttk.Label(self.date_tab, text="Date criteria:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.date_criteria = ttk.Combobox(self.date_tab, 
                                        values=["Created on", "Created after", "Created before", "Between dates"],
                                        state="readonly")
        self.date_criteria.current(0)
        self.date_criteria.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        self.date_criteria.bind("<<ComboboxSelected>>", self.update_date_ui)
        
        # Date input frame
        self.date_input_frame = ttk.Frame(self.date_tab)
        self.date_input_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        
        # Group by options
        ttk.Label(self.date_tab, text="Group by:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.date_grouping = ttk.Combobox(self.date_tab, 
                                        values=["Single folder", "Year", "Month", "Day", "Year-Month"],
                                        state="readonly")
        self.date_grouping.current(0)
        self.date_grouping.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        # Initialize date UI
        self.update_date_ui()

    def update_date_ui(self, event=None):
        # Clear previous widgets
        for widget in self.date_input_frame.winfo_children():
            widget.destroy()
        
        criteria = self.date_criteria.get()
        
        if criteria in ["Created on", "Created after", "Created before"]:
            ttk.Label(self.date_input_frame, text="Date (dd-mm-yyyy):").pack(side="left")
            self.date_entry = ttk.Entry(self.date_input_frame, width=10)
            self.date_entry.pack(side="left", padx=5)
        elif criteria == "Between dates":
            ttk.Label(self.date_input_frame, text="From (dd-mm-yyyy):").pack(side="left")
            self.date_from_entry = ttk.Entry(self.date_input_frame, width=10)
            self.date_from_entry.pack(side="left", padx=5)
            
            ttk.Label(self.date_input_frame, text="To (dd-mm-yyyy):").pack(side="left", padx=10)
            self.date_to_entry = ttk.Entry(self.date_input_frame, width=10)
            self.date_to_entry.pack(side="left", padx=5)    
    

    def setup_name_tab(self):
        # Position options
        ttk.Label(self.name_tab, text="Search position:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.position_var = tk.StringVar(value="Anywhere")
        position_frame = ttk.Frame(self.name_tab)
        position_frame.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        ttk.Radiobutton(position_frame, text="Anywhere", variable=self.position_var, value="Anywhere").pack(side="left")
        ttk.Radiobutton(position_frame, text="Starts with", variable=self.position_var, value="Starts with").pack(side="left", padx=5)
        ttk.Radiobutton(position_frame, text="Ends with", variable=self.position_var, value="Ends with").pack(side="left", padx=5)
        
        # Name contains
        ttk.Label(self.name_tab, text="Text to search:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.name_contains_entry = ttk.Entry(self.name_tab)
        self.name_contains_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        
        # Folder naming
        ttk.Label(self.name_tab, text="Folder name:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.name_folder_pattern = ttk.Combobox(self.name_tab, 
                                            values=["Files containing '{text}'", "Text '{text}' files", "Custom"],
                                            state="readonly")
        self.name_folder_pattern.current(0)
        self.name_folder_pattern.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        # Custom name pattern
        self.custom_name_pattern = ttk.Entry(self.name_tab)
        self.custom_name_pattern.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.custom_name_pattern.grid_remove()
        
        self.name_folder_pattern.bind("<<ComboboxSelected>>", self.update_name_pattern_ui)
    
    def update_name_position_ui(self, *args):
        position = self.position_var.get()
        if position == "Anywhere":
            self.char_count_frame.grid_remove()
        else:
            self.char_count_frame.grid()
    
    def update_name_pattern_ui(self, event=None):
        if self.name_folder_pattern.get() == "Custom":
            self.custom_name_pattern.grid()
        else:
            self.custom_name_pattern.grid_remove()
    
    def show_calendar(self, entry_widget):
        top = tk.Toplevel(self.tab)
        self.calendar_windows.append(top)
        top.title("Select Date")
        
        cal = Calendar(top, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=20)
        
        def set_date():
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, cal.get_date())
            top.destroy()
            self.calendar_windows.remove(top)
        
        ttk.Button(top, text="Select", command=set_date).pack(pady=10)
    
    def browse_folder(self, entry_widget):
        folder_path = filedialog.askdirectory()
        if folder_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folder_path)
    
    def log_message(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.tab.update()
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def validate_date(self, date_str):
        try:
            return datetime.strptime(date_str, "%d-%m-%Y").date()
        except ValueError:
            return None
    
    def get_size_in_bytes(self, size_str, unit):
        try:
            size = float(size_str)
            if unit == "KB":
                return size * 1024
            elif unit == "MB":
                return size * 1024 * 1024
            elif unit == "GB":
                return size * 1024 * 1024 * 1024
            return size  # bytes
        except ValueError:
            return None
    
    def get_destination_folder(self, file_path, file_stat):
        base_folder = self.dest_entry.get()
        if not base_folder:
            return None
            
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        ext = ext.lower()[1:] if ext else "no_extension"
        size_bytes = file_stat.st_size
        date_created = datetime.fromtimestamp(file_stat.st_ctime)
        
        current_tab = self.org_notebook.tab(self.org_notebook.select(), "text")
        
        if current_tab == "By Extension":
            if not self.extensions_entry.get().strip():
                return base_folder
                
            specified_exts = [e.strip().lower() for e in self.extensions_entry.get().split(",")]
            if ext in specified_exts:
                return os.path.join(base_folder, f"{ext} files")
            elif self.misc_var.get():
                return os.path.join(base_folder, "Miscellaneous extension files")
            return None
        
        elif current_tab == "By Size":
            size_value = self.size_value.get().strip()
            if not size_value:
                return base_folder
                
            target_size = self.get_size_in_bytes(size_value, self.size_unit.get())
            if target_size is None:
                return None
                
            operator = self.size_operator.get()
            file_size = size_bytes
            
            # Check if file matches size criteria
            if operator == "<" and not (file_size < target_size):
                return None
            elif operator == "<=" and not (file_size <= target_size):
                return None
            elif operator == "=" and not (file_size == target_size):
                return None
            elif operator == ">=" and not (file_size >= target_size):
                return None
            elif operator == ">" and not (file_size > target_size):
                return None
                
            # Create folder name based on pattern
            pattern = self.size_folder_pattern.get()
            folder_name = pattern.format(
                operator=self.size_operator.get(),
                value=self.size_value.get(),
                unit=self.size_unit.get()
            )
            return os.path.join(base_folder, folder_name)
        
        elif current_tab == "By Date":
            date_criteria = self.date_criteria.get()
            grouping = self.date_grouping.get()
            
            # Check date criteria
            if date_criteria == "Created on":
                date_str = self.date_entry.get().strip()
                if not date_str:
                    return None
                    
                target_date = self.validate_date(date_str)
                if not target_date or date_created.date() != target_date:
                    return None
                    
                if grouping == "Single folder":
                    return os.path.join(base_folder, f"Created on {date_str}")
                elif grouping == "Year":
                    return os.path.join(base_folder, str(date_created.year))
                elif grouping == "Month":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m"))
                elif grouping == "Day":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m-%d"))
                elif grouping == "Year-Month":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m"))
                
            
            elif date_criteria == "Created after":
                date_str = self.date_entry.get().strip()
                if not date_str:
                    return None
                    
                target_date = self.validate_date(date_str)
                if not target_date or date_created.date() <= target_date:
                    return None
                    
                if grouping == "Single folder":
                    return os.path.join(base_folder, f"Created after {date_str}")
                elif grouping == "Year":
                    return os.path.join(base_folder, str(date_created.year))
                elif grouping == "Month":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m"))
                elif grouping == "Day":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m-%d"))
                elif grouping == "Year-Month":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m"))
                
            
            elif date_criteria == "Created before":
                date_str = self.date_entry.get().strip()
                if not date_str:
                    return None
                    
                target_date = self.validate_date(date_str)
                if not target_date or date_created.date() >= target_date:
                    return None
                    
                if grouping == "Single folder":
                    return os.path.join(base_folder, f"Created before {date_str}")
                elif grouping == "Year":
                    return os.path.join(base_folder, str(date_created.year))
                elif grouping == "Month":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m"))
                elif grouping == "Day":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m-%d"))
                elif grouping == "Year-Month":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m"))
              
            
            elif date_criteria == "Between dates":
                date_from_str = self.date_from_entry.get().strip()
                date_to_str = self.date_to_entry.get().strip()
                if not date_from_str or not date_to_str:
                    return None
                    
                date_from = self.validate_date(date_from_str)
                date_to = self.validate_date(date_to_str)
                if not date_from or not date_to or not (date_from <= date_created.date() <= date_to):
                    return None
                    
                if grouping == "Single folder":
                    return os.path.join(base_folder, f"Created between {date_from_str} and {date_to_str}")
                elif grouping == "Year":
                    return os.path.join(base_folder, str(date_created.year))
                elif grouping == "Month":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m"))
                elif grouping == "Day":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m-%d"))
                elif grouping == "Year-Month":
                    return os.path.join(base_folder, date_created.strftime("%Y-%m"))
                
        elif current_tab == "By Name":
            name_contains = self.name_contains_entry.get().strip()
            if not name_contains:
                return None
                
            position = self.position_var.get()
            filename_lower = filename.lower()
            search_text = name_contains.lower()
            
            if position == "Anywhere":
                if search_text not in filename_lower:
                    return None
            elif position == "Starts with":
                if not filename_lower.startswith(search_text):
                    return None
            elif position == "Ends with":
                if not filename_lower.endswith(search_text):
                    return None
                
            pattern = self.name_folder_pattern.get()
            if pattern == "Custom":
                custom_pattern = self.custom_name_pattern.get().strip()
                if not custom_pattern:
                    return None
                folder_name = custom_pattern.format(text=name_contains)
            else:
                folder_name = pattern.format(text=name_contains)
                
            return os.path.join(base_folder, folder_name)
    
    def handle_duplicate(self, dest_path):
        if not os.path.exists(dest_path):
            return dest_path
            
        handle_method = self.duplicates_var.get()
        
        if handle_method == "overwrite":
            return dest_path
        elif handle_method == "skip":
            return None
            
        # Default is "rename"
        counter = 1
        name, ext = os.path.splitext(dest_path)
        while os.path.exists(f"{name}_{counter}{ext}"):
            counter += 1
            
        return f"{name}_{counter}{ext}"
    
    def preview_organization(self):
        self.clear_log()
        source_folder = self.input_entry.get()
        
        if not source_folder or not os.path.exists(source_folder):
            messagebox.showerror("Error", "Please select a valid source folder")
            return
        
        self.log_message("=== PREVIEW MODE ===")
        self.log_message(f"Scanning: {source_folder}")
        
        try:
            file_count = 0
            for root_dir, _, files in os.walk(source_folder):
                for filename in files:
                    file_path = os.path.join(root_dir, filename)
                    try:
                        file_stat = os.stat(file_path)
                        
                        dest_folder = self.get_destination_folder(file_path, file_stat)
                        if dest_folder:
                            dest_path = os.path.join(dest_folder, filename)
                            self.log_message(f"{filename} → {dest_folder}")
                            file_count += 1
                    except Exception as e:
                        self.log_message(f"Error processing {filename}: {str(e)}")
            
            self.log_message(f"\nFound {file_count} files matching criteria")
            self.status_var.set(f"Preview complete: {file_count} files would be processed")
        except Exception as e:
            self.log_message(f"\nError during preview: {str(e)}")
            self.status_var.set("Preview failed")
    
    def execute_organization(self):
        self.clear_log()
        source_folder = self.input_entry.get()
        dest_base = self.dest_entry.get()
        operation = self.operation_var.get()
        
        if not source_folder or not os.path.exists(source_folder):
            messagebox.showerror("Error", "Please select a valid source folder")
            return
        
        if not dest_base:
            messagebox.showerror("Error", "Please select a destination folder")
            return
        
        self.log_message("=== EXECUTION MODE ===")
        self.log_message(f"Source: {source_folder}")
        self.log_message(f"Destination: {dest_base}")
        self.log_message(f"Operation: {operation}")
        
        try:
            processed = 0
            skipped = 0
            errors = 0
            
            for root_dir, _, files in os.walk(source_folder):
                for filename in files:
                    file_path = os.path.join(root_dir, filename)
                    try:
                        file_stat = os.stat(file_path)
                        
                        dest_folder = self.get_destination_folder(file_path, file_stat)
                        if not dest_folder:
                            skipped += 1
                            continue
                            
                        # Ensure destination folder exists
                        os.makedirs(dest_folder, exist_ok=True)
                        
                        dest_path = os.path.join(dest_folder, filename)
                        
                        # Handle duplicates
                        final_dest_path = self.handle_duplicate(dest_path)
                        if not final_dest_path:
                            self.log_message(f"Skipped duplicate: {filename}")
                            skipped += 1
                            continue
                            
                        # Perform operation
                        if operation == "Move":
                            shutil.move(file_path, final_dest_path)
                            action = "Moved"
                        else:
                            shutil.copy2(file_path, final_dest_path)
                            action = "Copied"
                        
                        self.log_message(f"{action}: {filename} → {os.path.dirname(final_dest_path)}")
                        processed += 1
                    except Exception as e:
                        self.log_message(f"Error processing {filename}: {str(e)}")
                        errors += 1
            
            self.log_message(f"\nOperation complete!")
            self.log_message(f"Files processed: {processed}")
            self.log_message(f"Files skipped: {skipped}")
            self.log_message(f"Errors encountered: {errors}")
            
            self.status_var.set(
                f"Completed: {processed} processed, {skipped} skipped, {errors} errors")
            
            messagebox.showinfo(
                "Complete", 
                f"Operation completed!\nProcessed: {processed}\nSkipped: {skipped}\nErrors: {errors}")
        except Exception as e:
            self.log_message(f"\nFatal error during operation: {str(e)}")
            self.status_var.set("Operation failed")
            messagebox.showerror("Error", f"Operation failed: {str(e)}")

# ==================== MAIN APPLICATION ====================
class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        self.setup_ui()
        
    def setup_ui(self):
        # Configure style for notebook tabs
        style = ttk.Style()
        style.configure("TNotebook.Tab", foreground="black", background="white")
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        self.invisible_tab = ttk.Frame(self.notebook)
        self.grid_tab = ttk.Frame(self.notebook)
        self.flatten_tab = ttk.Frame(self.notebook)
        self.file_organizer_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.invisible_tab, text="Invisible Grid Converter")
        self.notebook.add(self.grid_tab, text="Grid-Based Converter")
        self.notebook.add(self.flatten_tab, text="Flatten Folder")
        self.notebook.add(self.file_organizer_tab, text="File Organizer")
        
        # Initialize all tools
        self.invisible_converter = InvisibleGridConverter(self.invisible_tab)
        self.grid_converter = GridBasedConverter(self.grid_tab)
        self.flatten_tool = FlattenFolderTool(self.flatten_tab)
        self.file_organizer = FileOrganizerTool(self.file_organizer_tab)

if __name__ == "__main__":
    try:
        # Set theme if available
        from ttkthemes import ThemedTk
        root = ThemedTk(theme="arc")
    except ImportError:
        root = tk.Tk()
    
    app = PDFConverterApp(root)
    root.mainloop()