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
        
        self.notebook.add(self.invisible_tab, text="Invisible Grid Converter")
        self.notebook.add(self.grid_tab, text="Grid-Based Converter")
        
        # Initialize both converters
        self.invisible_converter = InvisibleGridConverter(self.invisible_tab)
        self.grid_converter = GridBasedConverter(self.grid_tab)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()