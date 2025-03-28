import os
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkcalendar import Calendar

class AdvancedFileOrganizer:
    def __init__(self, root):
        self.root = root
        self.root.title("File Organizer")
        self.setup_ui()
        self.calendar_windows = []
        
    def setup_ui(self):
        # Configure grid weights
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(10, weight=1)
        
        # Style configuration for notebook tabs
        style = ttk.Style()
        style = ttk.Style()
        style.configure("TNotebook.Tab", foreground="black", background="white")
        style.map("TNotebook.Tab", foreground=[("selected", "black")])

        # Notebook for different organization methods
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=10, pady=5)

        # Create tabs (Initialize before calling setup functions)
        self.extension_tab = ttk.Frame(self.notebook)
        self.size_tab = ttk.Frame(self.notebook)
        self.date_tab = ttk.Frame(self.notebook)
        self.name_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.extension_tab, text="By Extension")
        self.notebook.add(self.size_tab, text="By Size")
        self.notebook.add(self.date_tab, text="By Date")
        self.notebook.add(self.name_tab, text="By Name")

        # Call setup functions AFTER initializing the tabs
        self.setup_extension_tab()

        
        # Common widgets for all tabs
        self.setup_common_widgets()
        
        # Tab-specific widgets
        self.setup_extension_tab()
        self.setup_size_tab()
        self.setup_date_tab()
        self.setup_name_tab()
    
    def setup_common_widgets(self):
        # Input Folder
        ttk.Label(self.root, text="Source Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.input_entry = ttk.Entry(self.root, width=60)
        self.input_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(self.root, text="Browse", command=lambda: self.browse_folder(self.input_entry)).grid(row=1, column=2, padx=10, pady=5)
        
        # Operation Type
        ttk.Label(self.root, text="Operation:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.operation_var = tk.StringVar(value="Copy")
        ttk.Combobox(self.root, textvariable=self.operation_var, 
                    values=["Copy", "Move"], state="readonly").grid(row=2, column=1, padx=10, pady=5, sticky="w")
        
        # Destination Folder
        ttk.Label(self.root, text="Destination Folder:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.dest_entry = ttk.Entry(self.root, width=60)
        self.dest_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(self.root, text="Browse", command=lambda: self.browse_folder(self.dest_entry)).grid(row=3, column=2, padx=10, pady=5)
        
        # Duplicate Handling
        ttk.Label(self.root, text="Handle Duplicates:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.duplicates_var = tk.StringVar(value="rename")
        duplicates_frame = ttk.Frame(self.root)
        duplicates_frame.grid(row=4, column=1, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(duplicates_frame, text="Rename", variable=self.duplicates_var, value="rename").pack(side="left")
        ttk.Radiobutton(duplicates_frame, text="Overwrite", variable=self.duplicates_var, value="overwrite").pack(side="left", padx=10)
        ttk.Radiobutton(duplicates_frame, text="Skip", variable=self.duplicates_var, value="skip").pack(side="left")
        
        # Action Buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=10)
        ttk.Button(btn_frame, text="Preview", command=self.preview_organization).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Execute", command=self.execute_organization).pack(side="left", padx=10)
        
        # Log Area
        ttk.Label(self.root, text="Operation Log:").grid(row=6, column=0, sticky="nw", padx=10, pady=5)
        
        log_frame = ttk.Frame(self.root)
        log_frame.grid(row=7, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        
        # Status Bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        ttk.Label(self.root, textvariable=self.status_var).grid(
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
                                        values=["Single folder", "Year", "Month", "Day", "Year-Month", "Custom"],
                                        state="readonly")
        self.date_grouping.current(0)
        self.date_grouping.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        # Custom date format
        ttk.Label(self.date_tab, text="Custom date format:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.custom_date_format = ttk.Entry(self.date_tab)
        self.custom_date_format.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        ttk.Label(self.date_tab, text="(e.g., %Y-%m-%d)").grid(row=4, column=1, sticky="w", padx=10)
        
        # Initialize date UI
        self.update_date_ui()
    
    def update_date_ui(self, event=None):
        # Clear previous widgets
        for widget in self.date_input_frame.winfo_children():
            widget.destroy()
        
        criteria = self.date_criteria.get()
        
        if criteria in ["Created on", "Created after", "Created before"]:
            ttk.Label(self.date_input_frame, text="Date:").pack(side="left")
            self.date_entry = ttk.Entry(self.date_input_frame, width=10)
            self.date_entry.pack(side="left", padx=5)
            ttk.Button(self.date_input_frame, text="📅", command=lambda: self.show_calendar(self.date_entry)).pack(side="left")
        elif criteria == "Between dates":
            ttk.Label(self.date_input_frame, text="From:").pack(side="left")
            self.date_from_entry = ttk.Entry(self.date_input_frame, width=10)
            self.date_from_entry.pack(side="left", padx=5)
            ttk.Button(self.date_input_frame, text="📅", command=lambda: self.show_calendar(self.date_from_entry)).pack(side="left")
            
            ttk.Label(self.date_input_frame, text="To:").pack(side="left", padx=10)
            self.date_to_entry = ttk.Entry(self.date_input_frame, width=10)
            self.date_to_entry.pack(side="left", padx=5)
            ttk.Button(self.date_input_frame, text="📅", command=lambda: self.show_calendar(self.date_to_entry)).pack(side="left")
    
    def setup_name_tab(self):
        # Position options
        ttk.Label(self.name_tab, text="Search position:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.position_var = tk.StringVar(value="Anywhere")
        position_frame = ttk.Frame(self.name_tab)
        position_frame.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        ttk.Radiobutton(position_frame, text="Anywhere", variable=self.position_var, value="Anywhere").pack(side="left")
        ttk.Radiobutton(position_frame, text="Starts with", variable=self.position_var, value="Starts with").pack(side="left", padx=5)
        ttk.Radiobutton(position_frame, text="Ends with", variable=self.position_var, value="Ends with").pack(side="left", padx=5)
        
        # Character count (only visible when position is not "Anywhere")
        self.char_count_frame = ttk.Frame(self.name_tab)
        self.char_count_frame.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        ttk.Label(self.char_count_frame, text="Number of characters:").pack(side="left")
        self.char_count = ttk.Entry(self.char_count_frame, width=5)
        self.char_count.pack(side="left", padx=5)
        
        # Initially hide character count frame
        self.char_count_frame.grid_remove()
        
        # Bind position change event
        self.position_var.trace_add("write", self.update_name_position_ui)
        
        # Name contains
        ttk.Label(self.name_tab, text="Text to search:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.name_contains_entry = ttk.Entry(self.name_tab)
        self.name_contains_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        # Folder naming
        ttk.Label(self.name_tab, text="Folder name:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.name_folder_pattern = ttk.Combobox(self.name_tab, 
                                              values=["Files containing '{text}'", "Text '{text}' files", "Custom"],
                                              state="readonly")
        self.name_folder_pattern.current(0)
        self.name_folder_pattern.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        
        # Custom name pattern
        self.custom_name_pattern = ttk.Entry(self.name_tab)
        self.custom_name_pattern.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
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
        top = tk.Toplevel(self.root)
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
        self.root.update()
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def validate_date(self, date_str):
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").date()
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
        
        current_tab = self.notebook.tab(self.notebook.select(), "text")
        
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
                elif grouping == "Custom":
                    fmt = self.custom_date_format.get().strip()
                    if not fmt:
                        fmt = "%Y-%m-%d"
                    return os.path.join(base_folder, date_created.strftime(fmt))
            
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
                elif grouping == "Custom":
                    fmt = self.custom_date_format.get().strip()
                    if not fmt:
                        fmt = "%Y-%m-%d"
                    return os.path.join(base_folder, date_created.strftime(fmt))
            
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
                elif grouping == "Custom":
                    fmt = self.custom_date_format.get().strip()
                    if not fmt:
                        fmt = "%Y-%m-%d"
                    return os.path.join(base_folder, date_created.strftime(fmt))
            
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
                elif grouping == "Custom":
                    fmt = self.custom_date_format.get().strip()
                    if not fmt:
                        fmt = "%Y-%m-%d"
                    return os.path.join(base_folder, date_created.strftime(fmt))
        
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
                char_count = self.char_count.get().strip()
                if char_count:
                    try:
                        n = int(char_count)
                        if not filename_lower.startswith(search_text[:n]):
                            return None
                    except ValueError:
                        pass
                else:
                    if not filename_lower.startswith(search_text):
                        return None
            elif position == "Ends with":
                char_count = self.char_count.get().strip()
                if char_count:
                    try:
                        n = int(char_count)
                        if not filename_lower.endswith(search_text[-n:]):
                            return None
                    except ValueError:
                        pass
                else:
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
        
        return base_folder
    
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

if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedFileOrganizer(root)
    root.mainloop()