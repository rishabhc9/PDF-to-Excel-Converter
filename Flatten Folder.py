import os
import shutil
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

def browse_folder(entry_widget):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_widget.delete(0, END)
        entry_widget.insert(0, folder_path)

def extract_files():
    source_folder = input_entry.get()
    destination_folder = output_entry.get()
    operation = operation_var.get()  # 'Move' or 'Copy'
    extensions = extensions_entry.get().strip()
    all_extensions = all_extensions_var.get()  # Boolean
    handle_duplicates = duplicates_var.get()  # 'keep', 'overwrite', or 'rename'
    
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
    log_text.delete(1.0, END)
    log_text.insert(END, f"Starting {operation.lower()} operation...\n")
    log_text.insert(END, f"Duplicate handling: {handle_duplicates}\n")
    log_text.see(END)
    root.update()
    
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
                                        log_text.insert(END, f"Skipped duplicate: {file_path}\n")
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
                                        log_text.insert(END, f"Moved: {file_path} → {dest_path}\n")
                                    else:
                                        shutil.copy2(file_path, dest_path)
                                        log_text.insert(END, f"Copied: {file_path} → {dest_path}\n")
                                    processed_files += 1
                                except Exception as e:
                                    log_text.insert(END, f"Error processing {file_path}: {str(e)}\n")
                                
                                log_text.see(END)
                                root.update()
                    else:
                        # If it's a file, check extension if needed
                        if extensions is None or os.path.splitext(item)[1].lower() in extensions:
                            dest_path = os.path.join(destination_folder, item)
                            
                            # Handle duplicates based on user selection
                            if os.path.exists(dest_path):
                                if handle_duplicates == 'skip':
                                    log_text.insert(END, f"Skipped duplicate: {item_path}\n")
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
                                    log_text.insert(END, f"Moved: {item_path} → {dest_path}\n")
                                else:
                                    shutil.copy2(item_path, dest_path)
                                    log_text.insert(END, f"Copied: {item_path} → {dest_path}\n")
                                processed_files += 1
                            except Exception as e:
                                log_text.insert(END, f"Error processing {item_path}: {str(e)}\n")
                            
                            log_text.see(END)
                            root.update()
        
        log_text.insert(END, f"\nOperation completed!\n")
        log_text.insert(END, f"Files processed: {processed_files}\n")
        log_text.insert(END, f"Files skipped: {skipped_files}\n")
        messagebox.showinfo("Success", f"Operation completed!\nProcessed: {processed_files} files\nSkipped: {skipped_files} files")
    except Exception as e:
        log_text.insert(END, f"\nError: {str(e)}\n")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    finally:
        log_text.see(END)

def main():
    global root, input_entry, output_entry, operation_var, extensions_entry, all_extensions_var, duplicates_var, log_text
    
    root = Tk()
    root.title("Flatten Folder Tool")
    
    # Configure grid weights for proper resizing
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(7, weight=1)
    
    # Input Folder
    Label(root, text="Source Folder:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    input_entry = Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
    Button(root, text="Browse", command=lambda: browse_folder(input_entry)).grid(row=0, column=2, padx=10, pady=5)
    
    # Output Folder
    Label(root, text="Destination Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    output_entry = Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
    Button(root, text="Browse", command=lambda: browse_folder(output_entry)).grid(row=1, column=2, padx=10, pady=5)
    
    # Operation Type (Move/Copy)
    Label(root, text="Operation:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    operation_var = StringVar(value="Move")
    operation_menu = ttk.Combobox(root, textvariable=operation_var, values=["Move", "Copy"], state="readonly", width=47)
    operation_menu.grid(row=2, column=1, padx=10, pady=5, sticky="w")
    
    # File Extensions
    Label(root, text="File Extensions (comma-separated):").grid(row=3, column=0, sticky="w", padx=10, pady=5)
    extensions_entry = Entry(root, width=50)
    extensions_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
    
    # All Extensions Checkbox
    all_extensions_var = BooleanVar()
    all_extensions_cb = Checkbutton(root, text="All Extensions", variable=all_extensions_var)
    all_extensions_cb.grid(row=3, column=2, padx=10, pady=5, sticky="w")
    
    # Duplicate Handling
    Label(root, text="Handle Duplicates:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
    duplicates_var = StringVar(value="rename")
    duplicates_frame = Frame(root)
    duplicates_frame.grid(row=4, column=1, sticky="w", padx=10, pady=5)
    Radiobutton(duplicates_frame, text="Rename", variable=duplicates_var, value="rename").pack(side=LEFT)
    Radiobutton(duplicates_frame, text="Overwrite", variable=duplicates_var, value="overwrite").pack(side=LEFT, padx=10)
    Radiobutton(duplicates_frame, text="Skip", variable=duplicates_var, value="skip").pack(side=LEFT)
    
    # Execute Button
    Button(root, text="Extract Files", command=extract_files, width=20).grid(row=5, column=1, pady=10)

    # Log Area
    Label(root, text="Operation Log:").grid(row=6, column=0, sticky="nw", padx=10, pady=5)
    
    # Create a frame for the log text with integrated scrollbar
    log_frame = Frame(root)
    log_frame.grid(row=7, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
    log_frame.grid_columnconfigure(0, weight=1)
    log_frame.grid_rowconfigure(0, weight=1)
    
    # Create Text widget with integrated scrollbar
    log_text = Text(log_frame, wrap=WORD)
    log_text.grid(row=0, column=0, sticky="nsew")
    
    # Add scrollbar
    scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    scrollbar.grid(row=0, column=1, sticky="ns")
    log_text['yscrollcommand'] = scrollbar.set
    
    root.mainloop()

if __name__ == "__main__":
    main()