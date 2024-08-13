import os
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def search_keyword_in_pdfs(directory, keyword, case_sensitive):
    results = []

    # Count the total number of PDF files
    total_files = sum([len(files) for r, d, files in os.walk(directory) if any(f.endswith('.pdf') for f in files)])
    current_file = 0

    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename.endswith('.pdf'):
                file_path = os.path.join(root, filename)
                try:
                    with open(file_path, 'rb') as pdf_file:
                        pdf_reader = PyPDF2.PdfReader(pdf_file)
                        num_pages = len(pdf_reader.pages)
                        pages_found = []

                        for page_num in range(num_pages):
                            page = pdf_reader.pages[page_num]
                            text = page.extract_text()

                            if not case_sensitive:
                                if keyword.lower() in text.lower():
                                    pages_found.append(page_num + 1)
                            else:
                                if keyword in text:
                                    pages_found.append(page_num + 1)

                        if pages_found:
                            results.append((filename, file_path, ', '.join(map(str, pages_found))))

                except PyPDF2.errors.PdfReadError as e:
                    print(f"Skipped unreadable PDF: {file_path} due to error: {e}")
                except Exception as e:
                    print(f"An unexpected error occurred with file {file_path}: {e}")

                current_file += 1
                update_progress_bar(current_file, total_files)

    return results

def update_progress_bar(current, total):
    progress_var.set((current / total) * 100)
    progress_bar.update()

def browse_directory():
    folder_selected = filedialog.askdirectory()
    folder_path.set(folder_selected)

def search():
    directory = folder_path.get()
    keyword = keyword_entry.get()
    case_sensitive = case_sensitive_var.get()

    if not directory:
        messagebox.showerror("Error", "Please select a directory containing PDFs.")
        return
    
    if not keyword:
        messagebox.showerror("Error", "Please enter a keyword to search.")
        return

    # Clear previous results
    for i in tree.get_children():
        tree.delete(i)
    
    progress_var.set(0)
    progress_bar.update()

    # Run search
    results = search_keyword_in_pdfs(directory, keyword, case_sensitive)

    if results:
        for filename, file_path, pages in results:
            tree.insert("", "end", values=(filename, file_path, pages))
    else:
        messagebox.showinfo("Results", "Keyword not found in any PDF files.")

def copy_path():
    selected_item = tree.selection()
    if selected_item:
        file_path = tree.item(selected_item[0], "values")[1]
        root.clipboard_clear()
        root.clipboard_append(file_path)
        root.update()  # Keep the clipboard updated

# Set up the main application window
root = tk.Tk()
root.title("PDF Keyword Search")

# Set the window icon
root.iconbitmap("icon.ico")  # Use your .ico file here

# Make the UI scalable
root.columnconfigure(1, weight=1)
root.rowconfigure(3, weight=1)

# Variables
folder_path = tk.StringVar()
progress_var = tk.DoubleVar()
case_sensitive_var = tk.BooleanVar()

# Folder selection
tk.Label(root, text="Select PDF Directory:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=folder_path, width=50).grid(row=0, column=1, padx=10, pady=10, sticky="ew")
tk.Button(root, text="Browse", command=browse_directory).grid(row=0, column=2, padx=10, pady=10)

# Keyword entry
tk.Label(root, text="Keyword:").grid(row=1, column=0, padx=10, pady=10)
keyword_entry = tk.Entry(root, width=50)
keyword_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

# Case sensitive checkbox
tk.Checkbutton(root, text="Case Sensitive", variable=case_sensitive_var).grid(row=1, column=2, padx=10, pady=10)

# Search button
tk.Button(root, text="Search", command=search).grid(row=2, column=2, padx=10, pady=10)

# Progress bar
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=400)
progress_bar.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

# Results table
tree = ttk.Treeview(root, columns=("File Name", "File Path", "Pages"), show="headings")
tree.heading("File Name", text="File Name")
tree.heading("File Path", text="File Path")
tree.heading("Pages", text="Pages Found")

tree.column("File Name", width=200, anchor="w")
tree.column("File Path", width=400, anchor="w")
tree.column("Pages", width=100, anchor="center")

# Add vertical scrollbar to the table
vsb = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
vsb.grid(row=4, column=3, sticky="ns")

tree.configure(yscrollcommand=vsb.set)

# Add horizontal scrollbar to the table
hsb = ttk.Scrollbar(root, orient="horizontal", command=tree.xview)
hsb.grid(row=5, column=0, columnspan=3, sticky="ew")

tree.configure(xscrollcommand=hsb.set)

tree.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Add context menu for copying file path
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Copy Path", command=copy_path)

def show_context_menu(event):
    context_menu.post(event.x_root, event.y_root)

tree.bind("<Button-3>", show_context_menu)

# Start the application
root.mainloop()
