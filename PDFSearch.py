import os
import json
import csv
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Constants for theme
THEME_FILE = "theme.json"

def load_theme():
    """Load the saved theme from the configuration file."""
    if os.path.exists(THEME_FILE):
        with open(THEME_FILE, 'r') as file:
            return json.load(file).get('theme', 'light')
    return 'light'

def save_theme(theme):
    """Save the current theme to the configuration file."""
    with open(THEME_FILE, 'w') as file:
        json.dump({'theme': theme}, file)

def apply_theme(theme):
    """Apply the given theme to the application."""
    if theme == 'dark':
        bg_color = '#2E2E2E'
        fg_color = '#FFFFFF'
        highlight_color = '#424242'
        progress_color = '#32CD32'
        tree_bg_color = '#2E2E2E'
        scrollbar_bg_color = '#2E2E2E'
        scrollbar_fg_color = '#FFFFFF'
        row_bg_even = '#2E2E2E'
        row_bg_odd = '#2E2E2E'
    else:
        bg_color = '#FFFFFF'
        fg_color = '#000000'
        highlight_color = '#D9D9D9'
        progress_color = '#32CD32'
        tree_bg_color = '#FFFFFF'
        scrollbar_bg_color = '#FFFFFF'
        scrollbar_fg_color = '#000000'
        row_bg_even = '#FFFFFF'
        row_bg_odd = '#F0F0F0'
    
    # Apply theme to root window
    root.configure(bg=bg_color)
    
    # Apply theme to tk widgets
    for widget in root.winfo_children():
        if isinstance(widget, (tk.Label, tk.Entry, tk.Button)):
            widget.configure(bg=bg_color, fg=fg_color)
            
        elif isinstance(widget, tk.Checkbutton):
            widget.configure(bg=bg_color, fg=fg_color, selectcolor=bg_color)
        
            
        elif isinstance(widget, tk.Frame):
            widget.configure(bg=bg_color)
    
    # Apply theme to ttk widgets
    style.configure('TButton', background=bg_color, foreground=fg_color)
    #style.configure('TCheckbutton', bg=bg_color, fg=bg_color)
    
    # Treeview styling
    style.configure('Treeview', background=tree_bg_color, foreground=fg_color, fieldbackground=tree_bg_color)
    style.configure('Treeview.Heading', background=bg_color, fieldbackground=bg_color, foreground=fg_color)
    
    # Configure Treeview row colors
    tree.tag_configure('evenrow', background=row_bg_even)
    tree.tag_configure('oddrow', background=row_bg_odd)
    
    # Apply theme to progress bar
    style.configure('TProgressbar', troughcolor=highlight_color, background=progress_color)

    # Apply theme to scrollbars
    style.configure('Vertical.TScrollbar', background=scrollbar_bg_color, troughcolor=highlight_color, arrowcolor=scrollbar_fg_color)
    style.configure('Horizontal.TScrollbar', background=scrollbar_bg_color, troughcolor=highlight_color, arrowcolor=scrollbar_fg_color)

def toggle_theme():
    """Toggle between light and dark themes."""
    current_theme = load_theme()
    new_theme = 'dark' if current_theme == 'light' else 'light'
    apply_theme(new_theme)
    save_theme(new_theme)

def search_keyword_in_pdfs(directory, keyword, case_sensitive):
    results = []
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

def save_results_as_csv(results, filename, query):
    """Save the search results to a CSV file, including the search query."""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(["Search Query", query])  # Write the search query
            csvwriter.writerow(["File Name", "File Path", "Pages Found"])  # Header row
            for result in results:
                csvwriter.writerow(result)
        messagebox.showinfo("Success", f"Results saved to {filename}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save results: {e}")

def search():
    global results  # Declare results as global to access in save_results
    directory = folder_path.get()
    keyword = keyword_entry.get()
    case_sensitive = case_sensitive_var.get()

    if not directory:
        messagebox.showerror("Error", "Please select a directory containing PDFs.")
        return
    
    if not keyword:
        messagebox.showerror("Error", "Please enter a keyword to search.")
        return

    for i in tree.get_children():
        tree.delete(i)
    
    progress_var.set(0)
    progress_bar.update()

    results = search_keyword_in_pdfs(directory, keyword, case_sensitive)

    if results:
        for filename, file_path, pages in results:
            tree.insert("", "end", values=(filename, file_path, pages))
        
        # Update progress bar to full
        progress_var.set(100)
        progress_bar.update()

        # Enable the Save Results button
        save_button.config(state=tk.NORMAL)

    else:
        messagebox.showinfo("Results", "Keyword not found in any PDF files.")
        progress_var.set(100)
        progress_bar.update()

def save_results():
    keyword = keyword_entry.get()
    if not keyword:
        messagebox.showerror("Error", "No search query found. Perform a search first.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                           filetypes=[("CSV files", "*.csv")],
                                           initialfile=f"search_results_{keyword}.csv")
    if file_path:
        save_results_as_csv(results, file_path, keyword)

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

# Define styles
style = ttk.Style(root)
style.theme_use("clam")

# Create widgets
tk.Label(root, text="Select PDF Directory:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
folder_path = tk.StringVar()
tk.Entry(root, textvariable=folder_path, width=50).grid(row=0, column=1, padx=10, pady=10, sticky="ew")
tk.Button(root, text="Browse", command=browse_directory).grid(row=0, column=2, padx=10, pady=10, sticky="w")

tk.Label(root, text="Keyword:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
keyword_entry = tk.Entry(root, width=50)
keyword_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

case_sensitive_var = tk.BooleanVar()
tk.Checkbutton(root, text="Case Sensitive", variable=case_sensitive_var).grid(row=1, column=2, padx=10, pady=10, sticky="w")

tk.Button(root, text="Search", command=search).grid(row=2, column=2, padx=10, pady=10, sticky="w")
tk.Button(root, text="Toggle Dark Theme", command=toggle_theme).grid(row=2, column=0, padx=10, pady=10, sticky="w")

# Save Results button (initially disabled)
save_button = tk.Button(root, text="Save Results", command=save_results, state=tk.DISABLED)
save_button.grid(row=2, column=1, padx=10, pady=10, sticky="e")

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=400)
progress_bar.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

# Create the Treeview and scrollbars
tree = ttk.Treeview(root, columns=("File Name", "File Path", "Pages"), show="headings")
tree.heading("File Name", text="File Name")
tree.heading("File Path", text="File Path")
tree.heading("Pages", text="Pages Found")

tree.column("File Name", width=200, anchor="w")
tree.column("File Path", width=400, anchor="w")
tree.column("Pages", width=100, anchor="center")

# Add scrollbars
vsb = ttk.Scrollbar(root, orient="vertical", style='Vertical.TScrollbar', command=tree.yview)
vsb.grid(row=4, column=3, sticky="ns")
tree.configure(yscrollcommand=vsb.set)

hsb = ttk.Scrollbar(root, orient="horizontal", style='Horizontal.TScrollbar', command=tree.xview)
hsb.grid(row=5, column=0, columnspan=3, sticky="ew")
tree.configure(xscrollcommand=hsb.set)

tree.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Add context menu for copying file path
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Copy Path", command=copy_path)

def show_context_menu(event):
    context_menu.post(event.x_root, event.y_root)

tree.bind("<Button-3>", show_context_menu)

# Load and apply the saved theme
current_theme = load_theme()
apply_theme(current_theme)

# Configure grid scaling
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)
root.columnconfigure(2, weight=1)
root.rowconfigure(4, weight=1)
root.rowconfigure(5, weight=0)  # Ensure the last row has no extra space

root.mainloop()
