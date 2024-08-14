import os
import json
import csv
import PyPDF2
import tkinter as tk
from tkinter import filedialog
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
    
    def apply_theme_to_widget(widget):
        """Apply theme to a single widget."""
        if isinstance(widget, (tk.Label, tk.Entry, tk.Button)):
            widget.configure(bg=bg_color, fg=fg_color)
        elif isinstance(widget, tk.Checkbutton):
            widget.configure(bg=bg_color, fg=fg_color, selectcolor=bg_color)
        elif isinstance(widget, tk.Frame):
            widget.configure(bg=bg_color)
            # Apply theme recursively to all child widgets
            for child in widget.winfo_children():
                apply_theme_to_widget(child)
    
    # Apply theme to the root window
    root.configure(bg=bg_color)
    # Apply theme to all widgets in the root window
    for widget in root.winfo_children():
        apply_theme_to_widget(widget)
    
    # Apply theme to ttk widgets
    style.configure('TButton', background=bg_color, foreground=fg_color)
    
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

def save_results_as_csv(results, filename):
    """Save the search results to a CSV file, including the search query."""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(["Query", "File Name", "File Path", "Pages Found"])  # Correct header row
            for result in results:
                csvwriter.writerow(result)
        status_label.config(text=f"Results saved to {filename}", fg="green")
    except Exception as e:
        status_label.config(text=f"Failed to save results: {e}", fg="red")

def search():
    global results  # Declare results as global to access in save_results
    directory = folder_path.get()
    keyword = keyword_entry.get()
    case_sensitive = case_sensitive_var.get()

    if not directory:
        status_label.config(text="Please select a directory containing PDFs.", fg="red")
        return
    
    if not keyword:
        status_label.config(text="Please enter a keyword to search.", fg="red")
        return

    for i in tree.get_children():
        tree.delete(i)
    
    progress_var.set(0)
    progress_bar.update()

    results = []
    search_results = search_keyword_in_pdfs(directory, keyword, case_sensitive)
    for filename, file_path, pages in search_results:
        results.append((keyword, filename, file_path, pages))  # Include query in the results

    if results:
        for result in results:
            query, filename, file_path, pages = result
            tree.insert("", "end", values=(query, filename, file_path, pages))
        
        # Update progress bar to full
        progress_var.set(100)
        progress_bar.update()

        # Enable the Save Results button
        save_button.config(state=tk.NORMAL)
        status_label.config(text="Search complete.", fg="green")

    else:
        status_label.config(text="Keyword not found in any PDF files.", fg="red")
        progress_var.set(100)
        progress_bar.update()

def batch_search():
    global results
    results = []  # Clear previous results
    tree.delete(*tree.get_children())  # Clear previous results in the treeview
    total_files = 0
    current_file = 0

    # Calculate total files for progress bar
    for query, case_sensitive in batch_queries:
        total_files += sum([len(files) for r, d, files in os.walk(folder_path.get()) if any(f.endswith('.pdf') for f in files)])

    for query, case_sensitive in batch_queries:
        search_results = search_keyword_in_pdfs(folder_path.get(), query, case_sensitive)
        for result in search_results:
            filename, filepath, pages = result
            results.append((query, filename, filepath, pages))  # Include query in the results

        # Update progress bar
        current_file += len(search_results)
        update_progress_bar(current_file, total_files)

    if results:
        for result in results:
            query, filename, filepath, pages = result
            tree.insert("", "end", values=(query, filename, filepath, pages))

        progress_var.set(100)
        progress_bar.update()
        status_label.config(text="Batch search complete.", fg="green")
    else:
        status_label.config(text="No results found in batch search.", fg="red")
        progress_var.set(100)
        progress_bar.update()

def save_batch_results():
    keyword = keyword_entry.get()
    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV files", "*.csv")],
                                             initialfile=f"search_results_{keyword}.csv")
    if file_path:
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                csvwriter = csv.writer(csvfile)
                csvwriter.writerow(["Search Query", "File Name", "File Path", "Pages Found"])  # Header row
                for result in results:
                    csvwriter.writerow(result)
            status_label.config(text=f"Batch results saved to {file_path}", fg="green")
        except Exception as e:
            status_label.config(text=f"Failed to save results: {e}", fg="red")

def add_query_to_batch():
    """Add the current query to the batch list."""
    keyword = keyword_entry.get()
    case_sensitive = case_sensitive_var.get()

    if keyword:
        batch_queries.append((keyword, case_sensitive))
        update_batch_list()
        keyword_entry.delete(0, tk.END)
    else:
        status_label.config(text="Please enter a keyword to add to batch.", fg="red")

def update_batch_list():
    """Update the batch list display."""
    batch_list.delete(0, tk.END)
    for i, (query, case_sensitive) in enumerate(batch_queries):
        cs_text = " (Case Sensitive)" if case_sensitive else ""
        batch_list.insert(tk.END, f"{i+1}. {query}{cs_text}")

def remove_selected_query():
    """Remove the selected query from the batch list."""
    selected = batch_list.curselection()
    if selected:
        index = selected[0]
        batch_queries.pop(index)
        update_batch_list()

def clear_batch_query():
    batch_queries.clear()
    update_batch_list()

def copy_path():
    selected_item = tree.selection()
    if selected_item:
        file_path = tree.item(selected_item[0], "values")[2]
        root.clipboard_clear()
        root.clipboard_append(file_path)
        root.update()  # Keep the clipboard updated
def open_file():
    selected_item = tree.selection()
    if selected_item:
        file_path = tree.item(selected_item[0], "values")[2]
        root.clipboard_clear()
        root.clipboard_append(file_path)
        root.update()  # Keep the clipboard updated 
        os.startfile(file_path)

# Initialize the Tk root window
root = tk.Tk()
# Define styles
style = ttk.Style(root)
style.theme_use("clam")
root.title("PDF Keyword Search")
root.geometry("800x600")

# Define variables after initializing root
folder_path = tk.StringVar()  # For directory entry
case_sensitive_var = tk.BooleanVar()  # For case sensitive checkbox
batch_queries = []  # For batch queries
results = []  # For storing results

style = ttk.Style()

# Widgets
directory_label = tk.Label(root, text="Directory:")
directory_label.grid(row=0, column=0, sticky="e", padx=10, pady=5)

directory_entry = tk.Entry(root, textvariable=folder_path, width=40)
directory_entry.grid(row=0, column=1, padx=5, pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_directory)
browse_button.grid(row=0, column=2, padx=5, pady=5)

keyword_label = tk.Label(root, text="Keyword:")
keyword_label.grid(row=1, column=0, sticky="e", padx=10, pady=5)

keyword_entry = tk.Entry(root, width=40)
keyword_entry.grid(row=1, column=1, padx=5, pady=5)

case_sensitive_checkbox = tk.Checkbutton(root, text="Case Sensitive", variable=case_sensitive_var)
case_sensitive_checkbox.grid(row=1, column=2, padx=5, pady=5)

search_button = tk.Button(root, text="Search", command=search)
search_button.grid(row=2, column=1, padx=10, pady=5)

# Batch Search Controls
batch_controls_frame = tk.Frame(root)
batch_controls_frame.grid(row=3, column=1, padx=10, pady=5)

add_to_batch_button = tk.Button(batch_controls_frame, text="Add to Batch", command=add_query_to_batch)
add_to_batch_button.grid(row=0, column=0, padx=5, pady=5)

batch_search_button = tk.Button(batch_controls_frame, text="Batch Search", command=batch_search)
batch_search_button.grid(row=0, column=1, padx=5, pady=5)

remove_from_batch_button = tk.Button(batch_controls_frame, text="Remove Selected", command=remove_selected_query)
remove_from_batch_button.grid(row=0, column=2, padx=5, pady=5)

clear_batch_button = tk.Button(batch_controls_frame, text="Clear Batch", command=clear_batch_query)
clear_batch_button.grid(row=0, column=3, padx=5, pady=5)

batch_list = tk.Listbox(batch_controls_frame, width=60, height=5)
batch_list.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

# Results
tree = ttk.Treeview(root, columns=("Query", "File Name", "File Path", "Pages Found"), show="headings", selectmode="browse")
tree.heading("Query", text="Query")
tree.heading("File Name", text="File Name")
tree.heading("File Path", text="File Path")
tree.heading("Pages Found", text="Pages Found")
tree.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

tree_scroll = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
tree_scroll.grid(row=4, column=3, sticky="ns")
tree.configure(yscrollcommand=tree_scroll.set)

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

save_button = tk.Button(root, text="Save Results", command=save_batch_results, state=tk.DISABLED)
save_button.grid(row=6, column=0, padx=10, pady=5)

toggle_theme_button = tk.Button(root, text="Toggle Theme", command=toggle_theme)
toggle_theme_button.grid(row=6, column=1, padx=10, pady=5)

status_label = tk.Label(root, text="", fg="red")
status_label.grid(row=7, column=0, columnspan=3, padx=10, pady=5)

# Configure grid scaling
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)
root.columnconfigure(2, weight=1)
root.rowconfigure(4, weight=1)  # Adjust row weight for Treeview
root.rowconfigure(8, weight=0)  # Ensure the last row has no extra space

# Load and apply the saved theme
current_theme = load_theme()
apply_theme(current_theme)

# Add context menu for copying file path
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Copy Path", command=copy_path)
context_menu.add_command(label="Open File", command=open_file)

def show_context_menu(event):
    context_menu.post(event.x_root, event.y_root)

tree.bind("<Button-3>", show_context_menu)
root.mainloop()


