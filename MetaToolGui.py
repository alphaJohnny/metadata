import os
import subprocess
import hashlib
from tkinter import Tk, filedialog, Label, Button, ttk, messagebox
from openpyxl import load_workbook
from datetime import datetime
from PyPDF2 import PdfFileReader
from PIL import Image
from PIL.ExifTags import TAGS

def get_excel_metadata(file_path):
    """Retrieve metadata from an Excel file."""
    wb = load_workbook(file_path)
    metadata = wb.properties
    
    metadata_dict = {
        'Title': metadata.title,
        'Subject': metadata.subject,
        'Creator': metadata.creator,
        'Keywords': metadata.keywords,
        'Description': metadata.description,
        'Last Modified By': metadata.lastModifiedBy,
        'Revision': metadata.revision,
        'Created': metadata.created,
        'Modified': metadata.modified,
        'Category': metadata.category,
        'Content Status': metadata.contentStatus,
        'Language': metadata.language,
        'Identifier': metadata.identifier
    }
    
    return metadata_dict

def get_pdf_metadata(file_path):
    """Retrieve metadata from a PDF file."""
    with open(file_path, 'rb') as f:
        reader = PdfFileReader(f)
        metadata = reader.getDocumentInfo()
    
    metadata_dict = {key[1:]: value for key, value in metadata.items()}
    return metadata_dict

def get_image_metadata(file_path):
    """Retrieve metadata from an image file."""
    image = Image.open(file_path)
    info = image._getexif()
    
    metadata_dict = {TAGS.get(tag): value for tag, value in info.items()} if info else {}
    return metadata_dict

def update_file_system_dates(file_path, created, modified):
    """Update file system creation and modification dates to match file metadata."""
    if created:
        created_timestamp = datetime.timestamp(created)
        subprocess.run(['touch', '-t', datetime.fromtimestamp(created_timestamp).strftime('%Y%m%d%H%M.%S'), file_path])

    if modified:
        modified_timestamp = datetime.timestamp(modified)
        subprocess.run(['touch', '-mt', datetime.fromtimestamp(modified_timestamp).strftime('%Y%m%d%H%M.%S'), file_path])

def calculate_hash(file_path, algorithm='sha256'):
    """Calculate and return the hash of the file."""
    hash_func = hashlib.new(algorithm)
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b''):
            hash_func.update(chunk)
    return hash_func.hexdigest()

def compare_to_industry_standards(metadata):
    """Compare metadata to industry standards and identify potential issues."""
    issues = {}
    # Example standards
    if metadata.get('Creator') is None:
        issues['Creator'] = 'Missing creator'
    if metadata.get('Revision') is None or not metadata['Revision'].isdigit():
        issues['Revision'] = 'Invalid revision number'
    
    # Add more checks as needed
    return issues

def display_metadata(metadata, issues):
    """Display metadata in a table with potential issues."""
    for item in tree.get_children():
        tree.delete(item)
    for row, (key, value) in enumerate(metadata.items()):
        tree.insert("", "end", values=(key, value, issues.get(key, '')))

def on_select_file():
    try: 
        file_path = filedialog.askopenfilename(filetypes=[
            ("All Files", "*.*"),
            ("Excel files", "*.xlsx"),
            ("PDF files", "*.pdf"),
            ("Image files", "*.jpg;*.jpeg;*.png;*.gif")
        ])

        if not file_path:
            messagebox.showwarning("No Selection", "No file was selected.")
            return

        if file_path:
            if file_path.endswith('.xlsx'):
                metadata = get_excel_metadata(file_path)
            elif file_path.endswith('.pdf'):
                metadata = get_pdf_metadata(file_path)
            elif file_path.lower().endswith(('.jpg', '.jpeg', '.png', '.gif')):
                metadata = get_image_metadata(file_path)
            else:
                messagebox.showerror("Unsupported File", "The selected file type is not supported.")
                return

        issues = compare_to_industry_standards(metadata)
        display_metadata(metadata, issues)
        file_label.config(text=f"Selected file: {file_path}")
        global selected_file
        selected_file = file_path

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while selecting the file: {str(e)}")


def on_save():
    if not selected_file:
        messagebox.showerror("Error", "No file selected")
        return

    new_metadata = {}
    for child in tree.get_children():
        key = tree.item(child, 'values')[0]
        value = tree.item(child, 'values')[1]
        new_metadata[key] = value
    
    if selected_file.endswith('.xlsx'):
        wb = load_workbook(selected_file)
        wb.properties.title = new_metadata.get('Title', wb.properties.title)
        wb.properties.subject = new_metadata.get('Subject', wb.properties.subject)
        wb.properties.creator = new_metadata.get('Creator', wb.properties.creator)
        wb.properties.keywords = new_metadata.get('Keywords', wb.properties.keywords)
        wb.properties.description = new_metadata.get('Description', wb.properties.description)
        wb.properties.lastModifiedBy = new_metadata.get('Last Modified By', wb.properties.lastModifiedBy)
        wb.properties.revision = str(int(wb.properties.revision) + 1) if wb.properties.revision else '1'
        wb.properties.created = new_metadata.get('Created', wb.properties.created)
        wb.properties.modified = new_metadata.get('Modified', wb.properties.modified)
        wb.properties.category = new_metadata.get('Category', wb.properties.category)
        wb.properties.contentStatus = new_metadata.get('Content Status', wb.properties.contentStatus)
        wb.properties.language = new_metadata.get('Language', wb.properties.language)
        wb.properties.identifier = new_metadata.get('Identifier', wb.properties.identifier)
        
        wb.save(selected_file)
        
        created = datetime.strptime(new_metadata['Created'], "%Y-%m-%d %H:%M:%S") if 'Created' in new_metadata else None
        modified = datetime.strptime(new_metadata['Modified'], "%Y-%m-%d %H:%M:%S") if 'Modified' in new_metadata else None
        update_file_system_dates(selected_file, created, modified)
    else:
        messagebox.showerror("Unsupported File", "Saving metadata is only supported for Excel files in this version.")
        return

    file_hash = calculate_hash(selected_file)
    hash_label.config(text=f"File Hash (SHA-256): {file_hash}")

    messagebox.showinfo("Success", "File metadata updated successfully")

# GUI setup
root = Tk()
root.title("Forensic Audit Tool")
root.geometry("800x600")

selected_file = None

# File selection
file_label = Label(root, text="No file selected")
file_label.pack(pady=10)
select_button = Button(root, text="Select File", command=on_select_file)
select_button.pack(pady=10)

# Metadata table
columns = ('Property', 'Value', 'Issues')
tree = ttk.Treeview(root, columns=columns, show='headings')
tree.heading('Property', text='Property')
tree.heading('Value', text='Value')
tree.heading('Issues', text='Issues')
tree.pack(pady=20, fill='both', expand=True)

# Save button
save_button = Button(root, text="Save Changes", command=on_save)
save_button.pack(pady=10)

# Hash label
hash_label = Label(root, text="File Hash (SHA-256): N/A")
hash_label.pack(pady=10)

root.mainloop()
