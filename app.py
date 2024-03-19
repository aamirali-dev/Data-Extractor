import PyPDF2
import re
import pandas as pd 
import csv
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io, os 
from pdf2image import convert_from_path
from reportlab.platypus import Table, TableStyle, BaseDocTemplate, Frame, PageTemplate
from reportlab.lib import colors
from PIL import Image
import threading
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, Spacer
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from datetime import date, timedelta

from invoice import Invoice
from page import PdfPage
from extractor import PageExtractor
from pick_list import PickList


progress_bar = {}
WORKER_THREAD = None
PNGS_NOT_FOUND = []



def create_label_map(labeller_path):
    reader2 = PyPDF2.PdfReader(labeller_path)
    map = []
    index = -1
    for i, page in enumerate(reader2.pages):
        text = page.extract_text()
        if 'CUSTOMS DECLARATION' in text:
            map[index]['custom'] = i
        else:
            map.append({'post': i, 'custom': None})
            index += 1
    return  map
garment_pick_list = []
invoice = {

}

def add_to_garment_pick_list(page):
    for item in page['items']:
        garment_pick_list.append({'name': item['Garment Type'], 'size': item['Size'], 'color': item['Colour'], 'quantity': int(item['Quantity']), 'SKU TYPE': item['SKU'].split('-')[1]})

def read_csv_to_dict(csv_file):
    sku_details = {}

    with open(csv_file, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            sku = row['SKU']
            details = {key: value for key, value in row.items() if key != 'SKU'}
            sku_details[sku] = details

    return sku_details







def read_csv(filename):
    try:
        with open(filename) as file:
            data = file.readlines()
        data = {line.strip().split(',')[0]: line.strip().split(',')[1] for line in data}
        return data 
    except:
        return None


import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox

# paths = {
#     'SKU FILE': None,
#     'PACK FILE': None,
#     'POST FILE': None,
#     'IMAGE FOLDER': None,
#     'OUTPUT IMAGE FOLDER': None,
#     'OUTPUT FOLDER': None,
# }

paths = {
    'SKU FILE': r"data/SKU LIST 2024.csv",
    'PACK FILE': r"samples/29.2.24 - 8 Orders Pack 10001.pdf",
    'POST FILE': r"samples/29.2.24 - 8 Orders Post 10001.pdf",
    'IMAGE FOLDER': r"samples/PNG",
    'OUTPUT IMAGE FOLDER': r"samples/hbl",
    'OUTPUT FOLDER': r"samples/output",
}

files = ['SKU FILE', 'PACK FILE', 'POST FILE']
folders = ['IMAGE FOLDER', 'OUTPUT IMAGE FOLDER', 'OUTPUT FOLDER']

entries = {
    'SKU FILE': None,
    'PACK FILE': None,
    'POST FILE': None,
    'IMAGE FOLDER': None,
    'OUTPUT IMAGE FOLDER': None,
    'OUTPUT FOLDER': None,
}

def browse_file(index, entry_type='file'):
    if entry_type == 'file':
        file_path = filedialog.askopenfilename() 
    elif entry_type == 'folder':
        file_path = filedialog.askdirectory()
    if file_path:
        paths[index] = file_path
        entries[index].config(state='normal')
        entries[index].delete(0, tk.END)
        entries[index].insert(0, file_path)
        entries[index].config(state='readonly')


def check_paths(paths):
    for key, value in paths.items():
        if not value:
            messagebox.showerror("Path Not Selected", f"{key} is not selected. please select all paths to proceed")
            return False
    return True

def generate_results():
    global progress_bar
    try:
        progress_bar['value'] = 10
        if not check_paths(paths):
            progress_bar['value'] = 100
            return
        customer_id = os.path.splitext(os.path.basename(paths['PACK FILE']))[0].split(' ')[-1]
        extracted_info = extract_info_from_pdf(paths['PACK FILE'], paths['POST FILE'], paths['IMAGE FOLDER'], paths['OUTPUT IMAGE FOLDER'], paths['SKU FILE'])
        os.makedirs(paths['OUTPUT FOLDER'], exist_ok=True)
        with open(paths['OUTPUT FOLDER'] + '/OUTPUT.pdf', 'wb') as f:
            extracted_info.write(f)
        
        df = pd.DataFrame(garment_pick_list)
        grouped_df = df.groupby(['name', 'size', 'color'])['quantity'].sum().reset_index().sort_values(by=['name', 'size', 'color'])
        PickList(grouped_df).to_pdf(paths['OUTPUT FOLDER'] + '/pick_list.pdf')
        
        grouped_df = df.groupby(['SKU TYPE'])['quantity'].sum().reset_index()
        Invoice(grouped_df, customer_id).to_pdf(paths['OUTPUT FOLDER'] + '/invoice.pdf')
        
        progress_bar['value'] = 100
        messagebox.showinfo("Operation Completed", f"Output Files has been generated successfully")
    
    except Exception as e:
        print(e.with_traceback(None))
        messagebox.showerror("Unknow Error", f"Some Unknow Error Occured. Find Details Below \n{e}")
        progress_bar['value'] = 100

def handle_click():
    global WORKER_THREAD
    if WORKER_THREAD and WORKER_THREAD.is_alive():
        messagebox.showerror('Task In Progress', 'Another task is already running')
        return
    WORKER_THREAD = threading.Thread(target=generate_results, daemon=True)
    WORKER_THREAD.start()


root = tk.Tk()
root.title("Data Extraction and Processing")
 
def draw_widgets(keys, dictionary, j=0, entry_type='file'):
    for i, key in enumerate(keys):
        label = tk.Label(root, text=key)
        label.grid(row=i+j, column=0, padx=10, pady=5, sticky="w")

        entry = tk.Entry(root, width=50, font=("Arial", 13), readonlybackground="white", state='readonly')
        entry.grid(row=i+j, column=1, padx=0, pady=10)

        button = tk.Button(root, text="Browse", width=10 ,command=lambda idx=key: browse_file(idx, entry_type))
        button.grid(row=i+j, column=2, padx=0, pady=10)

        entries[key] = entry


if __name__ == "__main__":
    # draw_widgets(files, paths)
    # draw_widgets(folders, paths, j=len(files), entry_type='folder')
    # button = tk.Button(root, text="Generate Results", command=handle_click)
    # button.grid(row=len(paths) + 1, column=0, columnspan=3, pady=10)
    # progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
    # progress_bar.grid(row=len(paths) + 2, column=0, columnspan=3, pady=10)
    # progress_bar['value'] = 100
    # root.mainloop()
    generate_results()
    
