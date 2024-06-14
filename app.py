import PyPDF2
import pandas as pd 
import csv
import os 
from pdf2image import convert_from_path
from PIL import Image
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r".\Tesseract-OCR\tesseract.exe"


from invoice import Invoice
from extractor import PdfExtractor
from pick_list import PickList


class SwanSeaPrintingApp:
    """
    This is the main UI class for this application.
    """
    # this is the dict to store the relevant paths selected by the user
    paths = {
        'SKU FILE': None,
        'PACK FILE': None,
        'POST FILE': None,
        'IMAGE FOLDER': None,
        'OUTPUT IMAGE FOLDER': None,
        'OUTPUT FOLDER': None,
        'SHARED FOLDER': None,
    }

    # this is the dummy information for the testing purposes only 
    paths = {
        'SKU FILE': r"data/SKU LIST 2024.csv",
        # 'PACK FILE': [r"samples/29.2.24 - 8 Orders Pack 10001.pdf", r"samples/29.2.24 - 24 Orders Pack 10001.pdf"],
        # 'PACK FILE': [r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 23 Orders Pack 10001.pdf', r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 13 Orders Pack 10001.pdf', r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 1 Order Pack Claire 10002.pdf', r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 1 Orders Pack 10001.pdf'],
        # 'POST FILE': [r"samples/29.2.24 - 8 Orders Post 10001.pdf", r"samples/29.2.24 - 24 Orders Post 10001.pdf"],
        # 'POST FILE': [r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 1 Order Post Claire 10002.pdf', r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 20 Orders Post Claire 10002.pdf', r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 13 Orders Post 10001.pdf', r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 1 Orders Post 10001.pdf', r'C:/Users/Aamir Ali/Downloads/new orders/7.4.24 - 23 Orders Post 10001.pdf'],
        'PACK FILE': [r'C:/Users/Aamir Ali/Downloads/3.4.24 - 13 Orders Pack Claire 10002.pdf'],
        'POST FILE': [r'C:/Users/Aamir Ali/Downloads/3.4.24 - 13 Orders Post Claire 10002.pdf'],
        'IMAGE FOLDER': r"samples/PNG",
        'OUTPUT IMAGE FOLDER': r"samples/hbl",
        'OUTPUT FOLDER': r"samples/output",
        'SHARED FOLDER': r"samples/shared",
    }

    # this contains tkinter input objects
    entries = {
        'SKU FILE': None,
        'PACK FILE': None,
        'POST FILE': None,
        'IMAGE FOLDER': None,
        'OUTPUT IMAGE FOLDER': None,
        'OUTPUT FOLDER': None,
        'SHARED FOLDER': None,
    }

    # below lines segregates paths into files and folders for tkinter selection
    files = ['SKU FILE']
    multi_files = ['PACK FILE', 'POST FILE']
    folders = ['IMAGE FOLDER', 'OUTPUT IMAGE FOLDER', 'OUTPUT FOLDER', 'SHARED FOLDER']

    def __init__(self, root):
        """
        It initializes the UI and draws necessary widgets
        Args:
            root (tk.Tk): parent object
        """
        self.root = root
        root.title("Data Extraction and Processing")
        self.entries = {}
        
        # draw input widgets for files and folders
        self.draw_widgets(self.files)
        self.draw_widgets(self.multi_files, j=len(self.files), entry_type='multi-file')
        self.draw_widgets(self.folders, j=len(self.files + self.multi_files), entry_type='folder')

        # draw button to generate results
        button = tk.Button(root, text="Generate Results", command=self.handle_click)
        button.grid(row=len(self.paths) + 1, column=0, columnspan=3, pady=10)
        
        # finally generate a progress bar to show progress as it could take plenty of time to process files.
        progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
        progress_bar.grid(row=len(self.paths) + 2, column=0, columnspan=3, pady=10)
        progress_bar['value'] = 100

        self.progress_bar = progress_bar
        self.worker_thread = None

    def draw_widgets(self, keys, j=0, entry_type='file'):
        """
        It draws input widgets that include a text entry and the browse button
        Args:
            keys (list): list of keys to be drawn
            j (int): the starting index to draw on UI
            entry_type (str): 'file' or 'folder'
        """
        for i, key in enumerate(keys):
            label = tk.Label(self.root, text=key)
            label.grid(row=i+j, column=0, padx=10, pady=5, sticky="w")

            entry = tk.Entry(self.root, width=50, font=("Arial", 13), readonlybackground="white", state='readonly')
            entry.grid(row=i+j, column=1, padx=0, pady=10)

            button = tk.Button(self.root, text="Browse", width=10 ,command=lambda idx=key: self.browse_file(idx, entry_type))
            button.grid(row=i+j, column=2, padx=0, pady=10)

            self.entries[key] = entry

    def handle_click(self):
        """
        It creates a new worker if the initial worker is either done processing or 
        no worker is processing yet. it is inplace to handle multi-click mistakes
        """
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showerror('Task In Progress', 'Another task is already running')
            return
        self.worker_thread = threading.Thread(target=self.generate_results, daemon=True)
        self.worker_thread.start()


    def check_paths(self):
        """
        It checks if all the paths has been provided by the user or not 
        """
        for key, value in self.paths.items():
            if not value:
                messagebox.showerror("Path Not Selected", f"{key} is not selected. please select all paths to proceed")
                return False
        return True
    
    def browse_file(self, index, entry_type='file'):
        """
        it uses tkinter browse file dialog. we created seperate functions because 
        we are gonna select multiple files along the way
        Args:
            index (str): this is the name of the path 
            entry_type (str): 'file' or 'folder'
        """
        entries = self.entries
        if entry_type == 'file':
            file_path = filedialog.askopenfilename() 
        elif entry_type == 'folder':
            file_path = filedialog.askdirectory()
        elif entry_type == 'multi-file':
            file_path = filedialog.askopenfilenames()
        if file_path:
            self.paths[index] = file_path
            entries[index].config(state='normal')
            entries[index].delete(0, tk.END)
            entries[index].insert(0, file_path)
            entries[index].config(state='readonly')
    
    def filter_labels(self, labels, labels_path):
        """
        Labels pdf can contain both postage and custom labels and this function removes custom labels
        Args:
            labels ([Image]): list to be filtered
            labels_path (str): path to the labels pdf
        """
        reader = PyPDF2.PdfReader(labels_path)
        filtered_labels = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if 'CUSTOMS DECLARATION' not in text:
                filtered_labels.append(labels[i])

        return filtered_labels
    
    def name_labels(self, labels):
        
        named_labels = {}
        ind_named_labels = {}
        roi_coordinates = (500, 0, 800, 150)
        evri_roi = (25, 590, 250, 800)
        mail_roi = (25, 540, 500, 800)
        for label in labels:
            try:
                roi_image = label.crop(roi_coordinates)
                text = pytesseract.image_to_string(roi_image)
                if text[:4] == 'EVRi':
                    roi_image = label.crop(evri_roi)
                elif text[:12] == 'Delivered by':
                    roi_image = label.crop(mail_roi)
                text = pytesseract.image_to_string(roi_image)
                name = text.split('\n')[0].strip()
                named_labels[name] = label
                ind_names = [x.strip() for x in name.split(' ')]
                ind_named_labels[tuple(ind_names)] = label
            except:
                pass
        
        return named_labels, ind_named_labels
    
    def get_sku_details(self):
        """
        read sku details from the csv file and return the details 
        """
        sku_details = {}

        with open(self.paths['SKU FILE'], 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                sku = row['SKU']
                details = {key: value for key, value in row.items() if key != 'SKU'}
                sku_details[sku.lower()] = details

        return sku_details

    def get_customer_ids(self, pack_files):
        ids = set()
        for file in pack_files:
            customer_id = os.path.splitext(os.path.basename(file))[0].split(' ')[-1]
            ids.add(customer_id)
        return list(ids)
    
    def writelines(self, file, lines):
        for line in lines:
            file.write(f'{line}\n')

    def print_exceptions(self, writer):
        paths = self.paths
        if len(writer.pngs_not_found) > 0:
            with open(paths['OUTPUT FOLDER'] + '/images not found.txt', 'w') as f:
                self.writelines(f, writer.pngs_not_found)
        if len(writer.exceptions) > 0:
            with open(paths['OUTPUT FOLDER'] + '/errors.txt', 'w') as f:
                self.writelines(f, writer.exceptions)
        if len(writer.skus_not_found) > 0:
            with open(paths['OUTPUT FOLDER'] + '/skus not found.txt', 'w') as f:
                self.writelines(f, writer.skus_not_found)
        if len(writer.labels_not_found) > 0:
            with open(paths['OUTPUT FOLDER'] + '/labels not found.txt', 'w') as f:
                self.writelines(f, writer.labels_not_found)

    def generate_results(self):
        """
        this is the main function that controls the processing of the results
        """
        paths = self.paths
        try:
            self.progress_bar['value'] = 5
            if not self.check_paths():
                self.progress_bar['value'] = 100
                return
            
            os.makedirs(paths['OUTPUT FOLDER'], exist_ok=True)
            labels = []
            for filepath in paths['POST FILE']:
                try:
                    llabels = convert_from_path(filepath, poppler_path='poppler/bin')
                    llabels = self.filter_labels(llabels, filepath)
                    labels.extend(llabels)
                except Exception as e:
                    continue
            self.progress_bar['value'] = 15
            named_labels, ind_named_labels = self.name_labels(labels)
            custom = Image.open('data/customs_label.png')
            sku_details = self.get_sku_details()
            writer = PdfExtractor(paths['PACK FILE'], named_labels, ind_named_labels, custom, paths['IMAGE FOLDER'],
                                            paths['OUTPUT IMAGE FOLDER'], paths['SHARED FOLDER'], sku_details, self.progress_bar)
                            
            with open(paths['OUTPUT FOLDER'] + '/output.pdf', 'wb') as f:
                writer.write(f)
            self.print_exceptions(writer)
            df = pd.DataFrame(writer.garment_pick_list)
            # grouped_df = df.groupby(['name', 'size', 'color'])['quantity'].sum().reset_index().sort_values(by=['name', 'size', 'color'])
            # grouped_df = df.groupby(['name', 'size', 'color']).agg({'quantity': 'sum', 'Sort Key': 'first'}).reset_index()
            grouped_df = df.groupby(['Sort Key']).agg({'quantity': 'sum', 'name': 'first', 'size': 'first', 'color': 'first'}).reset_index()
            grouped_df = grouped_df.sort_values(by=['Sort Key']).reset_index().drop(columns=['index'])
            # grouped_df = grouped_df.sort_values(by=['Sort Key']).drop(columns=['Sort Key'])
            PickList(grouped_df).to_pdf(paths['OUTPUT FOLDER'] + '/pick_list.pdf')
            for cid, group_df in df.groupby(['customer id']):
                customer_id = cid[0]
                new_df = pd.DataFrame(group_df)
                grouped_df = new_df.groupby(['SKU TYPE'])['quantity'].sum().reset_index()
                try:
                    Invoice(grouped_df, customer_id).to_pdf(paths['OUTPUT FOLDER'] + f'/invoice_{customer_id}.pdf')
                except Exception as e:
                    messagebox.showerror("Error", e)
            
            self.progress_bar['value'] = 100
            messagebox.showinfo("Operation Completed", f"Output Files has been generated successfully")
        
        except Exception as e:
            # import traceback
            # traceback.print_exc()
            messagebox.showerror("Unknow Error", f"Some Unknow Error Occured. Find Details Below \n{e}")
            self.progress_bar['value'] = 100
        
    def get_shared_storage(self):
        return self.paths['SHARED FOLDER']
        
        



if __name__ == "__main__":
    root = tk.Tk()
    SwanSeaPrintingApp(root)
    root.mainloop()
    
