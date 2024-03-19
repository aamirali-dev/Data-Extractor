import re 
import PyPDF2
from PIL import Image

class PageExtractor:
    
    key_to_expressions = {
        'address': r'Deliver to\s+(.*?)\s+Scheduled to',
        'dispatch_date': r'Scheduled to dispatch by\s+(.*?)\s+Shop',
        'shop_name': r'Shop\s+(.*?)\s+Order date',
        'order_date': r'Order date\s+(.*?)\n',
        'no_of_items': r'(\d+)\s+(item|items)'
    }
    
    item_keys_to_expressions = {
        'SKU': r'SKU:\s+(.*?)\n',
        'Quantity': r'Colour: .+?(\d+) x ',
        'Design Code': r'\s+-\s+(\d+)\s+SKU:',
        'Title': r'(.+?\n.+?)\s+SKU:'
    }
    
    config = {'moc': 0}
    
    def __init__(self, page_text, SKU_DETAILS):
        self.page_text = page_text
        self.SKU_DETAILS = SKU_DETAILS
        self.info = {}
        self.items = []
        self.count = 0
        self.extract_metadata()
        self.extract_items()
        self.assign_design_folder()
        
    def extract_metadata(self):        
        for key, regex in self.key_to_expressions.items():
            match = re.search(regex, self.page_text, re.DOTALL)
            if match:
                self.info[key] = match.group(1).strip()

    def extract_items(self):
        page_text = self.page_text
        info = self.info
        SKU_DETAILS = self.SKU_DETAILS
        
        self.count = info['no_of_items']
        if info['no_of_items'] > 1:
            self.config['moc'] += 1
        
        items_info = {key: re.findall(expression, page_text, re.DOTALL) for key, expression in self.item_keys_to_expressions.items()}
            
        items = [{key: items_info[key][i] for key in items_info} for i in range(info['no_of_items'])]
        
        for i, item in enumerate(items):
            item['Title'] = item['Title'].split('T-Shirt')[0]
            
            if item['SKU'] in self.SKU_DETAILS:
                item.update(self.SKU_DETAILS[item['SKU']])
                
                if info['no_of_items'] > 1:
                    item['Rename'] = f'4.{self.config["moc"]}.{i+1}.'
                else:
                    item['Rename'] = SKU_DETAILS[item['SKU']]['PDF PNG Rename (Add Seq(1.,2.,3.etc)'] + '1'
        
        info['items'] = items
        self.items = items
        
    def assign_design_folder(self):
        items = self.info['items']
        
        if len(items) > 1:
            self.info['Design Folder'] = '4. Multi Orders'
        elif len(items) == 1:
            self.info['Design Folder'] = self.SKU_DETAILS[items[0]['SKU']]['Design Folder'] 

    def info(self):
        return self.info
    

class PdfExtractor:
    def __init__(self, reader, labels, image_folder, target_image_folder, sku_details, update_progress):
        
        self.reader = reader
        self.labels = labels
        self.image_folder = image_folder
        self.target_image_folder = target_image_folder
        self.sku_details = sku_details 
        self.update_progress = update_progress
        self.writer = PyPDF2.PdfWriter()
        self.num_pages = len(reader.pages)

def extract_info_from_pdf(pdf_path, labeller_path, image_folder, image_output_folder, sku_file_path):
    global progress_bar
    all_info = []
    try:
        SKU_DETAILS = read_csv_to_dict(sku_file_path)
    except:
        messagebox.showerror("SKU File Error", "Error Reading SKU File, either file is not there or it is not in the right format")
        return
    with open(pdf_path, 'rb') as file, open(labeller_path, 'rb') as file2:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()
        num_pages = len(reader.pages)
        labeller = convert_from_path(labeller_path, poppler_path='poppler/bin')
        custom = Image.open('data/customs_label.png')
        label_map = create_label_map(labeller_path)
        progress_bar['value'] = 20
        page_value = 70 / num_pages
        for page_number in range(num_pages):
            progress_bar['value'] += page_value
            page = reader.pages[page_number]
            text = page.extract_text()
            post = labeller[label_map[page_number]['post']]
            # custom = labeller[label_map[page_number]['custom']] if label_map[page_number]['custom'] else None
            print('page no ', str(page_number))
            # for line in text.splitlines():
            #     print(line)
            page_info = PageExtractor(text, SKU_DETAILS).info()
            new_pdf_page = PdfPage(page_info, post, custom, image_folder, image_output_folder).get()
            add_to_garment_pick_list(page_info)
            writer.add_page(new_pdf_page.pages[0])
            all_info.append(new_pdf_page)

    return writer


