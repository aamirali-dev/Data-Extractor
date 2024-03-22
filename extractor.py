import re 
import PyPDF2
from PIL import Image
from page import PdfPage

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
        'Title': r'(?:items?\s+|Colour:[^\n]*\n)(.*?)\s+SKU:'
        # 'Title': r'items?\s+(.*?)\s+SKU:'
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

        self.count = int(self.info['no_of_items'])

    def extract_items(self):
        page_text = self.page_text
        info = self.info
        SKU_DETAILS = self.SKU_DETAILS
        
        
        if self.count > 1:
            self.config['moc'] += 1
        
        items_info = {key: re.findall(expression, page_text, re.DOTALL) for key, expression in self.item_keys_to_expressions.items()}
        print(items_info)
        print(page_text)
        items = [{key: items_info[key][i] for key in items_info} for i in range(self.count)]
        
        for i, item in enumerate(items):
            item['Title'] = item['Title'].split('T-Shirt')[0]
            
            if item['SKU'] in self.SKU_DETAILS:
                item.update(self.SKU_DETAILS[item['SKU']])
                
                if self.count > 1:
                    item['Rename'] = f'4.{self.config["moc"]}.{i+1}.'
                else:
                    item['Rename'] = SKU_DETAILS[item['SKU']]['PDF PNG Rename (Add Seq(1.,2.,3.etc)'] + '1'
        
        info['items'] = items
        self.items = items
        
    def assign_design_folder(self):
        items = self.info['items']
        
        if len(items) > 1:
            self.info['Design Folder'] = '4. Multi Orders'
            self.info['Sort Key'] = items[0]['Rename']
        elif len(items) == 1:
            self.info['Design Folder'] = self.SKU_DETAILS[items[0]['SKU']]['Design Folder']
            self.info['Sort Key'] = items[0]['Rename']

    def get_info(self):
        return self.info
    

class PdfExtractor:
    def __init__(self, reader, labels, custom, image_folder, target_image_folder, sku_details, progress_bar):
        
        self.reader = reader
        self.labels = labels
        self.custom = custom
        self.image_folder = image_folder
        self.target_image_folder = target_image_folder
        self.sku_details = sku_details 
        self.progress_bar = progress_bar
        self.writer = PyPDF2.PdfWriter()
        self.num_pages = len(reader.pages)
        self.garment_pick_list = []
        self.info = []
        self.images_not_found = []

        self.process_files()
        self.sort_files()

    def add_to_pick_list(self, page):
        for item in page['items']:
            self.garment_pick_list.append({
                'name': item['Garment Type'], 
                'size': item['Size'], 
                'color': item['Colour'], 
                'quantity': int(item['Quantity']), 
                'SKU TYPE': item['SKU'].split('-')[1],
                'Sort Key': item['Rename']
                })
 


    def process_files(self):

        self.progress_bar['value'] = 20
        page_value = 70 / self.num_pages
        for page_number in range(self.num_pages):
            self.progress_bar['value'] += page_value
            page = self.reader.pages[page_number]
            text = page.extract_text()
            post = self.labels[page_number]
            page_info = PageExtractor(text, self.sku_details).get_info()
            new_pdf_page = PdfPage(page_info, post, self.custom, self.image_folder, self.target_image_folder).get()
            self.add_to_pick_list(page_info)
            self.writer.add_page(new_pdf_page.pages[0])
            self.info.append({'data': page_info, 'page': new_pdf_page.pages[0], 'Sort Key': page_info['Sort Key']})
    
    def sort_files(self):
        self.info = sorted(self.info, key=lambda page: page['Sort Key'])

    def write(self, f):
        writer = PyPDF2.PdfWriter()
        for entry in self.info:
            writer.add_page(entry['page'])
        writer.write(f)
    
    def get_image_not_found(self):
        return PdfPage.PNGS_NOT_FOUND
    
    


