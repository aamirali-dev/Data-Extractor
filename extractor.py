import re 
import PyPDF2
from PIL import Image
from page import PdfPage

class PageExtractor:
    """
    This class extracts the desired information from a single page, we made this class because the template is same and we gotta extract same information from multiple pages.
    """
    
    # this is the list of keys to extract and the regular expressions that are used to extract those keys. 
    key_to_expressions = {
        'address': r'Deliver to\s+(.*?)\s+Scheduled to',
        'dispatch_date': r'Scheduled to dispatch by\s+(.*?)\s+Shop',
        'shop_name': r'Shop\s+(.*?)\s+Order date',
        'order_date': r'Order date\s+(.*?)\n',
        'no_of_items': r'(\d+)\s+(item|items)'
    }
    
    # this does the same thing, maps keys to regular expressions that are used to extract those keys. the only difference is that these keys could match multiple values in the same pdf depending on the no_of_items extracted in previous dictionary
    item_keys_to_expressions = {
        'SKU': r'SKU:\s+(.*?)\n',
        'Quantity': r'Colour: .+?(\d+) x ',
        'Design Code': r'\s+-\s+(\d+)\s+SKU:',
        'Title': r'(?:items?\s+|Colour:[^\n]*\n)(.*?)\s+SKU:'
        # 'Title': r'items?\s+(.*?)\s+SKU:'
    }
    
    # this is used to set and pass the number of of multi-item orders across different pages. moc stands for multiple order count.
    config = {'moc': 0}
    
    def __init__(self, page_text, SKU_DETAILS):
        """
        Args:
            page_text (str): The text of the page to extract information from.
            SKU_DETAILS (dict): A dictionary of SKU information, including any additional information needed to process the page.
        """
        self.page_text = page_text
        # SKU_DETAILS is just some business specific details and doesn't concern the program logic much neither requires understanding of the details
        self.SKU_DETAILS = SKU_DETAILS
        self.info = {}
        self.items = []
        self.count = 0
        self.extract_metadata()
        self.extract_items()
        self.assign_design_folder()

        
    def extract_metadata(self):
        """
        This function utilizes key_to_expressions dict to extract and store coresponding values in the self.info
        """      
        for key, regex in self.key_to_expressions.items():
            match = re.search(regex, self.page_text, re.DOTALL)
            if match:
                self.info[key] = match.group(1).strip()

        self.count = int(self.info['no_of_items'])

    def extract_items(self):
        """
        This function utilizes item_key_to_expressions dict to extract and store coresponding values in the self.info['items'] list. it also updates the item info from the SKU details as all info is not extracted from pdf.
        """    
        page_text = self.page_text
        info = self.info
        SKU_DETAILS = self.SKU_DETAILS
        
        # if order is multi-item, update moc 
        if self.count > 1:
            self.config['moc'] += 1
        
        # find matches for all the keys in item_keys_to_expressions
        items_info = {key: re.findall(expression, page_text, re.DOTALL) for key, expression in self.item_keys_to_expressions.items()}
        
        # since we have a dict of lists, we transform it to a list of dicts
        items = [{key: items_info[key][i] for key in items_info} for i in range(self.count)]
        
        for i, item in enumerate(items):
            # shorten the title by excluding some details
            item['Title'] = item['Title'].split('T-Shirt')[0]
            
            # extract & update some data from SKU details file
            if item['SKU'] in self.SKU_DETAILS:
                item.update(self.SKU_DETAILS[item['SKU']])
                
                # this logic was specified in business requirements. if there are multiple items, the file rename rule is different but for a single item, just pick the name from sku details.
                if self.count > 1:
                    item['Rename'] = f'4.{self.config["moc"]}.{i+1}.'
                else:
                    item['Rename'] = SKU_DETAILS[item['SKU']]['PDF PNG Rename (Add Seq(1.,2.,3.etc)'] + '1'
        
        info['items'] = items
        self.items = items
        
    def assign_design_folder(self):
        """
        Design folder is based on no of items and extracted from sku details. 
        """    
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
    """
    This class extracts information from the pdf. pdf will contain multiple pages but each page will have similar information. it uses PageExtractor to extract information from each page.
    """
    
    def __init__(self, reader, labels, custom, image_folder, target_image_folder, sku_details, progress_bar):
        
        """
        Initialize an instance of the PdfExtractor class.

        Args:
            reader (PdfReader): reader object to read pages from
            labels ([Image]): a list of postage labels to be used.
            custom (Image): a custom image to be used for every page 
            image_folder (str): folder to retrieve images from 
            target_image_folder (str): folder to save images to
            sku_details (dict): a dictionary of SKU information
            progress_bar (Progressbar): tkinter progress bar object, I know this is tightly coupled approach but business requirements aren't changing much.
        """
        
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
        """
        Garment pick list is just the list of items ordered. this function just extract some desired information for pick list from the item and append it. we are not checking for duplicates as this is the desired behavior.
        """
        for item in page['items']:
            self.garment_pick_list.append(
                {
                    'name': item['Garment Type'], 
                    'size': item['Size'], 
                    'color': item['Colour'], 
                    'quantity': int(item['Quantity']), 
                    'SKU TYPE': item['SKU'].split('-')[1],
                }
            )
 


    def process_files(self):
        """
        it does 3 tasks:
        1. extracts information from each page
        2. creates a resulting pdf page
        3. saves the information for each page in self.info 
        """
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
        """
        Sort files using the Sort Key
        """
        self.info = sorted(self.info, key=lambda page: page['Sort Key'])

    def write(self, filename):
        """
        Writes the PDF file to the specified filename.

        Args:
            filename (str): The name of the file to write to.

        Returns:
            None
        """
        writer = PyPDF2.PdfWriter()
        for entry in self.info:
            writer.add_page(entry['page'])
        writer.write(filename)
    
    def get_image_not_found(self):
        """
        It returns the list of images that were not found. we chose to proceed with missing images and print their names so that we can make them available
        """
        return PdfPage.PNGS_NOT_FOUND
    
    


