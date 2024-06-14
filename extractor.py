import re 
import PyPDF2
from PIL import Image
from page import PdfPage
import PyPDF2
import os
from fuzzywuzzy import fuzz

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
    png_name_counts = {}
    
    def __init__(self, file, page_no, page_text, sku_details):
        """
        Args:
            page_text (str): The text of the page to extract information from.
            SKU_DETAILS (dict): A dictionary of SKU information, including any additional information needed to process the page.
        """
        self.filename = file
        self.page_no = page_no
        self.page_text = page_text
        # SKU_DETAILS is just some business specific details and doesn't concern the program logic much neither requires understanding of the details
        self.sku_details = sku_details
        self.info = {}
        self.items = []
        self.count = 0
        self.items_not_found = []
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
        try:
            self.info['name'] = self.info['address'].split('\n')[0].strip()
        except:
            raise Exception(f'Unable to find name on page: {self.page_no}, file: {self.filename}')

    def extract_items(self):
        """
        This function utilizes item_key_to_expressions dict to extract and store coresponding values in the self.info['items'] list. it also updates the item info from the SKU details as all info is not extracted from pdf.
        """    
        page_text = self.page_text
        info = self.info
        sku_details = self.sku_details
        
        # find matches for all the keys in item_keys_to_expressions
        items_info = {key: re.findall(expression, page_text, re.DOTALL) for key, expression in self.item_keys_to_expressions.items()}
        
        # since we have a dict of lists, we transform it to a list of dicts
        try:
            items = [{key: items_info[key][i] for key in items_info} for i in range(len(items_info['SKU']))]
        except:
            raise Exception(f'Unable to parse page: {self.page_no}, file: {self.filename}')
        items_not_found = self.items_not_found
        items_not_found_index = []

        # if order is multi-item, update moc 
        self.count = len(items)
        if int(self.info['no_of_items']) > 1:
            self.config['moc'] += 1

        for i, item in enumerate(items):
            # shorten the title by excluding some details
            item['Title'] = item['Title'].split('T-Shirt')[0]
            
            # extract & update some data from SKU details file
            sku = item['SKU'].lower()
            if sku in self.sku_details:
                item.update(self.sku_details[sku])
                
                # this logic was specified in business requirements. if there are multiple items, the file rename rule is different but for a single item, just pick the name from sku details.
                if int(self.info['no_of_items']) > 1:
                    item['Rename'] = f'4.{self.config["moc"]}.{i+1}.'
                else:
                    name = sku_details[sku]['PDF PNG Rename (Add Seq(1.,2.,3.etc)']
                    # item['Rename'] = sku_details[item['SKU']]['PDF PNG Rename (Add Seq(1.,2.,3.etc)'] + '1'
                    if name in self.png_name_counts.keys():
                        item['Rename'] = name + str(self.png_name_counts[name])
                        self.png_name_counts[name] += 1
                    else:
                        item['Rename'] = name + '1'
                        self.png_name_counts[name] = 2
                         
            else:
                items_not_found.append(item['SKU'])
                items_not_found_index.append(i)
        
        # remove items for which SKU is not found
        items_not_found_index.sort(reverse=True)
        for index in items_not_found_index:
            del items[index]

        info['items'] = items
        self.items = items
        
    def assign_design_folder(self):
        """
        Design folder is based on no of items and extracted from sku details. 
        """    
        items = self.info['items']
        try:
            if int(self.info['no_of_items']) > 1 and len(items) >= 1:
                self.info['Design Folder'] = '4. Multi Orders'
                self.info['Sort Key'] = items[0]['Rename']
            elif int(self.info['no_of_items']) == 1 and len(items) >= 1:
                self.info['Design Folder'] = self.sku_details[items[0]['SKU'].lower()]['Design Folder']
                self.info['Sort Key'] = items[0]['Rename']
        except Exception as e:
            raise Exception(f'Unable to assign design folder and sort key, page: {self.page_no}, file: {self.filename}')

    def get_info(self):
        return self.info, self.items_not_found
    

class PdfExtractor:
    """
    This class extracts information from the pdf. pdf will contain multiple pages but each page will have similar information. it uses PageExtractor to extract information from each page.
    """
    
    def __init__(self, files, labels, ind_named_labels, custom, image_folder, target_image_folder, shared_storage, sku_details, progress_bar):
        
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
        
        self.files = files
        self.labels = labels
        self.ind_named_labels = ind_named_labels
        self.custom = custom
        self.image_folder = image_folder
        self.target_image_folder = target_image_folder
        self.sku_details = sku_details 
        self.progress_bar = progress_bar
        self.shared_storage = shared_storage
        # self.writer = PyPDF2.PdfWriter()
        self.garment_pick_list = []
        self.info = []
        self.images_not_found = []
        self.exceptions = []
        self.skus_not_found = []
        self.pngs_not_found = []
        self.labels_not_found = []

        self.process_files()
        self.sort_files()

    def add_to_pick_list(self, page, customer_id):
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
        'Sort Key': ".".join(item['Rename'].split('.')[:3]),
                    'customer id': customer_id,
                }
            )
 
    def get_customer_id(self, filepath):
        return os.path.splitext(os.path.basename(filepath))[0].split(' ')[-1]

    def process_files(self):
        """
        it does 3 tasks:
        1. extracts information from each page
        2. creates a resulting pdf page
        3. saves the information for each page in self.info 
        """
        self.progress_bar['value'] = 20
        file_value = 70 / len(self.files)
        for file in self.files:
            reader = PyPDF2.PdfReader(file)
            page_value = file_value / len(reader.pages)

            try:
                customer_id = self.get_customer_id(file)
            except:
                self.exceptions.append(f'Unable to get customer id, file: {file}')
                continue
            for page_number in range(len(reader.pages)):
                self.progress_bar['value'] += page_value
                page = reader.pages[page_number]
                text = page.extract_text()

                try:
                    page_info, skus_not_found = PageExtractor(file, page_number, text, self.sku_details).get_info()
                except Exception as e:
                    self.exceptions.append(e)
                    continue

                if len(page_info['items']) < 1:
                    self.skus_not_found.extend(skus_not_found)
                    continue

                try:
                    post = self.get_label(page_info['name'])
                except:
                    self.labels_not_found.append(page_info['name'])
                    post = Image.new('RGB', (400, 400), (255, 255, 255))

                new_pdf_page, pngs_not_found = PdfPage(page_info, post, self.custom, self.image_folder, self.target_image_folder, self.shared_storage).get()
                try:
                    self.add_to_pick_list(page_info, customer_id)
                except Exception as e:
                    self.exceptions.append(f'{e} not found, page: {page_number}, file: {file}')
                    continue
                self.skus_not_found.extend(skus_not_found)
                self.pngs_not_found.extend(pngs_not_found)
                # self.writer.add_page(new_pdf_page.pages[0])
                self.info.append({'data': page_info, 'page': new_pdf_page.pages[0], 'Sort Key': page_info['Sort Key']})

    def approximate_match(self, name):
        max_score = -1
        best_match = None

        def remove_whitespace(string):
            return ''.join(string.split())

        for key in self.labels.keys():
            score = fuzz.partial_ratio(remove_whitespace(name), remove_whitespace(key))
            if score > max_score:
                max_score = score
                best_match = key

        return best_match, max_score



    def get_label(self, name):
        post = self.labels.get(name, None)
        if post is not None:
            return post
        ind_name = [x.strip() for x in name.split(' ')]
        ind_name = tuple(ind_name)
        post = self.ind_named_labels.get(ind_name, None)
        if post is not None:
            return post 
        key, score = self.approximate_match(name)
        post = self.labels[key]
        return post


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
        return self.pngs_not_found
    
    


