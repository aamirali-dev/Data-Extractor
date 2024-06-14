import io, os 
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from PIL import Image
import PyPDF2
import qrcode
import uuid

class PdfPage:

    def __init__(self, data, label, custom, image_folder, image_output_folder, shared_storage):
        """
        Initialize the PDF page object.

        Args:
            data (dict): The data to be used for generating the PDF.
            label (Image): The postage label image.
            custom (Image): The custom label image.
            image_folder (str): The folder containing the product images.
            image_output_folder (str): The folder where the images will be saved for printing.
        """
        self.data = data
        self.label = label
        self.custom = custom
        self.image_folder = image_folder
        self.image_output_folder = image_output_folder
        self.shared_storage = shared_storage
        self.pngs_not_found = []
        os.makedirs(self.shared_storage, exist_ok=True)
    
    def init_canvas(self, width=210, height=300):
        """
        Initialize the canvas for pdf generation as we are gonna draw everyting 
        manually and aren't gonna use any flowable or pdf template.
        """
        self.packet = io.BytesIO()
        # width, height = letter
        width *= 2.83465
        height *= 2.83465
        self.canvas = canvas.Canvas(self.packet, pagesize=(width, height))
        
    def draw_items(self):
        """
        Draws each item from the order on the canvas. drawing item means drawing QTY x Size, Colour, Type, & Design code. Thumbnail is also included.
        """
        y = 800
        for i, item in enumerate(self.data['items']):
            # Draw each piece of text separately
            self.canvas.setFont("Helvetica-Bold", 25)
            qty = item.get('Quantity', '')
            size = item.get('Size', '')
            self.canvas.drawString(30, y - (i*60), f'{qty} x {size}')
            self.canvas.drawString(130, y - (i*60), item.get('Colour', ''))
            self.canvas.drawString(250, y - (i*60), item.get('Garment Type', ''))
            self.canvas.drawString(390, y - (i*60), item.get('Design Code', ''))
            self.canvas.setFont("Helvetica-Bold", 13)
            title, title2 = self.create_title(item.get('Title', ''))
            self.canvas.drawString(30, y - 20 - (i*60), title)
            if title2:
                self.canvas.drawString(30, y - 35 - (i*60), title2)
            try:
                self.draw_thumbnail(item, y, i)
            except:
                pass 
            
    def draw_thumbnail(self, item, y, i):
        """
        It performs 3 tasks
        1. copy thumnail image to the target location & rename it for sorting.
        2. it sets the background color to grey
        3. finally, it draws the thumbnail image on canvas
        """
        image_path = self.image_folder + f"/{item['Design Code']}.png"
        targe_image_folder = self.image_output_folder + f"/{self.data['Design Folder']}"
        
        try:
            image = Image.open(image_path)
            img = image.resize((55, 55), Image.LANCZOS)
            background_color = (192, 192, 192)  # Grey color
            background = Image.new('RGB', img.size, background_color)
            background.paste(img, (0, 0), img)
            self.canvas.drawInlineImage(background, 520, y - (i*60) - 20)
            os.makedirs(targe_image_folder, exist_ok=True)
            if int(item['Quantity']) > 1:
                for i in range(int(item['Quantity'])):
                    target_image_path = self.image_output_folder + f"/{self.data['Design Folder']}/{item['Rename']}.{i+1}.png"
                    image.save(target_image_path)
            else:
                target_image_path = self.image_output_folder + f"/{self.data['Design Folder']}/{item['Rename']}.png"
                image.save(target_image_path)
        except:
            self.pngs_not_found.append(image_path)
    
    def draw_total_items(self):
        y = 500
        self.canvas.setFont("Helvetica-Bold", 25)
        self.canvas.drawString(30, y, f'TOTAL = {self.data["no_of_items"]} Items')
        self.canvas.setFont("Helvetica", 12)

    def draw_order_details(self):
        """
        Adds order details such as total amount, store, address to send this order to, including few dates.
        """
        
        index = 0
        
        def increment_index(i=1):
            nonlocal index
            index += i
        
        y = 470
        for i, line in enumerate(self.data['address'].splitlines()):
            self.canvas.drawString(30, y - (index * 15), line) 
            increment_index()
        increment_index()
        self.canvas.drawString(30, y - (index * 15), 'Order Date:')
        increment_index()
        self.canvas.drawString(30, y - (index * 15), self.data.get('order_date', ''))
        increment_index(2)
        self.canvas.drawString(30, y - (index * 15), 'Dispatch Date:')
        increment_index()
        self.canvas.drawString(30, y - (index * 15), self.data.get('dispatch_date', ''))
        increment_index(2)
        self.canvas.drawString(30, y - (index * 15), self.data.get('shop_name', '')) 

    def draw_labels(self):
        """
        Draws both postage & custom labels.
        """
        letter_width, letter_height = letter
        width, height, x, y = 100 * 2.83465, 150 * 2.83465, letter_width - (115 * 2.83465), (10 * 2.83465)
        self.canvas.drawInlineImage(self.label, x, y, width=width, height=height)
        width, height, x, y = 80 * 2.83465, 80 * 2.83465, letter_width - (200 * 2.83465), (7.5 * 2.83465)
        self.canvas.drawInlineImage(self.custom, x, y, width=width, height=height) 

    def get(self):
        """
        creates the pdf from the canvas and returns
        
        Returns:
            PdfReader
        """
        filepath = self.save_qr_file()
        self.init_canvas()
        self.draw_items()
        self.draw_total_items()
        self.draw_order_details()
        self.draw_labels()
        self.draw_qr_code(filepath)
        self.canvas.save()
        self.packet.seek(0)
        return PyPDF2.PdfReader(self.packet), self.pngs_not_found

    def save_qr_file(self):
        """
        Save the current state of the canvas to a PDF file.
        
        Args:
            filename (str): The filename to save the PDF.
        """
        self.init_canvas(height=125)
        self.canvas.translate(0, -(175 * 2.83465))
        self.draw_items()
        self.draw_total_items()
        filename = self.shared_storage + f"/{uuid.uuid4()}.pdf"
        self.canvas.save()
        self.packet.seek(0)
        with open(filename, "wb") as f:
            f.write(self.packet.getvalue())
        
        return filename
    
    def draw_qr_code(self, filepath):
        code = qrcode.QRCode(
            version=1,
            error_correction=qrcode.ERROR_CORRECT_L,
            box_size=10,
            border=1
        ) 
        code.add_data(filepath)
        code.make(fit=True)
        img = code.make_image(fill_color='black', back_color='white')
        img_bytes = io.BytesIO()
        img.save(img_bytes, format='PNG')
        letter_width, letter_height = letter
        width, height, x, y = 30 * 2.83465, 30 * 2.83465, letter_width - (150 * 2.83465), (90 * 2.83465)
        self.canvas.drawImage(ImageReader(img_bytes), x, y, width=width, height=height)

    def create_title(self, title, max_characters=60):
        """
        It splits the title at newline and returns the title as a list of 2 objects. if the title is small enough, second object is None.
        """
        title = title.replace('\n', ' ')
        if len(title) > max_characters:
            start = max_characters
            while title[start] != ' ':
                start -= 1
            return [title[:start], self.create_title(title[start+1:])[0]]

        return [title, None]