import io, os 
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image
import PyPDF2

PNGS_NOT_FOUND = []

class PdfPage:
    def __init__(self, data, label, custom, image_folder, image_output_folder):
        self.data = data
        self.label = label
        self.custom = custom
        self.image_folder = image_folder
        self.image_output_folder = image_output_folder
        self.init_canvas()
        self.draw_items()
        self.draw_order_details()
        self.draw_labels()
    
    def init_canvas(self):
        self.packet = io.BytesIO()
        width, height = letter
        width = 210 * 2.83465
        height = 300 * 2.83465
        self.canvas = canvas.Canvas(self.packet, pagesize=(width, height))
        
    def draw_items(self):
        y = 800
        for i, item in enumerate(self.data['items']):
            # Draw each piece of text separately
            self.canvas.setFont("Helvetica-Bold", 25)
            qty = item['Quantity']
            size = item['Size']
            self.canvas.drawString(30, y - (i*60), f'{qty} x {size}')
            self.canvas.drawString(130, y - (i*60), item['Colour'])
            self.canvas.drawString(250, y - (i*60), item['Garment Type'])
            self.canvas.drawString(390, y - (i*60), item['Design Code'])
            self.canvas.setFont("Helvetica-Bold", 13)
            title, title2 = self.create_title(item['Title'])
            self.canvas.drawString(30, y - 20 - (i*60), title)
            if title2:
                self.canvas.drawString(30, y - 35 - (i*60), title2)
            
            self.draw_thumbnail(item, y, i)
            
    def draw_thumbnail(self, item, y, i):
        image_path = self.image_folder + f"/{item['Design Code']}.png"
        targe_image_folder = self.image_output_folder + f"/{item['Design Folder']}"
        target_image_path = self.image_output_folder + f"/{item['Design Folder']}/{item['Rename']}.png"
        try:
            image = Image.open(image_path)
            img = image.resize((55, 55), Image.LANCZOS)
            self.canvas.drawInlineImage(img, 520, y - (i*60) - 20)
            os.makedirs(targe_image_folder, exist_ok=True)
            image.save(target_image_path)
        except:
            PNGS_NOT_FOUND.append(image_path)
    
    def draw_order_details(self):
        y = 500
        index = 0
        
        def increment_index(i=1):
            index += i
        
        self.canvas.setFont("Helvetica-Bold", 25)
        self.canvas.drawString(30, y, f'TOTAL = {self.data["no_of_items"]} Items')
        self.canvas.setFont("Helvetica", 12)
        
        y = 470
        for i, line in enumerate(self.data['address'].splitlines()):
            self.canvas.drawString(30, y - (index * 15), line) 
            increment_index()
        increment_index()
        self.canvas.drawString(30, y - (index * 15), 'Order Date:')
        increment_index()
        self.canvas.drawString(30, y - (index * 15), self.data['order_date'])
        increment_index(2)
        self.canvas.drawString(30, y - (index * 15), 'Dispatch Date:')
        increment_index()
        self.canvas.drawString(30, y - (index * 15), self.data['dispatch_date'])
        increment_index(2)
        self.canvas.drawString(30, y - (index * 15), self.data['shop_name']) 

    def draw_labels(self):
        letter_width, letter_height = letter
        width, height, x, y = 100 * 2.83465, 150 * 2.83465, letter_width - (115 * 2.83465), (10 * 2.83465)
        self.canvas.drawInlineImage(self.label, x, y, width=width, height=height)
        width, height, x, y = 80 * 2.83465, 80 * 2.83465, letter_width - (197 * 2.83465), (7.5 * 2.83465)
        self.canvas.drawInlineImage(self.custom, x, y, width=width, height=height) 

    def get(self):
        self.canvas.save()
        self.packet.seek(0)
        return PyPDF2.PdfReader(self.packet)

    def create_title(self, title, max_characters=40):
        title = title.replace('\n', ' ')
        if len(title) > max_characters:
            start = max_characters
            while title[start] != ' ':
                start -= 1
            return [title[:start], title[start+1:]]

        return [title, None]