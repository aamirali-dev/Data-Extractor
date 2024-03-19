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


progress_bar = {}
WORKER_THREAD = None
PNGS_NOT_FOUND = []

def create_title(title, max_characters=40):
    title = title.replace('\n', ' ')
    if len(title) > max_characters:
        start = max_characters
        while title[start] != ' ':
            start -= 1
        return [title[:start], title[start+1:]]

    return [title, None] 

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

def create_pdf_page(data, label, custom, image_folder, image_output_folder):
    # Create a new PDF page
    packet = io.BytesIO()
    width, height = letter
    width = 210 * 2.83465
    height = 300 * 2.83465
    can = canvas.Canvas(packet, pagesize=(width, height))
    y = 800
    for i, item in enumerate(data['items']):
        # Draw each piece of text separately
        can.setFont("Helvetica-Bold", 25)
        qty = item['Quantity']
        size = item['Size']
        can.drawString(30, y - (i*60), f'{qty} x {size}')
        can.drawString(130, y - (i*60), item['Colour'])
        can.drawString(250, y - (i*60), item['Garment Type'])
        can.drawString(390, y - (i*60), item['Design Code'])
        can.setFont("Helvetica-Bold", 13)
        title, title2 = create_title(item['Title'])
        can.drawString(30, y - 20 - (i*60), title)
        if title2:
            can.drawString(30, y - 35 - (i*60), title2)
        image_path = image_folder + f"/{item['Design Code']}.png"
        targe_image_folder = image_output_folder + f"/{item['Design Folder']}"
        target_image_path = image_output_folder + f"/{item['Design Folder']}/{item['Rename']}.png"
        try:
            image = Image.open(image_path)
            img = image.resize((55, 55), Image.LANCZOS)
            can.drawInlineImage(img, 520, y - (i*60) - 20)
            os.makedirs(targe_image_folder, exist_ok=True)
            image.save(target_image_path)
        except:
            PNGS_NOT_FOUND.append(image_path) 

    can.setFont("Helvetica-Bold", 25)
    y = 500
    can.drawString(30, y, f'TOTAL = {data["no_of_items"]} Items')

    can.setFont("Helvetica", 12)
    index = 0
    y = 470
    for i, line in enumerate(data['address'].splitlines()):
        can.drawString(30, y - (index * 15), line) 
        index += 1
    index += 1
    can.drawString(30, y - (index * 15), 'Order Date:')
    index += 1
    can.drawString(30, y - (index * 15), data['order_date'])
    index += 2
    can.drawString(30, y - (index * 15), 'Dispatch Date:')
    index += 1
    can.drawString(30, y - (index * 15), data['dispatch_date'])
    index += 2
    can.drawString(30, y - (index * 15), data['shop_name'])
    letter_width, letter_height = letter
    width, height, x, y = 100 * 2.83465, 150 * 2.83465, letter_width - (115 * 2.83465), (10 * 2.83465)
    can.drawInlineImage(label, x, y, width=width, height=height)
    if custom:
        width, height, x, y = 80 * 2.83465, 80 * 2.83465, letter_width - (197 * 2.83465), (7.5 * 2.83465)
        can.drawInlineImage(custom, x, y, width=width, height=height)

    can.save()
    
    # Move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PyPDF2.PdfReader(packet)
    
    return new_pdf



config = {'moc': 0}
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


def extract_info_from_page(page_text, SKU_DETAILS):
    info = {}
    address_match = re.search(r'Deliver to\s+(.*?)\s+Scheduled to', page_text, re.DOTALL)
    if address_match:
        info['address'] = address_match.group(1).strip()
    
    dispatch_date_match = re.search(r'Scheduled to dispatch by\s+(.*?)\s+Shop', page_text, re.DOTALL)
    if dispatch_date_match:
        info['dispatch_date'] = dispatch_date_match.group(1).strip()
    
    shop_name_match = re.search(r'Shop\s+(.*?)\s+Order date', page_text, re.DOTALL)
    if shop_name_match:
        info['shop_name'] = shop_name_match.group(1).strip()
    
    order_date_match = re.search(r'Order date\s+(.*?)\n', page_text, re.DOTALL)
    if order_date_match:
        info['order_date'] = order_date_match.group(1).strip()
    
    items_match = re.search(r'(\d+)\s+(item|items)', page_text)
    if items_match:
        info['no_of_items'] = int(items_match.group(1))
    
    items = []
    sku_matches = re.findall(r'SKU:\s+(.*?)\n', page_text)
    # quantity_matches = re.findall(r'(\d+) x ' , page_text) // this is the simple yet working for now. next one is bit more specific
    quantity_matches = re.findall(r'Colour: .+?(\d+) x ' , page_text)
    design_code_matches = re.findall(r'\s+-\s+(\d+)\s+SKU:' , page_text)
    title_matches = re.findall(r'(.+?\n.+?)\s+SKU:' , page_text)
    # print(sku_matches)
    # print(quantity_matches)
    # print(design_code_matches)
    # print(title_matches)
    # print(title_matches)
    # for i in page_text.splitlines()[19]:
    #     print(i)
    if info['no_of_items'] > 1:
        config['moc'] += 1
    for i in range(info['no_of_items']):
        if i < len(sku_matches):
            sku = sku_matches[i]
            if sku in SKU_DETAILS:
                item_info = {'SKU': sku}
                item_info.update(SKU_DETAILS[sku])
                item_info['Quantity'] = quantity_matches[i]
                item_info['Design Code'] = design_code_matches[i]
                item_info['Title'] = title_matches[i].split('T-Shirt')[0]
                items.append(item_info)
                if info['no_of_items'] > 1:
                    item_info['Rename'] = f'4.{config["moc"]}.{i+1}.'
                else:
                    item_info['Rename'] = SKU_DETAILS[sku]['PDF PNG Rename (Add Seq(1.,2.,3.etc)'] + '1'
    info['items'] = items
    if len(items) > 1:
        info['Design Folder'] = '4. Multi Orders'
    elif len(items) == 1:
        info['Design Folder'] = SKU_DETAILS[items[0]['SKU']]['Design Folder']
    



    return info

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
            page_info = extract_info_from_page(text, SKU_DETAILS)
            new_pdf_page = create_pdf_page(page_info, post, custom, image_folder, image_output_folder)
            add_to_garment_pick_list(page_info)
            writer.add_page(new_pdf_page.pages[0])
            all_info.append(new_pdf_page)

    return writer

def export_to_pdf(df, filename):
    # doc = SimpleDocTemplate(filename, pagesize=letter)
    doc = BaseDocTemplate(filename)
    elements = [Table([list(df.columns)] + df.values.tolist())]
    style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)])

    elements[0].setStyle(style)
    frame1 = Frame(doc.leftMargin, doc.bottomMargin, doc.width/2-6, doc.height, id='col1')
    frame2 = Frame(doc.leftMargin+doc.width/2+6, doc.bottomMargin, doc.width/2-6, doc.height, id='col2')
    doc.addPageTemplates([PageTemplate(id='TwoCol',frames=[frame1,frame2]), ])
    doc.build(elements)

def read_csv(filename):
    try:
        with open(filename) as file:
            data = file.readlines()
        data = {line.strip().split(',')[0]: line.strip().split(',')[1] for line in data}
        return data 
    except:
        return None


def draw_on_canvas(c, doc, customer_details, order_details):
    # Open the existing PDF
    # c = canvas.Canvas('example.pdf')

    # customer_details = {
    #     'Customer ID': '10001',
    #     'Name': 'Jamie Durke',
    #     'Company Name': 'Swansea Designs',
    #     'Address': '888 Llangyfelach Rd',
    #     'City': 'Swansea',
    #     'Postcode': 'SA5 9AU',
    #     'Phone': '447367881134'
    # }

    c.drawImage('data/logo.png', 5, 740, width=50, height=50)
    c.setFontSize(30)
    c.drawString(60, 750, 'Swansea Printing Co Ltd')
    c.drawString(470, 750, 'INVOICE')
    c.setFontSize(10)
    c.drawString(10, 730, '31 Oxford Street')
    c.drawString(10, 715, 'Swansea, SA1 3AN')
    c.drawString(10, 700, 'Mobile: 07828522306')
    c.drawString(10, 685, 'Email: swanseaprintco@gmail.com')
    c.setFontSize(10)
    c.drawString(450, 715, 'DATE')
    c.drawString(450, 703, 'INVOICE #')
    c.drawString(450, 691, 'CUSTOMER ID')
    c.drawString(450, 679, 'DUE DATE')
    c.drawString(530, 715, date.today().strftime("%d/%m/%Y"))
    c.drawString(530, 703, str(customer_details['Customer ID']))
    c.drawString(530, 691, str(customer_details['Customer ID']))
    c.drawString(530, 679, (date.today() + timedelta(5)).strftime("%d/%m/%Y"))
    y = 635
    c.setFillColorRGB(42/255, 45/255, 117/255)
    c.rect(0, y - 4, width=150, height=15, stroke=0, fill=True)
    c.setFillColorRGB(1, 1, 1)
    c.drawString(10, y - 0, 'BILL TO')
    c.setFillColorRGB(0, 0, 0)
    c.drawString(10, y - 15, str(customer_details['Customer ID']))
    c.drawString(10, y - 28, customer_details['Company Name'])
    c.drawString(10, y - 41, customer_details['Address'])
    c.drawString(10, y - 54, f"{customer_details['City']}, {customer_details['Postcode']}")
    c.drawString(10, y - 67, str(customer_details['Phone']))
    c.drawImage('data/logo.png', 500, y - 80, width=100, height=100)
    y = 190
    x = 470
    c.drawString(x, y - 0, "Subtotal")
    c.drawString(x, y - 14, "Taxable")
    c.drawString(x, y - 28, "Tax rate")
    c.drawString(x, y - 42, "Tax due")
    c.drawString(x, y - 56, "Other")
    c.setLineWidth(0.3)
    c.line(x-10, y-60, x+110, y-60)
    c.drawString(x, y - 72, "Total")
    x = 540
    c.drawString(x, y - 0, str(order_details['AMOUNT'].sum()))
    c.drawString(x, y - 14, "-")
    c.drawString(x, y - 28, "0.000%")
    c.drawString(x, y - 42, "-")
    c.drawString(x, y - 56, "-")
    c.drawString(x, y - 70, str(order_details['AMOUNT'].sum()))

def invoice_to_pdf(filename, customer_details, order_details):
    doc = SimpleDocTemplate(filename, pagesize=letter)
    elements = []
    elements.append(Spacer(1, 2.5*inch))
    data = [list(order_details.columns)] + order_details.values.tolist()
    #print(order_details)
    #print(data)
    data.extend([["", "", "", "", "-"]] * (16 - len(order_details)))
    table_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2a2d75')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#e8e8ed')]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#878791')),
    ]
    table = Table(data, style=table_style, colWidths=[300, 70, 70, 70, 70])
    elements.append(table)

    disclaimer_data = [["OTHER COMMENTS", ""], ["""1. Total payment due before order processing
2. Please include the invoice number on your check
3. Payment via cash/card in store, or bank transfer
4. Tide Bank Transfer Details:
    Sortcode: 04-06-05 Account Number: 21805496
    """, ""]]

    disclaimer_table_style = [
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#2a2d75')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (0, -1), 0.3, colors.black),
    ]
    disclaimer_table = Table(disclaimer_data, style=disclaimer_table_style, colWidths=[300, 280])
    elements.append(Spacer(1, 0.3*inch))
    elements.append(disclaimer_table)

    doc.build(elements, onFirstPage=lambda can, doc: draw_on_canvas(can, doc, customer_details, order_details))


def create_invoice(df, filename, customer_id):
    sku_mapping_file = 'data/sku_to_name_mapping.csv'
    customer_details_file = 'data/Customer Details.csv'
    price_file = 'data/price_mapping.csv'
    customer_id = int(customer_id)
    sku_mapping = read_csv(sku_mapping_file)
    customer_details = pd.read_csv(customer_details_file)
    price = read_csv(price_file)
    print(customer_id)
    print(customer_details)
    customer_details = customer_details[customer_details['Customer ID'] == customer_id].to_dict(orient='records')[0]
    # print(customer_details)
    df['DESCRIPTION'] = df['SKU TYPE'].map(lambda x: sku_mapping[x])
    df['UNIT PRICE'] = df['SKU TYPE'].map(lambda x: float(price[x]))
    df['QTY'] = df['quantity'].map(lambda x: int(x))
    df['TAXED'] = df['quantity'].map(lambda x: "")
    df['AMOUNT'] = df['QTY'] * df['UNIT PRICE']
    cols = ['DESCRIPTION', 'UNIT PRICE', 'QTY', 'TAXED', 'AMOUNT']
    invoice_to_pdf(filename, customer_details, df[cols])



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
    'PACK FILE': r"C:\Users\Aamir Ali\Downloads\Order Examples-20240311T143007Z-001\Order Examples\PDF\29.2.24 - 24 Orders Pack 10001.pdf",
    'POST FILE': r"C:\Users\Aamir Ali\Downloads\Order Examples-20240311T143007Z-001\Order Examples\PDF\29.2.24 - 24 Orders Post 10001.pdf",
    'IMAGE FOLDER': r"C:\Users\Aamir Ali\Downloads\Order Examples-20240311T143007Z-001\Order Examples\PNG",
    'OUTPUT IMAGE FOLDER': r"C:\Users\Aamir Ali\Downloads\Order Examples-20240311T143007Z-001\Order Examples\hbl",
    'OUTPUT FOLDER': r"C:\Users\Aamir Ali\Downloads\Order Examples-20240311T143007Z-001\Order Examples",
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
        with open(paths['OUTPUT FOLDER'] + '/OUTPUT.pdf', 'wb') as f:
            extracted_info.write(f)
        df = pd.DataFrame(garment_pick_list)
        grouped_df = df.groupby(['name', 'size', 'color'])['quantity'].sum().reset_index().sort_values(by=['name', 'size', 'color'])
        export_to_pdf(grouped_df, paths['OUTPUT FOLDER'] + '/pick_list.pdf')
        grouped_df = df.groupby(['SKU TYPE'])['quantity'].sum().reset_index()
        create_invoice(grouped_df, paths['OUTPUT FOLDER'] + '/invoice.pdf', customer_id)
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
    
