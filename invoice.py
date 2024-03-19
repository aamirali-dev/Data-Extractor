
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

def read_csv(filename):
    try:
        with open(filename) as file:
            data = file.readlines()
        data = {line.strip().split(',')[0]: line.strip().split(',')[1] for line in data}
        return data 
    except:
        return None
    
class Invoice:
    def __init__(self, df, customer_id, data_dir='data'):
        self.customer_id = int(customer_id)
        self.data_dir = data_dir
        self.df = df
        self.load_data_from_files()
        self.initialize_order_details()
        
    def initialize_order_details(self):
        df = self.df
        df['DESCRIPTION'] = df['SKU TYPE'].map(lambda x: self.sku_mapping[x])
        df['UNIT PRICE'] = df['SKU TYPE'].map(lambda x: float(self.price[x]))
        df['QTY'] = df['quantity'].map(lambda x: int(x))
        df['TAXED'] = df['quantity'].map(lambda x: "")
        df['AMOUNT'] = df['QTY'] * df['UNIT PRICE']
        cols = ['DESCRIPTION', 'UNIT PRICE', 'QTY', 'TAXED', 'AMOUNT']
        self.order_details = df[cols]
    
    def load_data_from_files(self):
        sku_mapping_file = self.data_dir + '/sku_to_name_mapping.csv'
        customer_details_file = self.data_dir + '/Customer Details.csv'
        price_file = self.data_dir + '/price_mapping.csv'
        
        self.sku_mapping = read_csv(sku_mapping_file)
        self.price = read_csv(price_file)
        customer_details = pd.read_csv(customer_details_file)
        
        self.customer_details = customer_details[customer_details['Customer ID'] == self.customer_id].to_dict(orient='records')[0]
        
    def to_pdf(self, filename):
        doc = SimpleDocTemplate(filename, pagesize=letter)
        elements = []
        elements.append(Spacer(1, 2.5*inch))
        elements.append(self.get_order_table())

        elements.append(Spacer(1, 0.3*inch))
        elements.append(self.get_disclaimer_table())

        doc.build(elements, onFirstPage=self.draw_on_canvas)
        
    def get_order_table(self):
        data = [list(self.order_details.columns)] + self.order_details.values.tolist()
        data.extend([["", "", "", "", "-"]] * (16 - len(self.order_details)))
        table_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2a2d75')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#e8e8ed')]),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#878791')),
        ]
        return Table(data, style=table_style, colWidths=[300, 70, 70, 70, 70]) 
    
    def get_disclaimer_table(self):
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
        return Table(disclaimer_data, style=disclaimer_table_style, colWidths=[300, 280])
    
    def draw_on_canvas(self, c, doc):
        self.draw_branding(c)
        self.draw_invoice_metadata(c)
        self.draw_customer_details(c)
        self.draw_order_details(c)

    def draw_branding(self, c):
        c.drawImage('data/logo.png', 5, 740, width=50, height=50)
        c.setFontSize(30)
        c.drawString(60, 750, 'Swansea Printing Co Ltd')
        c.drawString(470, 750, 'INVOICE')
        c.setFontSize(10)
        c.drawString(10, 730, '31 Oxford Street')
        c.drawString(10, 715, 'Swansea, SA1 3AN')
        c.drawString(10, 700, 'Mobile: 07828522306')
        c.drawString(10, 685, 'Email: swanseaprintco@gmail.com')

    def draw_invoice_metadata(self, c):
        c.setFontSize(10)
        c.drawString(450, 715, 'DATE')
        c.drawString(450, 703, 'INVOICE #')
        c.drawString(450, 691, 'CUSTOMER ID')
        c.drawString(450, 679, 'DUE DATE')
        c.drawString(530, 715, date.today().strftime("%d/%m/%Y"))
        c.drawString(530, 703, str(self.customer_details['Customer ID']))
        c.drawString(530, 691, str(self.customer_details['Customer ID']))
        c.drawString(530, 679, (date.today() + timedelta(5)).strftime("%d/%m/%Y"))

    def draw_customer_details(self, c):
        customer_details = self.customer_details
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
        
    def draw_order_details(self, c):
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
        c.drawString(x, y - 0, str(self.order_details['AMOUNT'].sum()))
        c.drawString(x, y - 14, "-")
        c.drawString(x, y - 28, "0.000%")
        c.drawString(x, y - 42, "-")
        c.drawString(x, y - 56, "-")
        c.drawString(x, y - 70, str(self.order_details['AMOUNT'].sum()))