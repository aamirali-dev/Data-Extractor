
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Table
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, Spacer
from reportlab.lib import colors
from reportlab.lib.units import inch
from datetime import date, timedelta

def read_csv(filename):
    """
    Reads a CSV file and returns a dictionary of key-value pairs.

    Args:
        filename (str): The path to the CSV file.

    Returns:
        dict: A dictionary of key-value pairs, where the key is the first column of the CSV file and the value is the second column.
    """
    try:
        with open(filename) as file:
            data = file.readlines()
        data = {line.strip().split(',')[0]: line.strip().split(',')[1] for line in data}
        return data 
    except:
        return None
    
class Invoice:
    def __init__(self, df, customer_id, data_dir='data'):
        """
        Initialize an instance of the Invoice class.

        Args:
            df (DataFrame): A Pandas DataFrame containing the order details.
            customer_id (int): The customer ID.
            data_dir (str, optional): The directory containing the data files. Defaults to 'data'.
        """
        self.customer_id = int(customer_id)
        self.data_dir = data_dir
        self.df = df
        self.load_data_from_files()
        self.initialize_order_details()
        
    def initialize_order_details(self):
        """
        Initialize the order details by creating desired columns from existing 
        columns and extracting information from other sources particularly
        fetching description & price. Finally, it selects only desired columns.
        """
        df = self.df
        df['DESCRIPTION'] = df['SKU TYPE'].map(lambda x: self.sku_mapping[x])
        df['UNIT PRICE'] = df['SKU TYPE'].map(lambda x: float(self.price[x]))
        df['QTY'] = df['quantity'].map(lambda x: int(x))
        df['TAXED'] = df['quantity'].map(lambda x: "")
        df['AMOUNT'] = df['QTY'] * df['UNIT PRICE']
        cols = ['DESCRIPTION', 'UNIT PRICE', 'QTY', 'TAXED', 'AMOUNT']
        self.order_details = df[cols]
    
    def load_data_from_files(self):
        """
        Load data from CSV files.

        The function reads three CSV files:
        - sku_to_name_mapping.csv: maps SKU codes to product names
        - Customer Details.csv: contains customer information
        - price_mapping.csv: maps SKU codes to prices

        The function stores the data in class attributes for later use.
        """
        sku_mapping_file = self.data_dir + '/sku_to_name_mapping.csv'
        customer_details_file = self.data_dir + '/Customer Details.csv'
        price_file = self.data_dir + '/price_mapping.csv'
        
        self.sku_mapping = read_csv(sku_mapping_file)
        self.price = read_csv(price_file)
        customer_details = pd.read_csv(customer_details_file)
        
        self.customer_details = customer_details[customer_details['Customer ID'] == self.customer_id].to_dict(orient='records')[0]
        
    def to_pdf(self, filename):
        """
        Creates a PDF document containing the invoice and save it.

        Args:
            filename (str): The path to the PDF file.

        Returns:
            None
        """
        
        doc = SimpleDocTemplate(filename, pagesize=letter)
        elements = []
        elements.append(Spacer(1, 2.5*inch))
        elements.append(self.get_order_table())

        elements.append(Spacer(1, 0.3*inch))
        elements.append(self.get_disclaimer_table())

        doc.build(elements, onFirstPage=self.draw_on_canvas)
        
    def get_order_table(self):
        """
        Returns a table containing the order details.

        The table is created by concatenating the column names to the order details
        and adding blank rows to make the total columns equal to 16.

        Returns:
            Table: A table containing the order details.
        """
        data = list(self.order_details.columns) + self.order_details.values.tolist()
        # making total columns equal to 16
        data.extend([["", "", "", "", "-"]] * (16 - len(self.order_details)))
        
        table_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2a2d75')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#e8e8ed')]),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#878791')),
        ]
        return Table(data, style=table_style, colWidths=[300, 70, 70, 70, 70]) 
    
    def get_disclaimer_table(self):
        """
        Returns a box containing the disclaimer.

        A table element is adopted for this purpose because it was easier. second column is hided using styling

        Returns:
            Table: A table containing the disclaimer.
        """
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
        return Table(
            disclaimer_data, 
            style=disclaimer_table_style, 
            colWidths=[300, 280]
        )
    
    def draw_on_canvas(self, c, doc):
        """
        Draws the branding, invoice metadata, customer details, and order details on the canvas. This function is passed to doc.build method and is used to draw elements that were not easier to place otherwise in the normal pdf flow.

        Args:
            c (Canvas): The canvas object.
            doc (SimpleDocTemplate): The document template object.

        Returns:
            None
        """
        self.draw_branding(c)
        self.draw_invoice_metadata(c)
        self.draw_customer_details(c)
        self.draw_order_details(c)

    def draw_branding(self, c):
        """
        Draws the shop details on the invoice

        Args:
            c (Canvas): The canvas object.
        """
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
        """
        Draws invoice metadata such as date, invoice id, customer id, due date, etc.

        Args:
            c (Canvas): The canvas object.
        """
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
        """
        Draws the details of the customer, such as customer id, company name, 
        address, as the customer in this case is also a company. we are developing this for a B2B type of business. 

        Args:
            c (Canvas): The canvas object.
        """
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
        """
        Draws order details such as total, subtotal, tax, final amount, etc. These are drawn here because it was not easy to draw them using normal pdf flow.

        Args:
            c (Canvas): The canvas object.
        """
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