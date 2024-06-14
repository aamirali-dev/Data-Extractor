"""
This class creates a PDF document with a table that contains the headers and data from a pandas dataframe.

Args:
    df (pandas.DataFrame): The dataframe containing the data to be displayed in the table.

Attributes:
    df (pandas.DataFrame): The dataframe containing the data to be displayed in the table.

Methods:
    to_pdf(filename: str): Saves the PDF document to the specified filename.
"""

from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, BaseDocTemplate, Frame, PageTemplate


class PickList:
    def __init__(self, df) -> None:
        self.df = df
        # self.df.to_csv('picklist.csv')
        self.sort_key = self.df.pop('Sort Key')
        # self.sort_key = self.df['Sort Key']


    def to_pdf(self, filename):
        """
        Saves the pick list as pdf to the specified filename.

        Args:
            filename (str): The filename to save the PDF document to.
        """
        # doc = SimpleDocTemplate(filename, pagesize=letter)
        doc = BaseDocTemplate(filename)
        # creating the table by combining column names & rows
        headers = list(self.df.columns)
        elements = [
            Table(
                [headers] + self.df.values.tolist()
            )
        ]
        light_colors = [
            # '#F5F5F5',  # White Smoke
            # '#FFFFE0',  # Light Yellow
            # '#FFEFD5',  # Papaya Whip
            # '#FAF0E6',  # Linen
            # '#FFF5EE',  # Sea Shell
            '#F0FFF0',  # Honeydew
            '#E6E6FA',  # Lavender
            # '#FFF8DC',  # Cornsilk
            # '#FFE4E1',  # Misty Rose
            # '#F0F8FF',  # Alice Blue
            # '#F0E68C',  # Khaki
            # '#FAEBD7',  # Antique White
            # '#F5FFFA',  # Mint Cream
            # '#FFE4B5',  # Moccasin
            # '#F0FFFF',  # Azure
            # '#F5F5DC',  # Beige
            # '#FAFAD2',  # Light Goldenrod Yellow
            # '#FFFAF0',  # Floral White
            # '#FFE4C4',  # Bisque
            # '#FFEBCD'   # Blanched Almond
        ]
        row_formatting = []
        for index, row in enumerate(self.sort_key):
            if row.startswith('4'):
                color_index = int(row.split('.')[1])%2
                color = light_colors[color_index]
                # print(row, color_index, color)
                row_formatting.append(('BACKGROUND', (0, index + 1), (-1, index + 1), colors.HexColor(color)))
            else:
                # row_formatting.append(('BACKGROUND', (0, index + 1), (-1, index + 1), colors.beige))
                pass

        style = TableStyle(
                        [
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ] + row_formatting
                    )
        elements[0].setStyle(style)
        # creating 2 frames for 2 column view, since our table is quite slim, 
        # using 2 column approach will save printing papers by compressing 
        # 2 tables at 1 page
        frame1 = Frame(
            doc.leftMargin, 
            doc.bottomMargin, 
            doc.width/2-6, 
            doc.height, 
            id='col1',
        )
        frame2 = Frame(
            doc.leftMargin+doc.width/2+6, 
            doc.bottomMargin, 
            doc.width/2-6, 
            doc.height, 
            id='col2',
        )
        doc.addPageTemplates(
            [
                PageTemplate(id='TwoCol',frames=[frame1,frame2]), 
            ]
        )
        doc.build(elements)

# if __name__ == "__main__":
#     import pandas as pd 
#     df = pd.read_csv('picklist.csv')
#     df.pop('Unnamed: 0')
#     PickList(df).to_pdf('picklist.pdf')