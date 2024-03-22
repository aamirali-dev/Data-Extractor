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


    def to_pdf(self, filename):
        """
        Saves the pick list as pdf to the specified filename.

        Args:
            filename (str): The filename to save the PDF document to.
        """
        # doc = SimpleDocTemplate(filename, pagesize=letter)
        doc = BaseDocTemplate(filename)
        # creating the table by combining column names & rows
        elements = [
            Table(
                list(self.df.columns) + self.df.values.tolist()
            )
        ]
        style = TableStyle(
                        [
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]
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