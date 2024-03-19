from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, BaseDocTemplate, Frame, PageTemplate


class PickList:
    def __init__(self, df) -> None:
        self.df = df


    def to_pdf(self, filename):
        # doc = SimpleDocTemplate(filename, pagesize=letter)
        doc = BaseDocTemplate(filename)
        elements = [Table([list(self.df.columns)] + self.df.values.tolist())]
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