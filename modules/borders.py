from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def apply_page_borders(doc):
    for section in doc.sections:
        sectPr = section._sectPr

        pgBorders = OxmlElement('w:pgBorders')
        pgBorders.set(qn('w:offsetFrom'), 'page')

        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '12')
            border.set(qn('w:space'), '24')
            border.set(qn('w:color'), '000000')
            pgBorders.append(border)

        sectPr.append(pgBorders)