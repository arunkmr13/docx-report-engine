from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def apply_columns(doc):
    for para in doc.paragraphs:
        if "Section 3" in para.text:
            new_section = doc.add_section(WD_SECTION.CONTINUOUS)
            sectPr = new_section._sectPr

            cols = OxmlElement('w:cols')
            cols.set(qn('w:num'), '2')
            sectPr.append(cols)
            break