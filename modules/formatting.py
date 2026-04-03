from docx.shared import Pt

def apply_formatting(doc):
    for para in doc.paragraphs:
        if para.text.startswith("Section"):
            para.paragraph_format.line_spacing = 1
        else:
            para.paragraph_format.line_spacing = 1.5

        para.paragraph_format.space_after = Pt(8)