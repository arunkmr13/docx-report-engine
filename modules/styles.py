from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def get_style_safe(doc, style_name):
    try:
        return doc.styles[style_name]
    except KeyError:
        return None


def apply_global_styles(doc):
    normal = get_style_safe(doc, 'Normal')
    if normal:
        normal.font.name = "Calibri"
        normal.font.size = Pt(11)

    heading = get_style_safe(doc, 'Heading 1')
    if heading:
        heading.font.size = Pt(20)
        heading.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER