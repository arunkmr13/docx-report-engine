from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def apply_global_styles(doc):
    normal = doc.styles['Normal']
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    # Set heading size globally
    heading = doc.styles['Heading 1']
    heading.font.size = Pt(20)

    # Center all headings globally
    heading.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER