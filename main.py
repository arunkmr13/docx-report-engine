from docx import Document
from config import *
from content_generator import generate_content

from modules.header_footer import apply_header_footer
from modules.formatting import apply_formatting
from modules.sections import apply_columns
from modules.borders import apply_page_borders
from modules.styles import apply_global_styles

doc = Document()

# Generate heavy content
generate_content(doc)

# Apply features
apply_global_styles(doc)
apply_header_footer(doc, TITLE, QUARTER)
apply_formatting(doc)
apply_columns(doc)
apply_page_borders(doc)

# Save
doc.save(OUTPUT_PATH)

print("🚀 Script started")
print(f"✅ Report generated at {OUTPUT_PATH}")