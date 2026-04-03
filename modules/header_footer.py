from docx.shared import Inches

def apply_header_footer(doc, title, quarter):

    for section in doc.sections:

        usable_width = section.page_width - section.left_margin - section.right_margin

        # ---------------- HEADER ----------------
        header = section.header
        header.paragraphs[0].text = ""

        table = header.add_table(rows=1, cols=2, width=usable_width)

        left_cell = table.rows[0].cells[0]
        right_cell = table.rows[0].cells[1]

        left_cell.text = "Molecular Connections"
        right_cell.text = f"{title} | {quarter}"

        right_cell.paragraphs[0].alignment = 2  # RIGHT

        # ---------------- FOOTER ----------------
        footer = section.footer
        footer.paragraphs[0].text = ""

        table = footer.add_table(rows=1, cols=2, width=usable_width)

        left_cell = table.rows[0].cells[0]
        right_cell = table.rows[0].cells[1]

        # Page number
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        para = left_cell.paragraphs[0]
        run = para.add_run("Page no: ")

        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.text = "PAGE"

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

        # Right text
        right_cell.text = "Arun Kumar"
        right_cell.paragraphs[0].alignment = 2  # RIGHT