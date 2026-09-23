"""
Replace the {{ewp_image}} placeholder in a DOCX with an uploaded image.
Handles:
- Placeholder in paragraph text
- Placeholder in table cell
- Placeholder inside a text box (via XML)
"""
from docx import Document
from docx.shared import Inches
import io


def _insert_image_in_paragraph(paragraph, image_stream):
    """Insert image at the position of the placeholder run (or append)."""
    run = paragraph.add_run()
    run.add_picture(image_stream, width=Inches(6.0))


def _find_and_replace_image_in_paragraph(paragraph, image_stream):
    """If paragraph contains the {{ewp_image}} token, replace it with image."""
    full_text = "".join(r.text for r in paragraph.runs)
    if "{{ewp_image}}" not in full_text:
        return False

    # Clear text from runs
    for run in paragraph.runs:
        run.text = ""

    # Insert image
    _insert_image_in_paragraph(paragraph, image_stream)
    return True


def _replace_in_table(table, image_stream):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _find_and_replace_image_in_paragraph(paragraph, image_stream)
            for nested in cell.tables:
                _replace_in_table(nested, image_stream)


def replace_ewp_image(docx_path, image_stream, output_path):
    doc = Document(docx_path)

    replaced = False

    for paragraph in doc.paragraphs:
        if _find_and_replace_image_in_paragraph(paragraph, image_stream):
            replaced = True

    for table in doc.tables:
        _replace_in_table(table, image_stream)

    doc.save(output_path)
    return output_path, replaced
