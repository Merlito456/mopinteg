"""
Replace {{PLACEHOLDER}} tokens in a .docx file while preserving formatting.
Handles placeholders that may be split across multiple runs.
"""
from docx import Document


def _replace_in_paragraph(paragraph, mapping):
    """Replace placeholders in a single paragraph."""
    full_text = "".join(run.text for run in paragraph.runs)
    if "{{" not in full_text:
        return

    new_text = full_text
    for key, value in mapping.items():
        token = "{{" + key + "}}"
        if token in new_text:
            new_text = new_text.replace(token, str(value))

    if new_text != full_text:
        for run in paragraph.runs:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = new_text
        else:
            paragraph.add_run(new_text)


def _replace_in_table(table, mapping):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _replace_in_paragraph(paragraph, mapping)
            for nested_table in cell.tables:
                _replace_in_table(nested_table, mapping)


def replace_placeholders(docx_path, mapping, output_path):
    """Replace all placeholders in the DOCX file."""
    doc = Document(docx_path)

    for paragraph in doc.paragraphs:
        _replace_in_paragraph(paragraph, mapping)

    for table in doc.tables:
        _replace_in_table(table, mapping)

    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for paragraph in header.paragraphs:
                    _replace_in_paragraph(paragraph, mapping)
                for table in header.tables:
                    _replace_in_table(table, mapping)
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for paragraph in footer.paragraphs:
                    _replace_in_paragraph(paragraph, mapping)
                for table in footer.tables:
                    _replace_in_table(table, mapping)

    doc.save(output_path)
    return output_path
