"""
Replace the {{ewp_image}} placeholder in a DOCX with an uploaded image.

Handles:
- Placeholders in body paragraphs, tables, nested tables, headers, footers
- Multiple placeholders (replaces each with a FRESH stream)
- Image format conversion (WEBP, HEIC, BMP → PNG) if needed
- Stream re-creation for each insertion
"""
from docx import Document
from docx.shared import Inches
from io import BytesIO
import tempfile
import os


# ============================================================
# Image normalization — ensure python-docx can read it
# ============================================================
def _normalize_image_bytes(image_bytes: bytes) -> bytes:
    """
    Ensure the image bytes are in a format python-docx can read.
    If Pillow can open it, re-encode as PNG for maximum compatibility.
    """
    try:
        from PIL import Image
    except ImportError:
        # Pillow not available — assume bytes are already valid
        return image_bytes

    try:
        img = Image.open(BytesIO(image_bytes))
        # Force load to validate
        img.load()

        # Convert to RGB/RGBA (handles paletted images, CMYK, etc.)
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA")

        # Re-encode as PNG for consistency
        buf = BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        # If Pillow can't read it either, return original
        return image_bytes


# ============================================================
# Stream factory — creates a FRESH BytesIO each time
# ============================================================
def _make_stream_factory(image_bytes: bytes):
    """
    Returns a function that creates a fresh BytesIO stream on each call.
    Use this to insert the SAME image in multiple places.
    """
    def factory():
        return BytesIO(image_bytes)
    return factory


# ============================================================
# Paragraph-level image insertion
# ============================================================
def _insert_image_in_paragraph(paragraph, image_stream):
    """Insert image at the end of the paragraph."""
    run = paragraph.add_run()
    run.add_picture(image_stream, width=Inches(6.0))


def _find_and_replace_image_in_paragraph(paragraph, stream_factory):
    """
    If paragraph contains {{ewp_image}}, replace with a fresh image stream.

    Args:
        paragraph: python-docx paragraph
        stream_factory: callable that returns a NEW BytesIO each time

    Returns:
        True if replacement happened, False otherwise
    """
    full_text = "".join(r.text for r in paragraph.runs)
    if "{{ewp_image}}" not in full_text:
        return False

    # Clear existing text in runs
    for run in paragraph.runs:
        run.text = ""

    # Insert image using a FRESH stream
    stream = stream_factory()
    _insert_image_in_paragraph(paragraph, stream)
    return True


# ============================================================
# Table-level image replacement
# ============================================================
def _replace_in_table(table, stream_factory):
    """Recursively process tables — pass stream_factory, not a shared stream."""
    replaced_any = False

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                if _find_and_replace_image_in_paragraph(paragraph, stream_factory):
                    replaced_any = True
            # Handle nested tables
            for nested in cell.tables:
                if _replace_in_table(nested, stream_factory):
                    replaced_any = True

    return replaced_any


# ============================================================
# Main entry point
# ============================================================
def replace_ewp_image(docx_path, image_stream, output_path):
    """
    Replace {{ewp_image}} placeholders in the DOCX with the uploaded image.

    Args:
        docx_path: Path to input DOCX
        image_stream: BytesIO containing the EWP image
        output_path: Path to save output DOCX

    Returns:
        (output_path, replaced: bool)
    """
    # Read the image bytes ONCE and normalize format
    image_stream.seek(0)
    raw_bytes = image_stream.read()

    # Normalize to ensure python-docx can read it
    image_bytes = _normalize_image_bytes(raw_bytes)

    # Create a factory that produces FRESH streams for each insertion
    stream_factory = _make_stream_factory(image_bytes)

    # Load the DOCX
    doc = Document(docx_path)

    replaced = False

    # 1. Body paragraphs
    for paragraph in doc.paragraphs:
        if _find_and_replace_image_in_paragraph(paragraph, stream_factory):
            replaced = True

    # 2. Tables (including nested)
    for table in doc.tables:
        if _replace_in_table(table, stream_factory):
            replaced = True

    # 3. Headers & Footers
    for section in doc.sections:
        # Headers
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for paragraph in header.paragraphs:
                    if _find_and_replace_image_in_paragraph(paragraph, stream_factory):
                        replaced = True
                for table in header.tables:
                    if _replace_in_table(table, stream_factory):
                        replaced = True

        # Footers
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for paragraph in footer.paragraphs:
                    if _find_and_replace_image_in_paragraph(paragraph, stream_factory):
                        replaced = True
                for table in footer.tables:
                    if _replace_in_table(table, stream_factory):
                        replaced = True

    # Save
    doc.save(output_path)
    return output_path, replaced
