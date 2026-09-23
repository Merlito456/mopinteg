"""
Embed an Excel file (.xlsx) as an OLE object into a DOCX.

Uses low-level OOXML manipulation to create an authentic embedded Excel icon
that users can double-click to open in Excel.

Reference:
- OOXML spec: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/
- Word OLE structure: /word/embeddings/oleObject1.bin + rels + content types
"""

import os
import re
import shutil
import zipfile
from pathlib import Path
from io import BytesIO


# ============================================================
# Constants
# ============================================================
OLE_PROG_ID = "Excel.Sheet.12"          # Excel 2007+ (.xlsx)
OLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.oleObject"
OLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
ICON_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


# ============================================================
# Main Entry Point
# ============================================================
def embed_excel_in_docx(
    docx_path: str,
    xlsx_bytes: bytes,
    xlsx_filename: str,
    placeholder: str = "{{FIO_EMBED}}",
    output_path: str = None,
) -> str:
    """
    Replace a {{FIO_EMBED}} placeholder in a DOCX with a real OLE-embedded Excel file.
    
    Args:
        docx_path: Path to the input .docx
        xlsx_bytes: Raw bytes of the .xlsx file to embed
        xlsx_filename: Original filename (used for icon display name)
        placeholder: The placeholder token to replace (default: {{FIO_EMBED}})
        output_path: Where to save the output. If None, uses *_with_attachment.docx
    
    Returns:
        Path to the generated .docx
    """
    docx_path = Path(docx_path)
    if output_path is None:
        output_path = docx_path.parent / f"{docx_path.stem}_with_attachment.docx"
    else:
        output_path = Path(output_path)

    # 1. Extract DOCX (it's a ZIP)
    temp_dir = docx_path.parent / f"_temp_ole_{docx_path.stem}"
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(temp_dir)

    # 2. Check if the placeholder exists in document.xml
    doc_xml_path = temp_dir / "word" / "document.xml"
    doc_xml = doc_xml_path.read_text(encoding="utf-8")

    if placeholder not in doc_xml:
        raise ValueError(
            f"Placeholder '{placeholder}' not found in document.xml. "
            "Add it to the template where the FIO attachment should appear."
        )

    # 3. Prepare embeddings folder
    embeddings_dir = temp_dir / "word" / "embeddings"
    embeddings_dir.mkdir(parents=True, exist_ok=True)

    # Find next available oleObject index
    existing_ole = list(embeddings_dir.glob("oleObject*.bin"))
    ole_index = len(existing_ole) + 1
    ole_filename = f"oleObject{ole_index}.bin"
    ole_path = embeddings_dir / ole_filename

    # 4. Write the OLE object (raw xlsx bytes)
    with open(ole_path, "wb") as f:
        f.write(xlsx_bytes)

    # 5. Create Excel icon (or reuse existing)
    # For best results, embed a real PNG icon. Here we generate a simple one.
    icon_path, icon_filename = _ensure_excel_icon(temp_dir, ole_index)

    # 6. Find/create relationship IDs
    rels_path = temp_dir / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")

    rids = re.findall(r'Id="rId(\d+)"', rels_xml)
    next_rid_ole = max(int(r) for r in rids) + 1 if rids else 1
    next_rid_icon = next_rid_ole + 1

    # Add OLE relationship
    ole_rel = (
        f'<Relationship Id="rId{next_rid_ole}" '
        f'Type="{OLE_REL_TYPE}" '
        f'Target="embeddings/{ole_filename}"/>'
    )
    # Add icon image relationship
    icon_rel = (
        f'<Relationship Id="rId{next_rid_icon}" '
        f'Type="{ICON_REL_TYPE}" '
        f'Target="media/{icon_filename}"/>'
    )

    rels_xml = rels_xml.replace(
        "</Relationships>",
        ole_rel + icon_rel + "</Relationships>",
    )
    rels_path.write_text(rels_xml, encoding="utf-8")

    # 7. Build the OLE object XML to replace the placeholder
    ole_xml = _build_ole_xml(
        ole_rid=f"rId{next_rid_ole}",
        icon_rid=f"rId{next_rid_icon}",
        display_name=xlsx_filename,
    )

    # 8. Replace placeholder in document.xml
    # NOTE: The placeholder may be split across runs in Word. We handle both cases.
    doc_xml = _replace_placeholder_with_ole(doc_xml, placeholder, ole_xml)
    doc_xml_path.write_text(doc_xml, encoding="utf-8")

    # 9. Update [Content_Types].xml
    ct_path = temp_dir / "[Content_Types].xml"
    ct_xml = ct_path.read_text(encoding="utf-8")

    # Add Default for .bin (OLE object) if missing
    if 'Extension="bin"' not in ct_xml:
        ct_xml = ct_xml.replace(
            "</Types>",
            f'<Default Extension="bin" ContentType="{OLE_CONTENT_TYPE}"/></Types>',
        )
    # Add Default for .png (icon) if missing
    if 'Extension="png"' not in ct_xml:
        ct_xml = ct_xml.replace(
            "</Types>",
            '<Default Extension="png" ContentType="image/png"/></Types>',
        )

    ct_path.write_text(ct_xml, encoding="utf-8")

    # 10. Re-zip as DOCX
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(temp_dir)
                z.write(file_path, str(arcname).replace("\\", "/"))

    # Cleanup
    shutil.rmtree(temp_dir, ignore_errors=True)

    return str(output_path)


# ============================================================
# Icon Handling
# ============================================================
def _ensure_excel_icon(temp_dir: Path, ole_index: int) -> tuple:
    """
    Ensure an Excel icon exists in /word/media/.
    Returns (icon_path, icon_filename).
    """
    media_dir = temp_dir / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)

    icon_filename = f"excel_icon_{ole_index}.png"
    icon_path = media_dir / icon_filename

    # Check if an existing Excel icon already exists
    existing_icons = list(media_dir.glob("excel_icon*.png"))
    if existing_icons:
        return existing_icons[0], existing_icons[0].name

    # Generate a simple Excel icon (16x16 PNG)
    icon_bytes = _generate_excel_icon_bytes()
    with open(icon_path, "wb") as f:
        f.write(icon_bytes)

    return icon_path, icon_filename


def _generate_excel_icon_bytes() -> bytes:
    """
    Generate a minimal Excel-green icon PNG.
    In production, replace this with a real Excel icon asset.
    """
    try:
        from PIL import Image, ImageDraw
        img = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        # Excel green square
        draw.rectangle([(2, 2), (30, 30)], fill=(33, 115, 70, 255))
        # White "X" for Excel
        draw.line([(10, 10), (22, 22)], fill="white", width=3)
        draw.line([(22, 10), (10, 22)], fill="white", width=3)
        
        buf = BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except ImportError:
        # Fallback: minimal 1x1 transparent PNG
        return (
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
            b"\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
            b"\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n"
            b"\x2d\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
        )


# ============================================================
# OLE Object XML Builder
# ============================================================
def _build_ole_xml(ole_rid: str, icon_rid: str, display_name: str) -> str:
    """
    Build the OLE object XML for embedding into document.xml.
    This creates an Excel icon that users can double-click to open.
    """
    # Word OLE objects need: v:shape (icon visual) + o:OLEObject (the embedded file)
    ole_xml = (
        '<w:r>'
        '<w:object w:dxaOrig="1440" w:dyaOrig="1440">'
        # Icon visual
        '<v:shape id="_x0000_i1025" type="#_x0000_t75" '
        'style="width:24pt;height:24pt" fillcolor="auto" stroked="f">'
        f'<v:imagedata r:id="{icon_rid}" o:title="{display_name}"/>'
        '</v:shape>'
        # OLE object reference
        f'<o:OLEObject Type="Embed" ProgID="{OLE_PROG_ID}" '
        'ShapeID="_x0000_i1025" DrawAspect="Icon" '
        'ObjectID="_1234567890" '
        f'r:id="{ole_rid}">'
        f'<o:FieldCodes>\\s</o:FieldCodes>'
        '</o:OLEObject>'
        '</w:object>'
        '<w:rPr><w:noProof/></w:rPr>'
        '</w:r>'
    )
    return ole_xml


# ============================================================
# Placeholder Replacement (run-safe)
# ============================================================
def _replace_placeholder_with_ole(doc_xml: str, placeholder: str, ole_xml: str) -> str:
    """
    Replace placeholder in document.xml. Handles cases where the placeholder
    is split across multiple <w:t> runs (common in Word).
    """
    # Case 1: Placeholder is intact in one run
    if placeholder in doc_xml:
        # Replace inside <w:t> tags
        escaped_placeholder = re.escape(placeholder)
        doc_xml = re.sub(
            rf"<w:t[^>]*>{escaped_placeholder}</w:t>",
            ole_xml,
            doc_xml,
            count=1,
        )
        if placeholder not in doc_xml:
            return doc_xml

    # Case 2: Placeholder split across runs
    # Find the paragraph containing the placeholder
    placeholder_chars = list(placeholder)
    pattern = r"<w:t[^>]*>(.*?)</w:t>"
    
    # Find all text runs in document
    matches = list(re.finditer(pattern, doc_xml, re.DOTALL))
    
    # Reconstruct the concatenated text and find placeholder position
    for i, match in enumerate(matches):
        text = match.group(1)
        if placeholder_chars[0] in text:
            # Check if the placeholder spans multiple runs
            accumulated = text
            j = i + 1
            while j < len(matches) and len(accumulated) < len(placeholder):
                accumulated += matches[j].group(1)
                j += 1
            
            if placeholder in accumulated:
                # Found it! Replace the whole span
                start_pos = match.start()
                end_pos = matches[j - 1].end()
                
                # Replace with OLE XML
                doc_xml = doc_xml[:start_pos] + ole_xml + doc_xml[end_pos:]
                return doc_xml

    # Case 3: Not found — leave as-is (placeholder remains visible)
    return doc_xml


# ============================================================
# Test Utility
# ============================================================
def test_embedding(docx_path: str, xlsx_path: str, output_path: str = None) -> str:
    """
    Quick test function to verify embedding works.
    """
    with open(xlsx_path, "rb") as f:
        xlsx_bytes = f.read()
    
    xlsx_filename = os.path.basename(xlsx_path)
    
    return embed_excel_in_docx(
        docx_path=docx_path,
        xlsx_bytes=xlsx_bytes,
        xlsx_filename=xlsx_filename,
        placeholder="{{FIO_EMBED}}",
        output_path=output_path,
    )
