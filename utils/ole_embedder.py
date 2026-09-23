"""
Embed an Excel file (.xlsx) as an OLE object into a DOCX.

Two approaches:
1. embed_excel_as_package() — Word 2016+ native (recommended, mimics drag-and-drop)
2. embed_excel_in_docx()     — Legacy OLE (fallback for older Word)

Reference:
- OOXML spec: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/
- Word OLE structure: /word/embeddings/oleObject1.bin + rels + content types
"""

import os
import re
import shutil
import zipfile
import html
import time
import random
from pathlib import Path
from io import BytesIO


# ============================================================
# Constants
# ============================================================
OLE_PROG_ID = "Excel.Sheet.12"          # Excel 2007+ (.xlsx)
PACKAGE_PROG_ID = "Package"             # Modern package embed
OLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.oleObject"
XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

OLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
PACKAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
ICON_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


# ============================================================
# Required namespaces for OLE embedding
# ============================================================
REQUIRED_NAMESPACES = {
    "xmlns:o": "urn:schemas-microsoft-com:office:office",
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:w10": "urn:schemas-microsoft-com:office:word",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xmlns:wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


# ============================================================
# Diagnostic logger
# ============================================================
def _log(msg: str):
    """Print to stdout (Streamlit logs it) AND to session state for UI."""
    print(f"[OLE] {msg}")
    try:
        import streamlit as st
        if "ole_debug_log" not in st.session_state:
            st.session_state.ole_debug_log = []
        st.session_state.ole_debug_log.append(msg)
    except Exception:
        pass


# ============================================================
# Namespace fix — CRITICAL for OLE to render
# ============================================================
def _ensure_required_namespaces(doc_xml: str) -> str:
    """Ensure document.xml has all namespaces required for OLE embedding."""
    doc_tag_match = re.search(r"<w:document\b[^>]*>", doc_xml)
    if not doc_tag_match:
        _log("  [NS] ⚠️ Could not find <w:document> tag")
        return doc_xml

    doc_tag = doc_tag_match.group(0)
    new_tag = doc_tag
    added = []

    for ns_prefix, ns_uri in REQUIRED_NAMESPACES.items():
        if f"{ns_prefix}=" not in doc_tag:
            new_tag = new_tag[:-1] + f' {ns_prefix}="{ns_uri}">'
            added.append(ns_prefix)

    if new_tag != doc_tag:
        doc_xml = doc_xml.replace(doc_tag, new_tag, 1)
        _log(f"  [NS] Added namespaces: {', '.join(added)}")
    else:
        _log("  [NS] All namespaces already present")

    return doc_xml


# ============================================================
# Unique ID generation — prevents collisions
# ============================================================
def _find_existing_ole_ids(doc_xml: str) -> dict:
    """Find existing ShapeID, ObjectID, and anchor values in the document."""
    return {
        "shape_ids": set(re.findall(r'ShapeID="([^"]+)"', doc_xml)),
        "object_ids": set(re.findall(r'ObjectID="([^"]+)"', doc_xml)),
        "shape_attrs": set(re.findall(r'<v:shape[^>]*\sid="([^"]+)"', doc_xml)),
        "anchors": set(re.findall(r'w14:anchorId="([^"]+)"', doc_xml)),
    }


def _generate_unique_ole_ids(doc_xml: str) -> dict:
    """Generate unique ShapeID, ObjectID, and anchorId."""
    existing = _find_existing_ole_ids(doc_xml)

    shape_id = None
    for _ in range(200):
        candidate = f"_x0000_i{random.randint(1028, 99999)}"
        if (candidate not in existing["shape_ids"]
                and candidate not in existing["shape_attrs"]):
            shape_id = candidate
            break
    if shape_id is None:
        shape_id = f"_x0000_i{int(time.time() * 1000) % 100000}"

    object_id = None
    for _ in range(200):
        candidate = f"_{random.randint(1000000000, 9999999999)}"
        if candidate not in existing["object_ids"]:
            object_id = candidate
            break
    if object_id is None:
        object_id = f"_{int(time.time() * 1000000)}"

    anchor_id = None
    for _ in range(200):
        candidate = ''.join(random.choices('0123456789ABCDEF', k=8))
        if candidate not in existing["anchors"]:
            anchor_id = candidate
            break
    if anchor_id is None:
        anchor_id = ''.join(random.choices('0123456789ABCDEF', k=8))

    _log(f"  [ID] Existing: {len(existing['shape_ids'])} shapeIDs, "
         f"{len(existing['object_ids'])} objectIDs, "
         f"{len(existing['anchors'])} anchors")
    _log(f"  [ID] New: shape={shape_id}, object={object_id}, anchor={anchor_id}")

    return {
        "shape_id": shape_id,
        "object_id": object_id,
        "anchor_id": anchor_id,
    }


# ============================================================
# PACKAGE Approach — Modern Word (RECOMMENDED)
# ============================================================
def embed_excel_as_package(
    docx_path: str,
    xlsx_bytes: bytes,
    xlsx_filename: str,
    placeholder: str = "{{FIO_REF}}",
    output_path: str = None,
) -> str:
    """
    Embed an .xlsx file using Word's PACKAGE format (most reliable for Word 2016+).

    This mimics Word's modern drag-and-drop behavior:
    - Stores the .xlsx as an embedded package (oleObject1.xlsx, not .bin)
    - Uses ProgID="Package" with Content-Type Override
    - Word renders it as a clickable icon that opens Excel on double-click
    """
    docx_path = Path(docx_path)
    if output_path is None:
        output_path = docx_path.parent / f"{docx_path.stem}_with_package.docx"
    else:
        output_path = Path(output_path)

    _log("=" * 60)
    _log("EMBED EXCEL AS PACKAGE — START")
    _log(f"  Source:      {docx_path}")
    _log(f"  Output:      {output_path}")
    _log(f"  Placeholder: {placeholder!r}")
    _log(f"  XLSX bytes:  {len(xlsx_bytes):,}")
    _log(f"  XLSX name:   {xlsx_filename}")

    # 1. Extract DOCX
    temp_dir = docx_path.parent / f"_temp_pkg_{docx_path.stem}"
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            z.extractall(temp_dir)
        _log(f"  [1] Extracted DOCX → {temp_dir}")
    except Exception as e:
        raise ValueError(f"Failed to extract DOCX: {e}")

    # 2. Load document.xml
    doc_xml_path = temp_dir / "word" / "document.xml"
    try:
        doc_xml = doc_xml_path.read_text(encoding="utf-8")
    except Exception as e:
        raise ValueError(f"Failed to read document.xml: {e}")

    _log(f"  [2] document.xml size: {len(doc_xml):,} chars")
    _log(f"  [2] Placeholder found in XML: {placeholder in doc_xml}")

    if placeholder not in doc_xml:
        # Case-insensitive fallback
        if placeholder.upper() in doc_xml.upper():
            _log("  [2] ⚠️ Found case-insensitive match")
            idx = doc_xml.upper().find(placeholder.upper())
            actual = doc_xml[idx:idx + len(placeholder)]
            _log(f"  [2] ⚠️ Actual text: {actual!r}")
            placeholder = actual
        else:
            raise ValueError(
                f"Placeholder '{placeholder}' not found in document.xml."
            )

    # 3. Create embeddings folder
    embeddings_dir = temp_dir / "word" / "embeddings"
    embeddings_dir.mkdir(parents=True, exist_ok=True)

    # Find next available index (checking both .bin and .xlsx)
    existing_bin = list(embeddings_dir.glob("oleObject*.bin"))
    existing_xlsx = list(embeddings_dir.glob("oleObject*.xlsx"))
    ole_index = len(existing_bin) + len(existing_xlsx) + 1

    package_filename = f"oleObject{ole_index}.xlsx"
    package_path = embeddings_dir / package_filename

    # 4. Write the xlsx directly (as .xlsx, not .bin)
    with open(package_path, "wb") as f:
        f.write(xlsx_bytes)
    _log(f"  [4] Wrote package: {package_filename} ({len(xlsx_bytes):,} bytes)")

    # 5. Create Excel icon
    icon_path, icon_filename = _ensure_excel_icon(temp_dir, ole_index)
    _log(f"  [5] Icon ready: {icon_filename}")

    # 6. Add relationships
    rels_path = temp_dir / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")

    rids = re.findall(r'Id="rId(\d+)"', rels_xml)
    next_rid_pkg = max(int(r) for r in rids) + 1 if rids else 1
    next_rid_icon = next_rid_pkg + 1

    # ⚠️ PACKAGE relationship (different from oleObject!)
    package_rel = (
        f'<Relationship Id="rId{next_rid_pkg}" '
        f'Type="{PACKAGE_REL_TYPE}" '
        f'Target="embeddings/{package_filename}"/>'
    )
    icon_rel = (
        f'<Relationship Id="rId{next_rid_icon}" '
        f'Type="{ICON_REL_TYPE}" '
        f'Target="media/{icon_filename}"/>'
    )

    rels_xml = rels_xml.replace(
        "</Relationships>",
        package_rel + icon_rel + "</Relationships>",
    )
    rels_path.write_text(rels_xml, encoding="utf-8")
    _log(f"  [6] Added rels: rId{next_rid_pkg} (package), rId{next_rid_icon} (icon)")

    # 7. Generate unique IDs
    unique_ids = _generate_unique_ole_ids(doc_xml)

    # 8. Build Package XML
    package_xml = _build_package_xml(
        package_rid=f"rId{next_rid_pkg}",
        icon_rid=f"rId{next_rid_icon}",
        display_name=xlsx_filename,
        shape_id=unique_ids["shape_id"],
        object_id=unique_ids["object_id"],
        anchor_id=unique_ids["anchor_id"],
    )
    _log(f"  [8] Package XML length: {len(package_xml)} chars")

    # 9. Replace placeholder
    doc_xml, replaced = _replace_placeholder_with_ole(doc_xml, placeholder, package_xml)
    _log(f"  [9] Replaced: {replaced}")

    if not replaced:
        raise ValueError(
            f"Placeholder '{placeholder}' was found but could not be replaced."
        )

    # 9b. Ensure namespaces
    doc_xml = _ensure_required_namespaces(doc_xml)
    doc_xml_path.write_text(doc_xml, encoding="utf-8")

    # 10. Update [Content_Types].xml — need Override for the .xlsx package
    ct_path = temp_dir / "[Content_Types].xml"
    ct_xml = ct_path.read_text(encoding="utf-8")

    override_xml = (
        f'<Override PartName="/word/embeddings/{package_filename}" '
        f'ContentType="{XLSX_CONTENT_TYPE}"/>'
    )

    if f"/word/embeddings/{package_filename}" not in ct_xml:
        ct_xml = ct_xml.replace("</Types>", override_xml + "</Types>")
        _log(f"  [10] Added Override for {package_filename}")

    if 'Extension="png"' not in ct_xml:
        ct_xml = ct_xml.replace(
            "</Types>",
            '<Default Extension="png" ContentType="image/png"/></Types>',
        )
        _log("  [10] Added .png content type")

    ct_path.write_text(ct_xml, encoding="utf-8")

    # 11. Re-zip
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(temp_dir)
                z.write(file_path, str(arcname).replace("\\", "/"))

    _log(f"  [11] Re-zipped → {output_path}")
    _log("EMBED PACKAGE — ✅ SUCCESS")
    _log("=" * 60)

    shutil.rmtree(temp_dir, ignore_errors=True)
    return str(output_path)


def _build_package_xml(
    package_rid: str,
    icon_rid: str,
    display_name: str,
    shape_id: str,
    object_id: str,
    anchor_id: str,
) -> str:
    """Build the Package (modern) embedding XML."""
    safe_name = html.escape(display_name)

    return (
        '<w:r>'
        '<w:rPr><w:noProof/></w:rPr>'
        f'<w:object w:dxaOrig="1440" w:dyaOrig="1440" w14:anchorId="{anchor_id}">'
        f'<v:shape id="{shape_id}" type="#_x0000_t75" '
        'style="width:32pt;height:32pt" o:ole="">'
        f'<v:imagedata r:id="{icon_rid}" o:title="{safe_name}"/>'
        '</v:shape>'
        # Package uses ProgID="Package" and NO FieldCodes
        f'<o:OLEObject Type="Embed" ProgID="{PACKAGE_PROG_ID}" '
        f'ShapeID="{shape_id}" DrawAspect="Icon" ObjectID="{object_id}" '
        f'r:id="{package_rid}"/>'
        '</w:object>'
        '</w:r>'
        # Filename text after icon
        '<w:r>'
        '<w:rPr><w:noProof/></w:rPr>'
        f'<w:t xml:space="preserve"> {safe_name}</w:t>'
        '</w:r>'
    )


# ============================================================
# LEGACY OLE Approach — Fallback
# ============================================================
def embed_excel_in_docx(
    docx_path: str,
    xlsx_bytes: bytes,
    xlsx_filename: str,
    placeholder: str = "{{FIO_REF}}",
    output_path: str = None,
) -> str:
    """
    Replace a {{FIO_REF}} placeholder with a legacy OLE-embedded Excel file.
    Use this for older Word versions or as a fallback.
    """
    docx_path = Path(docx_path)
    if output_path is None:
        output_path = docx_path.parent / f"{docx_path.stem}_with_attachment.docx"
    else:
        output_path = Path(output_path)

    _log("=" * 60)
    _log("EMBED EXCEL (LEGACY OLE) — START")
    _log(f"  Source:      {docx_path}")
    _log(f"  Output:      {output_path}")
    _log(f"  Placeholder: {placeholder!r}")
    _log(f"  XLSX bytes:  {len(xlsx_bytes):,}")
    _log(f"  XLSX name:   {xlsx_filename}")

    # 1. Extract DOCX
    temp_dir = docx_path.parent / f"_temp_ole_{docx_path.stem}"
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            z.extractall(temp_dir)
        _log(f"  [1] Extracted DOCX → {temp_dir}")
    except Exception as e:
        raise ValueError(f"Failed to extract DOCX: {e}")

    # 2. Load document.xml
    doc_xml_path = temp_dir / "word" / "document.xml"
    try:
        doc_xml = doc_xml_path.read_text(encoding="utf-8")
    except Exception as e:
        raise ValueError(f"Failed to read document.xml: {e}")

    _log(f"  [2] document.xml size: {len(doc_xml):,} chars")
    _log(f"  [2] Placeholder found in XML: {placeholder in doc_xml}")

    if placeholder not in doc_xml:
        if placeholder.upper() in doc_xml.upper():
            _log("  [2] ⚠️ Found case-insensitive match")
            idx = doc_xml.upper().find(placeholder.upper())
            actual = doc_xml[idx:idx + len(placeholder)]
            _log(f"  [2] ⚠️ Actual text: {actual!r}")
            placeholder = actual
        else:
            raise ValueError(
                f"Placeholder '{placeholder}' not found in document.xml."
            )

    # 3. Prepare embeddings folder
    embeddings_dir = temp_dir / "word" / "embeddings"
    embeddings_dir.mkdir(parents=True, exist_ok=True)

    existing_ole = list(embeddings_dir.glob("oleObject*.bin"))
    ole_index = len(existing_ole) + 1
    ole_filename = f"oleObject{ole_index}.bin"
    ole_path = embeddings_dir / ole_filename

    # 4. Write the OLE object
    with open(ole_path, "wb") as f:
        f.write(xlsx_bytes)
    _log(f"  [4] Wrote OLE object: {ole_filename} ({len(xlsx_bytes):,} bytes)")

    # 5. Create Excel icon
    icon_path, icon_filename = _ensure_excel_icon(temp_dir, ole_index)
    _log(f"  [5] Icon ready: {icon_filename}")

    # 6. Add relationships
    rels_path = temp_dir / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")

    rids = re.findall(r'Id="rId(\d+)"', rels_xml)
    next_rid_ole = max(int(r) for r in rids) + 1 if rids else 1
    next_rid_icon = next_rid_ole + 1

    ole_rel = (
        f'<Relationship Id="rId{next_rid_ole}" '
        f'Type="{OLE_REL_TYPE}" '
        f'Target="embeddings/{ole_filename}"/>'
    )
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
    _log(f"  [6] Added rels: rId{next_rid_ole} (OLE), rId{next_rid_icon} (icon)")

    # 7. Generate unique IDs
    unique_ids = _generate_unique_ole_ids(doc_xml)

    ole_xml = _build_ole_xml(
        ole_rid=f"rId{next_rid_ole}",
        icon_rid=f"rId{next_rid_icon}",
        display_name=xlsx_filename,
        shape_id=unique_ids["shape_id"],
        object_id=unique_ids["object_id"],
        anchor_id=unique_ids["anchor_id"],
    )
    _log(f"  [7] OLE XML length: {len(ole_xml)} chars")

    # 8. Replace placeholder
    doc_xml, replaced = _replace_placeholder_with_ole(doc_xml, placeholder, ole_xml)
    _log(f"  [8] Replaced: {replaced}")

    if not replaced:
        raise ValueError(
            f"Placeholder '{placeholder}' was found but could not be replaced."
        )

    doc_xml = _ensure_required_namespaces(doc_xml)
    doc_xml_path.write_text(doc_xml, encoding="utf-8")
    _log(f"  [8] document.xml updated ({len(doc_xml):,} chars)")

    # 9. Update [Content_Types].xml
    ct_path = temp_dir / "[Content_Types].xml"
    ct_xml = ct_path.read_text(encoding="utf-8")

    if 'Extension="bin"' not in ct_xml:
        ct_xml = ct_xml.replace(
            "</Types>",
            f'<Default Extension="bin" ContentType="{OLE_CONTENT_TYPE}"/></Types>',
        )
        _log("  [9] Added .bin content type")

    if 'Extension="png"' not in ct_xml:
        ct_xml = ct_xml.replace(
            "</Types>",
            '<Default Extension="png" ContentType="image/png"/></Types>',
        )
        _log("  [9] Added .png content type")

    ct_path.write_text(ct_xml, encoding="utf-8")

    # 10. Re-zip
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(temp_dir)
                z.write(file_path, str(arcname).replace("\\", "/"))

    _log(f"  [10] Re-zipped → {output_path}")
    _log("EMBED OLE — ✅ SUCCESS")
    _log("=" * 60)

    shutil.rmtree(temp_dir, ignore_errors=True)
    return str(output_path)


def _build_ole_xml(
    ole_rid: str,
    icon_rid: str,
    display_name: str,
    shape_id: str,
    object_id: str,
    anchor_id: str,
) -> str:
    """Build legacy OLE XML."""
    safe_name = html.escape(display_name)

    return (
        '<w:r>'
        '<w:rPr><w:noProof/></w:rPr>'
        f'<w:object w:dxaOrig="1440" w:dyaOrig="1440" w14:anchorId="{anchor_id}">'
        f'<v:shape id="{shape_id}" type="#_x0000_t75" '
        'style="width:32pt;height:32pt" o:ole="">'
        f'<v:imagedata r:id="{icon_rid}" o:title="{safe_name}"/>'
        '</v:shape>'
        f'<o:OLEObject Type="Embed" ProgID="{OLE_PROG_ID}" '
        f'ShapeID="{shape_id}" DrawAspect="Icon" ObjectID="{object_id}" '
        f'r:id="{ole_rid}">'
        '<o:FieldCodes>\\s</o:FieldCodes>'
        '</o:OLEObject>'
        '</w:object>'
        '</w:r>'
        '<w:r>'
        '<w:rPr><w:noProof/></w:rPr>'
        f'<w:t xml:space="preserve"> {safe_name}</w:t>'
        '</w:r>'
    )


# ============================================================
# Icon Handling
# ============================================================
def _ensure_excel_icon(temp_dir: Path, ole_index: int) -> tuple:
    """Ensure an Excel icon exists in /word/media/."""
    media_dir = temp_dir / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)

    icon_filename = f"excel_icon_{ole_index}.png"
    icon_path = media_dir / icon_filename

    existing_icons = list(media_dir.glob("excel_icon*.png"))
    if existing_icons:
        return existing_icons[0], existing_icons[0].name

    icon_bytes = _generate_excel_icon_bytes()
    with open(icon_path, "wb") as f:
        f.write(icon_bytes)

    return icon_path, icon_filename


def _generate_excel_icon_bytes() -> bytes:
    """
    Load a real Excel icon from assets/excel_icon.png, or generate one.
    """
    # Try to load a real icon file first
    icon_file = Path(__file__).parent.parent / "assets" / "excel_icon.png"
    if icon_file.exists():
        try:
            with open(icon_file, "rb") as f:
                return f.read()
        except Exception:
            pass

    # Generate a 48×48 icon
    try:
        from PIL import Image, ImageDraw
        img = Image.new("RGBA", (48, 48), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        # Excel green rounded square
        draw.rounded_rectangle([(2, 2), (46, 46)], radius=6, fill=(33, 115, 70, 255))
        # White "X" for Excel
        draw.line([(14, 14), (34, 34)], fill="white", width=4)
        draw.line([(34, 14), (14, 34)], fill="white", width=4)

        buf = BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except ImportError:
        # Fallback: 1x1 transparent PNG
        return (
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
            b"\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
            b"\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n"
            b"\x2d\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
        )


# ============================================================
# Placeholder Replacement
# ============================================================
def _replace_placeholder_with_ole(doc_xml: str, placeholder: str, ole_xml: str) -> tuple:
    """
    Replace placeholder with OLE XML.
    Returns (updated_xml, replaced: bool).
    """
    # Strategy 1: Exact patterns
    patterns = [
        f"<w:t>{placeholder}</w:t>",
        f'<w:t xml:space="preserve">{placeholder}</w:t>',
    ]

    for pattern in patterns:
        if pattern in doc_xml:
            doc_xml = doc_xml.replace(pattern, ole_xml, 1)
            return doc_xml, True

    # Strategy 2: Regex match
    regex = re.compile(
        r"<w:t(?:\s[^>]*)?>" + re.escape(placeholder) + r"</w:t>",
        re.DOTALL,
    )
    match = regex.search(doc_xml)
    if match:
        doc_xml = doc_xml[:match.start()] + ole_xml + doc_xml[match.end():]
        return doc_xml, True

    # Strategy 3: Brute-force raw replace
    if placeholder in doc_xml:
        doc_xml = doc_xml.replace(placeholder, ole_xml, 1)
        return doc_xml, True

    return doc_xml, False


# ============================================================
# Standalone diagnostic
# ============================================================
def diagnose_placeholder(docx_path: str, placeholder: str = "{{FIO_REF}}") -> dict:
    """Standalone diagnostic tool."""
    docx_path = Path(docx_path)
    result = {
        "placeholder": placeholder,
        "exists_raw": False,
        "exists_concatenated": False,
        "context": "",
        "w_t_patterns": [],
        "run_count": 0,
        "has_ole_xml": False,
        "has_ole_bin": False,
        "has_package": False,
        "has_ole_rel": False,
        "namespaces": {},
        "ole_ids": {},
    }

    temp_dir = docx_path.parent / f"_diag_{docx_path.stem}"
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            z.extractall(temp_dir)

        doc_xml = (temp_dir / "word" / "document.xml").read_text(encoding="utf-8")
        result["exists_raw"] = placeholder in doc_xml

        for pattern_type in [
            f"<w:t>{placeholder}</w:t>",
            f'<w:t xml:space="preserve">{placeholder}</w:t>',
        ]:
            if pattern_type in doc_xml:
                result["w_t_patterns"].append(pattern_type)

        wt_runs = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", doc_xml, re.DOTALL)
        result["run_count"] = len(wt_runs)
        concatenated = "".join(wt_runs)
        result["exists_concatenated"] = placeholder in concatenated

        result["has_ole_xml"] = "<o:OLEObject" in doc_xml

        for ns in ["xmlns:o", "xmlns:v", "xmlns:w10", "xmlns:r", "xmlns:w14"]:
            result["namespaces"][ns] = f'{ns}=' in doc_xml

        embeddings = list((temp_dir / "word" / "embeddings").glob("*"))
        result["has_ole_bin"] = any(e.suffix == ".bin" for e in embeddings)
        result["has_package"] = any(e.suffix == ".xlsx" for e in embeddings)

        rels_path = temp_dir / "word" / "_rels" / "document.xml.rels"
        if rels_path.exists():
            rels_xml = rels_path.read_text(encoding="utf-8")
            result["has_ole_rel"] = ("oleObject" in rels_xml or "package" in rels_xml)

        # Check OLE ID uniqueness
        shape_ids = re.findall(r'ShapeID="([^"]+)"', doc_xml)
        object_ids = re.findall(r'ObjectID="([^"]+)"', doc_xml)
        result["ole_ids"] = {
            "shape_ids": shape_ids,
            "object_ids": object_ids,
            "shape_ids_unique": len(shape_ids) == len(set(shape_ids)),
            "object_ids_unique": len(object_ids) == len(set(object_ids)),
        }

        if placeholder in doc_xml:
            idx = doc_xml.find(placeholder)
            result["context"] = doc_xml[max(0, idx - 150):idx + 150]

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    return result


# ============================================================
# Test Utility
# ============================================================
def test_embedding(docx_path: str, xlsx_path: str, output_path: str = None,
                   use_package: bool = True) -> str:
    """
    Quick test function.
    Uses PACKAGE approach by default (recommended for Word 2016+).
    """
    with open(xlsx_path, "rb") as f:
        xlsx_bytes = f.read()

    xlsx_filename = os.path.basename(xlsx_path)

    if use_package:
        return embed_excel_as_package(
            docx_path=docx_path,
            xlsx_bytes=xlsx_bytes,
            xlsx_filename=xlsx_filename,
            placeholder="{{FIO_REF}}",
            output_path=output_path,
        )
    else:
        return embed_excel_in_docx(
            docx_path=docx_path,
            xlsx_bytes=xlsx_bytes,
            xlsx_filename=xlsx_filename,
            placeholder="{{FIO_REF}}",
            output_path=output_path,
        )
