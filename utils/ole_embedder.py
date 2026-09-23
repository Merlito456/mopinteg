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
import html
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
    """
    Ensure document.xml has all namespaces required for OLE embedding.
    Word needs: xmlns:o, xmlns:v, xmlns:w10, xmlns:r, xmlns:w
    """
    # Find the <w:document ...> opening tag
    doc_tag_match = re.search(r"<w:document\b[^>]*>", doc_xml)
    if not doc_tag_match:
        _log("  [NS] ⚠️ Could not find <w:document> tag")
        return doc_xml

    doc_tag = doc_tag_match.group(0)
    new_tag = doc_tag
    added = []

    for ns_prefix, ns_uri in REQUIRED_NAMESPACES.items():
        # Check if namespace is already declared (as xmlns:prefix or xmlns:prefix=)
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
# Main Entry Point
# ============================================================
def embed_excel_in_docx(
    docx_path: str,
    xlsx_bytes: bytes,
    xlsx_filename: str,
    placeholder: str = "{{FIO_REF}}",
    output_path: str = None,
) -> str:
    """
    Replace a {{FIO_REF}} placeholder in a DOCX with a real OLE-embedded Excel file.
    """
    docx_path = Path(docx_path)
    if output_path is None:
        output_path = docx_path.parent / f"{docx_path.stem}_with_attachment.docx"
    else:
        output_path = Path(output_path)

    _log("=" * 60)
    _log("EMBED EXCEL — START")
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

    if placeholder in doc_xml:
        idx = doc_xml.find(placeholder)
        snippet = doc_xml[max(0, idx - 100):idx + 100]
        _log(f"  [2] Context: ...{snippet}...")

    if placeholder not in doc_xml:
        if placeholder.upper() in doc_xml.upper():
            _log("  [2] ⚠️ Found case-insensitive match")
            idx = doc_xml.upper().find(placeholder.upper())
            actual = doc_xml[idx:idx + len(placeholder)]
            _log(f"  [2] ⚠️ Actual text: {actual!r}")
            placeholder = actual
        else:
            raise ValueError(
                f"Placeholder '{placeholder}' not found in document.xml. "
                "Add it to the template where the FIO attachment should appear."
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

    # 7. Build OLE XML
    ole_xml = _build_ole_xml(
        ole_rid=f"rId{next_rid_ole}",
        icon_rid=f"rId{next_rid_icon}",
        display_name=xlsx_filename,
    )
    _log(f"  [7] OLE XML length: {len(ole_xml)} chars")

    # 8. Replace placeholder
    doc_xml, replaced = _replace_placeholder_with_ole(doc_xml, placeholder, ole_xml)
    _log(f"  [8] Replaced: {replaced}")

    if not replaced:
        raise ValueError(
            f"Placeholder '{placeholder}' was found but could not be replaced. "
        )

    # ⭐ 8b. Ensure required namespaces are declared
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

    # 10. Re-zip as DOCX
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(temp_dir)
                z.write(file_path, str(arcname).replace("\\", "/"))

    _log(f"  [10] Re-zipped → {output_path}")
    _log("EMBED EXCEL — ✅ SUCCESS")
    _log("=" * 60)

    shutil.rmtree(temp_dir, ignore_errors=True)
    return str(output_path)


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
    """Generate a 48x48 Excel-green icon PNG."""
    icon_file = Path(__file__).parent.parent / "assets" / "excel_icon.png"
    if icon_file.exists():
        with open(icon_file, "rb") as f:
            return f.read()

    try:
        from PIL import Image, ImageDraw
        img = Image.new("RGBA", (48, 48), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([(2, 2), (46, 46)], radius=6, fill=(33, 115, 70, 255))
        draw.line([(14, 14), (34, 34)], fill="white", width=4)
        draw.line([(34, 14), (14, 34)], fill="white", width=4)

        buf = BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except ImportError:
        return (
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
            b"\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
            b"\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n"
            b"\x2d\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
        )


# ============================================================
# OLE Object XML Builder — Word-compatible
# ============================================================
def _build_ole_xml(ole_rid: str, icon_rid: str, display_name: str) -> str:
    """
    Build the OLE object XML matching Word's native embedded object structure.

    Key details from real Word output:
    - <w:rPr><w:noProof/></w:rPr> comes BEFORE <w:object>
    - w14:anchorId is present
    - v:shape has o:ole="" attribute
    - ShapeID matches v:shape id
    - o:FieldCodes uses \s (backslash s)
    """
    safe_name = html.escape(display_name)

    ole_xml = (
        '<w:r>'
        '<w:rPr><w:noProof/></w:rPr>'
        '<w:object w:dxaOrig="1440" w:dyaOrig="1440" w14:anchorId="4C8D24A8">'
        # ---- Icon shape (matches ShapeID below) ----
        '<v:shape id="_x0000_i1027" type="#_x0000_t75" '
        'style="width:32pt;height:32pt" o:ole="">'
        f'<v:imagedata r:id="{icon_rid}" o:title="{safe_name}"/>'
        '</v:shape>'
        # ---- OLE object reference ----
        f'<o:OLEObject Type="Embed" ProgID="{OLE_PROG_ID}" '
        'ShapeID="_x0000_i1027" DrawAspect="Icon" ObjectID="_1625429161" '
        f'r:id="{ole_rid}">'
        '<o:FieldCodes>\\s</o:FieldCodes>'
        '</o:OLEObject>'
        '</w:object>'
        '</w:r>'
        # ---- Filename text after the icon ----
        '<w:r>'
        '<w:rPr><w:noProof/></w:rPr>'
        f'<w:t xml:space="preserve"> {safe_name}</w:t>'
        '</w:r>'
    )
    return ole_xml


# ============================================================
# Placeholder Replacement — SIMPLE & ROBUST
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
        "has_ole_rel": False,
        "namespaces": {},
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

        # Check OLE presence
        result["has_ole_xml"] = "<o:OLEObject" in doc_xml

        # Check namespaces
        for ns in ["xmlns:o", "xmlns:v", "xmlns:w10", "xmlns:r", "xmlns:w14"]:
            result["namespaces"][ns] = f'{ns}=' in doc_xml

        # Check embeddings
        embeddings = list((temp_dir / "word" / "embeddings").glob("*.bin"))
        result["has_ole_bin"] = len(embeddings) > 0

        # Check rels
        rels_path = temp_dir / "word" / "_rels" / "document.xml.rels"
        if rels_path.exists():
            rels_xml = rels_path.read_text(encoding="utf-8")
            result["has_ole_rel"] = "oleObject" in rels_xml

        # Context
        if placeholder in doc_xml:
            idx = doc_xml.find(placeholder)
            result["context"] = doc_xml[max(0, idx - 150):idx + 150]

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    return result


# ============================================================
# Test Utility
# ============================================================
def test_embedding(docx_path: str, xlsx_path: str, output_path: str = None) -> str:
    """Quick test function to verify embedding works."""
    with open(xlsx_path, "rb") as f:
        xlsx_bytes = f.read()

    xlsx_filename = os.path.basename(xlsx_path)

    return embed_excel_in_docx(
        docx_path=docx_path,
        xlsx_bytes=xlsx_bytes,
        xlsx_filename=xlsx_filename,
        placeholder="{{FIO_REF}}",
        output_path=output_path,
    )
