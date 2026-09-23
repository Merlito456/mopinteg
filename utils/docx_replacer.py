"""
Replace {{PLACEHOLDER}} tokens in a .docx file while preserving formatting.

Handles:
- Placeholders split across multiple <w:r> runs (common Word behavior)
- Placeholders in tables, headers, footers, nested tables
- Placeholders in text boxes / shapes (via raw XML)
- Whitespace variations inside {{ }}
- Extra braces ({{KEY}}, {{KEY}}}, {{{KEY}}, etc.)
"""
from docx import Document
from docx.oxml.ns import qn
import re
import zipfile


# ============================================================
# Normalization
# ============================================================
def _normalize_placeholder_text(text: str) -> str:
    """
    Normalize placeholder text before replacement.
    - {{ KEY }} → {{KEY}}
    - {{  KEY  }} → {{KEY}}
    """
    return re.sub(
        r"\{\{\s*([A-Z_][A-Z0-9_]*)\s*\}\}",
        r"{{\1}}",
        text,
    )


def _replace_in_text(text: str, mapping: dict) -> str:
    """
    Replace all placeholders in a text string.
    Handles variations: {{KEY}}, {{KEY}}}, {{{KEY}}, {KEY}}, {KEY}
    """
    if "{" not in text:
        return text

    # Normalize spaces
    text = _normalize_placeholder_text(text)

    # Replace each key with all brace variations
    for key, value in mapping.items():
        if not value:
            continue

        str_value = str(value)
        variations = [
            "{{" + key + "}}",       # correct
            "{{" + key + "}}}",      # extra closing
            "{{{" + key + "}}",      # extra opening
            "{" + key + "}}",        # single opening
            "{" + key + "}",         # single braces
            "{{" + key + "}",        # missing closing
        ]

        for variant in variations:
            if variant in text:
                text = text.replace(variant, str_value)

    return text


# ============================================================
# Paragraph replacement
# ============================================================
def _replace_in_paragraph(paragraph, mapping, debug_log=None):
    """
    Replace placeholders in a single paragraph.
    Concatenates all runs, replaces in the concatenated text, writes back.
    This handles placeholders split across runs.
    """
    full_text = "".join(run.text for run in paragraph.runs)
    if "{" not in full_text:
        return

    new_text = _replace_in_text(full_text, mapping)

    if new_text != full_text:
        if debug_log is not None:
            debug_log.append({
                "before": full_text[:120],
                "after": new_text[:120],
            })

        # Write back: clear all runs, put new text in first run
        if paragraph.runs:
            for run in paragraph.runs:
                run.text = ""
            paragraph.runs[0].text = new_text
        else:
            paragraph.add_run(new_text)
    else:
        # No change — check if there were unmatched placeholders
        if debug_log is not None and "{{" in full_text:
            found = re.findall(r"\{\{([A-Z_][A-Z0-9_]*)\}\}", full_text)
            if found:
                unmatched = [f for f in found if f not in mapping]
                if unmatched:
                    debug_log.append({
                        "warning": "Unmatched placeholders",
                        "text": full_text[:120],
                        "missing_keys": unmatched,
                    })


# ============================================================
# Table replacement
# ============================================================
def _replace_in_table(table, mapping, debug_log=None):
    """Recursively replace in tables."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _replace_in_paragraph(paragraph, mapping, debug_log)
            for nested_table in cell.tables:
                _replace_in_table(nested_table, mapping, debug_log)


# ============================================================
# Text box replacement (raw XML)
# ============================================================
def _replace_in_text_boxes(doc, mapping, debug_log=None):
    """
    Handle text boxes / shapes. python-docx doesn't expose these directly.
    We modify the raw XML.
    """
    try:
        for txbx in doc.element.body.iter(qn("w:txbxContent")):
            for para_elem in txbx.iter(qn("w:p")):
                runs = list(para_elem.iter(qn("w:t")))
                if not runs:
                    continue

                full_text = "".join(r.text or "" for r in runs)
                if "{" not in full_text:
                    continue

                new_text = _replace_in_text(full_text, mapping)

                if new_text != full_text:
                    if debug_log is not None:
                        debug_log.append({
                            "location": "text_box",
                            "before": full_text[:120],
                            "after": new_text[:120],
                        })

                    # Put everything in first run, clear others
                    runs[0].text = new_text
                    for r in runs[1:]:
                        r.text = ""
    except Exception as e:
        if debug_log is not None:
            debug_log.append({"error": f"Text box replacement failed: {e}"})


# ============================================================
# Main entry point
# ============================================================
def replace_placeholders(docx_path, mapping, output_path, debug=False):
    """
    Replace all placeholders in the DOCX file.

    Args:
        docx_path: Input DOCX path
        mapping: {placeholder_key: value}
        output_path: Output DOCX path
        debug: If True, prints a summary of all changes

    Returns:
        output_path (str)
    """
    debug_log = [] if debug else None

    doc = Document(docx_path)

    # 1. Body paragraphs
    for paragraph in doc.paragraphs:
        _replace_in_paragraph(paragraph, mapping, debug_log)

    # 2. Tables (including nested)
    for table in doc.tables:
        _replace_in_table(table, mapping, debug_log)

    # 3. Headers and footers
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for paragraph in header.paragraphs:
                    _replace_in_paragraph(paragraph, mapping, debug_log)
                for table in header.tables:
                    _replace_in_table(table, mapping, debug_log)

        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for paragraph in footer.paragraphs:
                    _replace_in_paragraph(paragraph, mapping, debug_log)
                for table in footer.tables:
                    _replace_in_table(table, mapping, debug_log)

    # 4. Text boxes / shapes
    _replace_in_text_boxes(doc, mapping, debug_log)

    # 5. Save
    doc.save(output_path)

    # Debug output
    if debug and debug_log:
        print("\n" + "=" * 60)
        print(f"REPLACER DEBUG — {len(debug_log)} changes")
        print("=" * 60)
        for entry in debug_log:
            if "warning" in entry:
                print(f"⚠️ WARNING: {entry['warning']}")
                print(f"   Text: {entry['text']}")
                print(f"   Missing keys: {entry['missing_keys']}")
            elif "error" in entry:
                print(f"❌ ERROR: {entry['error']}")
            else:
                print(f"✅ {entry.get('location', 'body')}")
                print(f"   Before: {entry['before']}")
                print(f"   After:  {entry['after']}")
        print("=" * 60 + "\n")

    return output_path


# ============================================================
# Diagnostic: scan a DOCX for placeholder patterns
# ============================================================
def scan_docx_for_placeholders(docx_path):
    """
    Scan a DOCX and return all "curly brace" patterns found.
    Useful for debugging broken placeholders.

    Returns:
        dict with keys: valid, broken, all_placeholder_strings
    """
    result = {
        "valid": [],
        "broken": [],
        "all_placeholder_strings": set(),
    }

    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
    except Exception as e:
        return {"error": str(e)}

    # Extract all <w:t> contents
    wt_runs = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", doc_xml, re.DOTALL)
    concatenated = "".join(wt_runs)

    # Find all { ... } patterns with 1-3 braces
    all_patterns = re.findall(r"(\{{1,3})([A-Z_][A-Z0-9_]*)(\}{1,3})", concatenated)

    seen = set()
    for open_b, key, close_b in all_patterns:
        signature = f"{open_b}{key}{close_b}"
        if signature in seen:
            continue
        seen.add(signature)

        entry = {
            "text": signature,
            "key": key,
            "open_braces": len(open_b),
            "close_braces": len(close_b),
        }

        if open_b == "{{" and close_b == "}}":
            result["valid"].append(entry)
        else:
            result["broken"].append(entry)

        result["all_placeholder_strings"].add(key)

    # Convert set to sorted list
    result["all_placeholder_strings"] = sorted(result["all_placeholder_strings"])

    return result
