import streamlit as st
from docx import Document
from io import BytesIO
from pathlib import Path
import tempfile
import os
import json
import re
import hashlib
import zipfile

from utils.placeholder_map import (
    PLACEHOLDER_MAP, get_groups, get_by_group,
    get_fio_keys, get_ewp_keys, get_doc_keys, get_vault_keys,
)
from utils.fio_parser import parse_fio
from utils.docx_replacer import replace_placeholders, scan_docx_for_placeholders
from utils.image_replacer import replace_ewp_image

try:
    from utils.ewp_parser import parse_ewp, get_ocr_text, extract_candidates
    EWP_OCR_AVAILABLE = True
except Exception:
    EWP_OCR_AVAILABLE = False
    parse_ewp = None
    get_ocr_text = None
    extract_candidates = None

try:
    from utils.ole_embedder import (
        embed_excel_as_package,
        embed_excel_in_docx,
        diagnose_placeholder,
    )
    OLE_EMBED_AVAILABLE = True
    PACKAGE_EMBED_AVAILABLE = True
except ImportError:
    try:
        from utils.ole_embedder import embed_excel_in_docx, diagnose_placeholder
        OLE_EMBED_AVAILABLE = True
        PACKAGE_EMBED_AVAILABLE = False
        embed_excel_as_package = None
    except Exception:
        OLE_EMBED_AVAILABLE = False
        PACKAGE_EMBED_AVAILABLE = False
        embed_excel_in_docx = None
        embed_excel_as_package = None
        diagnose_placeholder = None

# ⭐ Dynamic prompt + parser
from utils.ai_prompt import build_gemini_prompt, parse_ai_response


# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="MOP Automation",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

TEMPLATE_PATH = "template/MOP_INTEGRATION_TEMPLATE.docx"


# ============================================================
# SESSION STATE
# ============================================================
SESSION_DEFAULTS = {
    "placeholder_values": {},
    "fio_uploaded": False,
    "fio_fingerprint": None,
    "ewp_uploaded": False,
    "ewp_fingerprint": None,
    "ewp_image_bytes": None,
    "ewp_ocr_text": "",
    "ewp_candidates": {},
    "fio_attachment_bytes": None,
    "fio_attachment_name": None,
    "generated_file": None,
    "ai_updated_keys": set(),
    "debug_log": [],
    "ole_debug_log": [],
    "fio_parse_debug": None,
}

for key, default in SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default


PROTECTED_SESSION_KEYS = set(SESSION_DEFAULTS.keys())


# ============================================================
# HELPERS
# ============================================================
def set_value(key, value):
    st.session_state.placeholder_values[key] = value


def get_value(key, default=""):
    return st.session_state.placeholder_values.get(key, default)


def reset_app():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def reset_fio_state():
    """Reset only FIO-related state to force re-parse."""
    st.session_state.fio_uploaded = False
    st.session_state.fio_parse_debug = None
    st.session_state.fio_fingerprint = None

    for p in PLACEHOLDER_MAP:
        if p["source"] in ("FIO", "EWP"):
            st.session_state.placeholder_values.pop(p["key"], None)


def reset_ewp_state():
    """Reset only EWP-related state to force re-parse."""
    st.session_state.ewp_uploaded = False
    st.session_state.ewp_image_bytes = None
    st.session_state.ewp_ocr_text = ""
    st.session_state.ewp_candidates = {}
    st.session_state.ewp_fingerprint = None

    for p in PLACEHOLDER_MAP:
        if p["source"] == "EWP":
            st.session_state.placeholder_values.pop(p["key"], None)


def file_fingerprint(file_obj):
    """Generate a unique fingerprint for an uploaded file."""
    if file_obj is None:
        return None
    try:
        file_obj.seek(0)
        data = file_obj.read()
        file_obj.seek(0)
        fp = f"{file_obj.name}::{len(data)}::{hashlib.md5(data[:1024]).hexdigest()[:8]}"
        return fp
    except Exception:
        return None


def clear_widget_caches():
    """Clear widget-specific cache keys, preserve session flags."""
    for k in list(st.session_state.keys()):
        if k in PROTECTED_SESSION_KEYS:
            continue

        is_widget_cache = False

        for prefix in ("fio_", "ewp_", "doc_", "ewp_dd_", "ewp_ti_"):
            if k.startswith(prefix):
                suffix = k[len(prefix):]
                if re.match(r"^[A-Z][A-Z0-9_]*$", suffix):
                    is_widget_cache = True
                    break

        if k.startswith("_last_"):
            is_widget_cache = True

        if is_widget_cache:
            del st.session_state[k]


def sync_widget_state(widget_key, external_value, default=""):
    effective_value = external_value if external_value else default
    last_key = f"_last_{widget_key}"
    last_synced_val = st.session_state.get(last_key, None)

    if widget_key not in st.session_state:
        st.session_state[widget_key] = effective_value
        st.session_state[last_key] = effective_value
        return

    if last_synced_val != effective_value:
        st.session_state[widget_key] = effective_value
        st.session_state[last_key] = effective_value


def get_missing_keys():
    missing = []
    for p in PLACEHOLDER_MAP:
        key = p["key"]
        if key == "FIO_REF":
            continue
        val = get_value(key, "")
        default = p.get("default", "")
        if not val and not default:
            missing.append(p)
    return missing


def check_placeholder_in_docx(docx_path, placeholder):
    """Check if a placeholder exists in the DOCX and how it's stored."""
    result = {
        "exists_raw": False,
        "exists_concatenated": False,
        "near_matches": [],
        "runs_with_ref": [],
    }
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")

        result["exists_raw"] = placeholder in doc_xml

        wt_runs = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", doc_xml, re.DOTALL)
        concatenated = "".join(wt_runs)
        result["exists_concatenated"] = placeholder in concatenated

        for term in ["FIO", "REF", "{FIO", "FIO_REF", "{{FIO"]:
            idx = concatenated.find(term)
            if idx >= 0:
                snippet = concatenated[max(0, idx - 30):idx + 40]
                result["near_matches"].append({
                    "term": term,
                    "snippet": snippet,
                })

        for i, run in enumerate(wt_runs):
            if "FIO" in run or "REF" in run or "{" in run:
                result["runs_with_ref"].append({
                    "index": i,
                    "text": run,
                })
    except Exception as e:
        result["error"] = str(e)

    return result


def sync_all_widgets_to_placeholders():
    """Sync widget state (doc_*, fio_*, ewp_*) into placeholder_values."""
    for p in PLACEHOLDER_MAP:
        key = p["key"]

        if st.session_state.placeholder_values.get(key):
            continue

        for prefix in ("doc_", "fio_", "ewp_"):
            widget_key = f"{prefix}{key}"
            if widget_key in st.session_state:
                widget_val = st.session_state[widget_key]
                if widget_val:
                    st.session_state.placeholder_values[key] = widget_val
                    break


def get_value_with_fallback(key, default=""):
    """Read from placeholder_values, then widget state, then PLACEHOLDER_MAP default."""
    val = st.session_state.placeholder_values.get(key, "")
    if val:
        return val

    for prefix in ("doc_", "fio_", "ewp_"):
        widget_val = st.session_state.get(f"{prefix}{key}", "")
        if widget_val:
            return widget_val

    for p in PLACEHOLDER_MAP:
        if p["key"] == key:
            default_val = p.get("default", "")
            if default_val:
                return default_val
            break

    return default


def build_output_filename():
    """
    Build the output filename:
      MOP_{SITE_NAME}_{OLT_SITE}_{OLT_PRODUCT}_Mini_OLT_Integration_{DATE_PRIMARY}_v{VERSION}.docx
    """
    sync_all_widgets_to_placeholders()

    site_name = get_value_with_fallback("SITE_NAME", "")
    if not site_name:
        an_site = get_value_with_fallback("AN_SITE", "")
        if an_site:
            site_name = re.split(r"[_\-]", an_site)[0].strip()
    if not site_name:
        site_name = "SITE"

    olt_site = get_value_with_fallback("OLT_SITE", "OLT_SITE")
    olt_product = get_value_with_fallback("OLT_PRODUCT", "MF-2")
    date_primary = get_value_with_fallback("DATE_PRIMARY", "TBD")
    version = get_value_with_fallback("VERSION", "1.0")

    def sanitize(s):
        s = str(s)
        for ch in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']:
            s = s.replace(ch, "-")
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    date_short = re.sub(r",\s*\d{4}$", "", date_primary)

    parts = [
        "MOP",
        sanitize(site_name),
        sanitize(olt_site),
        sanitize(olt_product),
        "Mini",
        "OLT",
        "Integration",
        sanitize(date_short),
        f"v{sanitize(version)}",
    ]

    return "_".join(parts) + ".docx"


# ============================================================
# HEADER
# ============================================================
st.title("📄 MOP Automation — Nokia Lightspan MF-2 OLT Integration")
st.markdown(
    """
    Upload the **FIO Excel** and **EWP Image**, verify the mapped values,
    then generate the filled **MOP Integration Template**.

    The **FIO Excel** is embedded as a Package (Word 2016+ native format)
    at `{{FIO_REF}}` in Section 13. Double-click the icon to open it in Excel.

    **Fill methods:**
    1. 📊 **Auto-FIO** — parses values from the FIO Excel
    2. 🖼️ **Auto-OCR** — extracts values from the EWP image
    3. 🤖 **AI Paste** — paste Gemini-extracted JSON
    """
)


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("⚙️ Configuration")

    st.markdown("### Step 1 — Upload Files")
    fio_file = st.file_uploader(
        "Upload FIO (.xlsx)",
        type=["xlsx"],
        key="fio_uploader",
        help="Uploading a new FIO replaces the current session data.",
    )
    ewp_image = st.file_uploader(
        "Upload EWP Image (.jpg/.png)",
        type=["jpg", "jpeg", "png"],
        key="ewp_uploader",
        help="Uploading a new EWP replaces the current session data.",
    )

    if fio_file:
        current_fp = file_fingerprint(fio_file)
        stored_fp = st.session_state.get("fio_fingerprint")

        if current_fp and current_fp != stored_fp:
            reset_fio_state()
            st.session_state.fio_fingerprint = current_fp
            st.info("🔄 New FIO detected — re-parsing...")

        if st.session_state.get("fio_attachment_name") != fio_file.name:
            try:
                fio_file.seek(0)
                fio_bytes = fio_file.read()
                st.session_state.fio_attachment_bytes = fio_bytes
                st.session_state.fio_attachment_name = fio_file.name
                fio_file.seek(0)
            except Exception as e:
                st.error(f"❌ Failed to read FIO: {e}")

    if ewp_image:
        current_fp = file_fingerprint(ewp_image)
        stored_fp = st.session_state.get("ewp_fingerprint")

        if current_fp and current_fp != stored_fp:
            reset_ewp_state()
            st.session_state.ewp_fingerprint = current_fp
            st.info("🔄 New EWP detected — re-parsing...")

    if st.session_state.get("fio_attachment_bytes"):
        size_kb = len(st.session_state.fio_attachment_bytes) / 1024
        st.caption(
            f"📎 **Attachment ready:** `{st.session_state.fio_attachment_name}` "
            f"({size_kb:.1f} KB)"
        )

    with st.expander("⚙️ Advanced — Different file for attachment", expanded=False):
        st.caption("By default, the FIO above is embedded at `{{FIO_REF}}`.")
        override_file = st.file_uploader(
            "Override (.xlsx)",
            type=["xlsx"],
            key="fio_attachment_override",
        )
        if override_file:
            try:
                st.session_state.fio_attachment_bytes = override_file.read()
                st.session_state.fio_attachment_name = override_file.name
                st.success(f"✅ Override: {override_file.name}")
            except Exception as e:
                st.error(f"❌ Failed: {e}")

    st.markdown("---")
    st.markdown("### Step 2 — Review & Edit")
    st.markdown("FIO & EWP values are auto-filled. Verify and edit as needed.")

    st.markdown("---")
    st.markdown("### Step 3 — Generate MOP")
    generate_btn = st.button(
        "🚀 Generate MOP",
        use_container_width=True,
        type="primary",
        key="generate_btn",
    )

    st.markdown("---")
    st.markdown("### 🔄 Reload")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        if st.button("♻️ Re-parse FIO", use_container_width=True, key="reparse_fio_btn"):
            reset_fio_state()
            st.success("FIO re-parse triggered")
            st.rerun()
    with col_r2:
        if st.button("♻️ Re-parse EWP", use_container_width=True, key="reparse_ewp_btn"):
            reset_ewp_state()
            st.success("EWP re-parse triggered")
            st.rerun()

    if st.button("🔄 Reset App (Full)", use_container_width=True, key="reset_btn"):
        reset_app()

    st.markdown("---")
    st.markdown("### 📌 Status")
    st.markdown(
        f"- FIO: {'✅ Loaded' if st.session_state.fio_uploaded else '❌ Not uploaded'}\n"
        f"- EWP Image: {'✅ Loaded' if st.session_state.ewp_uploaded else '❌ Not uploaded'}\n"
        f"- FIO Attachment: {'✅ Ready' if st.session_state.fio_attachment_bytes else '❌ None'}\n"
        f"- Template: {'✅ Found' if os.path.exists(TEMPLATE_PATH) else '❌ Missing'}\n"
        f"- OCR Engine: {'✅ Available' if EWP_OCR_AVAILABLE else '⚠️ Unavailable'}\n"
        f"- Package Embed: {'✅ Available' if PACKAGE_EMBED_AVAILABLE else '⚠️ Unavailable'}\n"
        f"- OLE Embed: {'✅ Available' if OLE_EMBED_AVAILABLE else '⚠️ Unavailable'}"
    )

    missing_keys = get_missing_keys()
    if missing_keys:
        st.markdown("---")
        st.markdown("### ⚠️ Missing Fields")
        st.markdown(f"**{len(missing_keys)}** fields still empty")

    if st.session_state.ai_updated_keys:
        st.markdown("---")
        st.markdown("### 🆕 AI-Updated")
        st.markdown(f"**{len(st.session_state.ai_updated_keys)}** fields updated")


# ============================================================
# PARSE FIO
# ============================================================
if fio_file and not st.session_state.fio_uploaded:
    with st.spinner("Parsing FIO..."):
        try:
            fio_file.seek(0)
            parsed = parse_fio(fio_file)

            expected_keys = [
                "OLT_SITE", "OLT_MGMT_IP", "OLT_OM_VLAN", "OLT_OM_GW",
                "VLAN_SIP", "VLAN_HSI", "VLAN_IPOE1", "VLAN_IPOE2",
                "AN_SITE", "AN_TRUNK_ID", "AN_UPLINK_PORT",
                "AG1_NODE", "AG2_NODE", "CX600_NODE",
                "AGG_NODE", "AGG_NODE_OM", "AGG_NODE_IPOE1",
                "BNG_NODE_SIP_HSI", "BNG_NODE_IPOE2",
                "DHCP_BLOCK_OM", "DHCP_GW_OM",
                "DHCP_BLOCK_SIP", "DHCP_GW_SIP",
                "CX600_VE", "CX600_VE_L2", "CX600_VE_L3", "CX600_UPLINK_PORT",
                "VCID_OM_PRIMARY", "VCID_SIP_PRIMARY", "VCID_HSI_PRIMARY",
                "VCID_IPOE1_PRIMARY", "VCID_IPOE2_PRIMARY",
                "VCID_OM_SECONDARY", "VCID_SIP_SECONDARY", "VCID_HSI_SECONDARY",
                "VCID_IPOE1_SECONDARY", "VCID_IPOE2_SECONDARY",
                "FIO_REF", "OLT_REGION", "CABINET_NAME", "SITE_NAME",
            ]
            missing_from_parse = [k for k in expected_keys if k not in parsed]

            st.session_state.fio_parse_debug = {
                "total_keys": len(parsed),
                "parsed": parsed,
                "missing_from_parse": missing_from_parse,
                "expected_count": len(expected_keys),
            }

            for k, v in parsed.items():
                st.session_state.placeholder_values[k] = v

            if not parsed.get("FIO_REF"):
                if st.session_state.get("fio_attachment_name"):
                    name = os.path.splitext(st.session_state.fio_attachment_name)[0]
                    st.session_state.placeholder_values["FIO_REF"] = name

            st.session_state.fio_uploaded = True
            clear_widget_caches()

            st.success(f"✅ FIO parsed — {len(parsed)} values mapped")

        except Exception as e:
            st.error(f"❌ Failed to parse FIO: {e}")
            st.exception(e)


# ============================================================
# STORE EWP IMAGE + AUTO-PARSE VIA OCR
# ============================================================
if ewp_image and not st.session_state.ewp_uploaded:
    try:
        ewp_image.seek(0)
        ewp_bytes = ewp_image.read()
        st.session_state.ewp_image_bytes = ewp_bytes
        st.session_state.ewp_uploaded = True

        if EWP_OCR_AVAILABLE:
            with st.spinner("🔍 Extracting values from EWP image (OCR)..."):
                try:
                    candidates = extract_candidates(ewp_bytes)
                    st.session_state.ewp_candidates = candidates

                    ewp_parsed = parse_ewp(ewp_bytes)
                    ocr_raw = get_ocr_text(ewp_bytes)
                    st.session_state.ewp_ocr_text = ocr_raw

                    applied = 0
                    for k, v in ewp_parsed.items():
                        if not st.session_state.placeholder_values.get(k):
                            st.session_state.placeholder_values[k] = v
                            applied += 1

                    clear_widget_caches()

                    st.success(
                        f"✅ EWP loaded — {len(ewp_parsed)} values extracted, "
                        f"{applied} applied, {len(candidates)} candidate sets for dropdowns"
                    )

                except Exception as ocr_err:
                    st.warning(f"⚠️ OCR failed: {ocr_err}. Use dropdowns or AI paste.")
        else:
            st.info("ℹ️ EWP loaded. OCR unavailable — use dropdowns or AI paste.")
    except Exception as e:
        st.error(f"❌ Failed to load EWP image: {e}")
        st.exception(e)


# ============================================================
# TABS
# ============================================================
tab_fio, tab_ewp, tab_ai, tab_doc, tab_preview, tab_generate, tab_ocr, tab_about = st.tabs([
    "📊 FIO-Mapped",
    "🖼️ EWP-Only",
    "🤖 AI Paste",
    "📝 Document Metadata",
    "👁️ Preview",
    "📄 Generate",
    "🔍 OCR Debug",
    "ℹ️ About",
])


# ============================================================
# TAB 1: FIO-MAPPED
# ============================================================
with tab_fio:
    st.subheader("FIO-Mapped Placeholders")
    st.caption("Auto-extracted from FIO. AI Paste and manual edits are reflected here.")

    if st.session_state.get("fio_parse_debug"):
        dbg = st.session_state.fio_parse_debug
        with st.expander(
            f"🐛 FIO Parse Debug — {dbg['total_keys']} keys (expected {dbg['expected_count']})",
            expanded=False,
        ):
            st.markdown(f"**Total keys parsed:** {dbg['total_keys']}")

            if dbg["missing_from_parse"]:
                st.warning(
                    f"⚠️ **{len(dbg['missing_from_parse'])} expected keys NOT parsed:**"
                )
                for k in dbg["missing_from_parse"]:
                    current_val = get_value(k, "")
                    marker = "🟡 has default" if current_val else "❌ empty"
                    st.code(f"{k}  →  {marker}")
            else:
                st.success("✅ All expected keys parsed")

            st.markdown("**Full parsed dict:**")
            st.json(dbg["parsed"])

    if not st.session_state.fio_uploaded:
        st.warning("⚠️ Upload FIO Excel first.")
    else:
        groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "FIO"))
        for group in groups:
            items = [p for p in PLACEHOLDER_MAP if p["source"] == "FIO" and p["group"] == group]
            with st.expander(f"📁 {group}", expanded=True):
                cols = st.columns(2)
                for i, item in enumerate(items):
                    key, label = item["key"], item["label"]
                    default = item.get("default", "")
                    current = get_value(key) or default
                    widget_key = f"fio_{key}"

                    sync_widget_state(widget_key, current, default)

                    display_label = f"🆕 {label}" if key in st.session_state.ai_updated_keys else label

                    help_text = None
                    if key == "FIO_REF":
                        help_text = "This value is used as the anchor for the embedded Excel file."
                        display_label = f"📎 {label} (embed anchor)"

                    with cols[i % 2]:
                        val = st.text_input(display_label, key=widget_key, help=help_text)
                        if val:
                            set_value(key, val)
                            st.session_state[f"_last_{widget_key}"] = val


# ============================================================
# TAB 2: EWP-ONLY
# ============================================================
with tab_ewp:
    st.subheader("EWP-Only Placeholders")
    st.caption(
        "Values are **auto-extracted via OCR**. Where multiple candidates were found, "
        "**pick from the dropdown**. AI Paste also populates these fields. "
        "Defaults are used when OCR misses a value."
    )

    if not st.session_state.ewp_uploaded:
        st.warning("⚠️ Upload EWP image first.")

    groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "EWP"))
    for group in groups:
        items = [p for p in PLACEHOLDER_MAP if p["source"] == "EWP" and p["group"] == group]
        with st.expander(f"📁 {group}", expanded=True):
            cols = st.columns(2)
            for i, item in enumerate(items):
                key, label = item["key"], item["label"]
                candidates = st.session_state.ewp_candidates.get(key, [])
                default = item.get("default", "")
                current = get_value(key) or default
                widget_key = f"ewp_{key}"

                sync_widget_state(widget_key, current, default)

                display_label = f"🆕 {label}" if key in st.session_state.ai_updated_keys else label

                with cols[i % 2]:
                    if candidates and len(candidates) > 1:
                        options = ["(none)"] + candidates
                        try:
                            default_idx = options.index(current) if current in options else 0
                        except Exception:
                            default_idx = 0

                        if st.session_state.get(widget_key) not in options:
                            st.session_state[widget_key] = options[default_idx]

                        pick = st.selectbox(
                            f"{display_label} ({len(candidates)} candidates)",
                            options=options,
                            key=widget_key,
                        )
                        new_val = "" if pick == "(none)" else pick
                    else:
                        hint = candidates[0] if candidates else ""

                        val = st.text_input(
                            display_label,
                            key=widget_key,
                            help=f"OCR: {hint}" if hint else (
                                f"Default: {default}" if default else "No OCR match — enter manually"
                            ),
                        )
                        new_val = val

                    if new_val:
                        set_value(key, new_val)
                        st.session_state[f"_last_{widget_key}"] = new_val


# ============================================================
# TAB 3: AI PASTE
# ============================================================
with tab_ai:
    st.subheader("🤖 AI-Assisted Extraction (Gemini)")
    st.caption(
        "Use **Google Gemini** to read the FIO and EWP image and output a JSON block. "
        "Paste the JSON below — the app will auto-fill all placeholders."
    )

    missing_keys = get_missing_keys()

    # ⭐ Build DYNAMIC prompt with ALL placeholders
    full_prompt = build_gemini_prompt(include_all=True)

    st.markdown("### Step 1 — Copy this prompt to Gemini")

    if missing_keys:
        # Build focused prompt for missing fields only
        focused_lines = []
        for p in missing_keys:
            key = p["key"]
            label = p.get("label", "")
            focused_lines.append(f'  "{key}": ""   // {label}')

        focused_prompt = (
            "CRITICAL OUTPUT INSTRUCTIONS:\n"
            "- Return ONLY a valid JSON object\n"
            "- Use DOUBLE quotes for all keys and string values\n"
            "- No markdown fences, no explanation, no trailing commas\n"
            "- Use null for values you cannot find\n"
            "- Output MUST start with `{` and end with `}`\n\n"
            "---\n\n"
            "You are an expert network engineer reading a Facility Implementation Order (FIO) "
            "and an Engineering Work Plan (EWP) image for a Nokia Lightspan MF-2 OLT integration.\n\n"
            f"Extract ONLY these {len(missing_keys)} fields and return them as a JSON object:\n\n"
            "{\n"
            + ",\n".join(focused_lines)
            + "\n}\n\n"
            "FIELD DESCRIPTIONS:\n"
            + "\n".join(f'- {p["key"]}: {p["label"]}' for p in missing_keys)
        )

        st.info(
            f"💡 **{len(missing_keys)} fields still empty.** "
            f"Prompt focuses on those fields. "
            f"(Total available: **{len(PLACEHOLDER_MAP)}** placeholders)"
        )
        st.code(focused_prompt, language="markdown")

        st.markdown("#### 📋 Or copy this JSON template directly")
        st.code(
            json.dumps({p["key"]: "" for p in missing_keys}, indent=2),
            language="json",
        )
    else:
        st.success(
            f"✅ All **{len(PLACEHOLDER_MAP)}** placeholders are filled!"
        )

    # Full prompt with ALL placeholders
    with st.expander(
        f"🔍 Show FULL prompt (all {len(PLACEHOLDER_MAP)} placeholders)",
        expanded=False,
    ):
        st.code(full_prompt, language="markdown")
        st.caption(
            f"This prompt includes all {len(PLACEHOLDER_MAP)} placeholders. "
            "Use this when you want Gemini to extract every field in one pass."
        )

    st.markdown("### Step 2 — Upload FIO + EWP to Gemini, then paste the response below")

    ai_response = st.text_area(
        "Paste Gemini JSON response here (must be a JSON object)",
        height=300,
        key="ai_response_textarea",
        placeholder='{"OLT_SITE": "CDO_013_GPONA_02", "OLT_OM_VLAN": "734", ...}',
    )

    col_a, col_b = st.columns([1, 1])

    with col_a:
        if st.button("✅ Apply AI Response", use_container_width=True, key="apply_ai_btn"):
            if not ai_response.strip():
                st.warning("⚠️ Paste a JSON response first.")
            else:
                try:
                    parsed = parse_ai_response(ai_response)
                    applied_keys = []
                    unknown_keys = []
                    all_known_keys = {p["key"] for p in PLACEHOLDER_MAP}

                    for k, v in parsed.items():
                        if k in all_known_keys:
                            set_value(k, v)
                            applied_keys.append(k)
                        else:
                            unknown_keys.append(k)

                    st.session_state.ai_updated_keys = set(applied_keys)
                    clear_widget_caches()

                    st.success(f"✅ Applied {len(applied_keys)} values from AI response")
                    if unknown_keys:
                        st.warning(f"⚠️ {len(unknown_keys)} keys not recognized — skipped.")
                    with st.expander("📋 Applied values", expanded=False):
                        st.json({k: parsed[k] for k in applied_keys})
                    st.rerun()
                except ValueError as e:
                    st.error(f"❌ Failed to parse AI response: {e}")
                    with st.expander("🔍 Debug — view pasted text", expanded=False):
                        st.code(ai_response[:2000], language="text")
                except Exception as e:
                    st.error(f"❌ Unexpected error: {e}")
                    st.exception(e)

    with col_b:
        if st.button("🔍 Preview AI Response", use_container_width=True, key="preview_ai_btn"):
            if not ai_response.strip():
                st.warning("⚠️ Paste a JSON response first.")
            else:
                try:
                    parsed = parse_ai_response(ai_response)
                    st.json(parsed)
                    st.caption(f"✅ {len(parsed)} keys parsed successfully")
                except ValueError as e:
                    st.error(f"❌ Failed to parse: {e}")
                    with st.expander("🔍 View raw pasted text", expanded=True):
                        st.code(ai_response[:2000], language="text")

    st.markdown("---")
    st.markdown("### 🎯 Missing Fields")

    if not missing_keys:
        st.success("✅ All placeholders are filled!")
    else:
        st.warning(f"⚠️ {len(missing_keys)} fields still empty")
        empty_df = [
            {
                "Placeholder": f"{{{{{p['key']}}}}}",
                "Label": p["label"],
                "Source": p["source"],
                "Group": p["group"],
            }
            for p in missing_keys
        ]
        st.dataframe(empty_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 🚀 Quick Actions")
    col_x, col_y = st.columns(2)

    with col_x:
        if st.button("🧹 Clear All Fields", use_container_width=True, key="clear_all_btn"):
            st.session_state.placeholder_values = {}
            st.session_state.ai_updated_keys = set()
            st.session_state.fio_uploaded = False
            st.session_state.fio_parse_debug = None
            clear_widget_caches()
            st.success("✅ Cleared all fields. Re-upload FIO to parse again.")
            st.rerun()

    with col_y:
        if st.button("🧹 Clear 🆕 Highlights", use_container_width=True, key="clear_ai_highlights"):
            st.session_state.ai_updated_keys = set()
            st.rerun()


# ============================================================
# TAB 4: DOCUMENT METADATA
# ============================================================
with tab_doc:
    st.subheader("Document Metadata & Fixed Values")
    st.caption("Sensible defaults — editable. AI Paste can override these.")

    groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "DOC"))
    for group in groups:
        items = [p for p in PLACEHOLDER_MAP if p["source"] == "DOC" and p["group"] == group]
        with st.expander(f"📁 {group}", expanded=True):
            cols = st.columns(2)
            for i, item in enumerate(items):
                key, label = item["key"], item["label"]
                default = item.get("default", "")
                current = get_value(key) or default
                widget_key = f"doc_{key}"

                sync_widget_state(widget_key, current, default)
                display_label = f"🆕 {label}" if key in st.session_state.ai_updated_keys else label

                with cols[i % 2]:
                    val = st.text_input(display_label, key=widget_key)
                    if val:
                        set_value(key, val)
                    st.session_state[f"_last_{widget_key}"] = val


# ============================================================
# TAB 5: PREVIEW
# ============================================================
with tab_preview:
    st.subheader("Preview — All Placeholder Values")

    st.markdown("### 📄 Output Filename Preview")
    st.code(build_output_filename(), language="text")

    ewp_bytes = st.session_state.get("ewp_image_bytes")
    if ewp_bytes:
        st.markdown("### 🖼️ EWP Image")
        try:
            st.image(ewp_bytes, use_container_width=True)
        except Exception as e:
            st.error(f"❌ Failed to render: {e}")

    if st.session_state.get("fio_attachment_bytes"):
        st.markdown("### 📎 FIO Attachment (Embed at `{{FIO_REF}}`)")
        st.info(
            f"**{st.session_state.fio_attachment_name}** "
            f"({len(st.session_state.fio_attachment_bytes) / 1024:.1f} KB) — "
            "will be embedded as a Package object at the `{{FIO_REF}}` placeholder."
        )

    st.markdown("### 📋 Placeholder Values")
    rows = []
    for p in PLACEHOLDER_MAP:
        key = p["key"]
        val = get_value_with_fallback(key, p.get("default", ""))
        marker = "🆕" if key in st.session_state.ai_updated_keys else ""
        if key == "FIO_REF":
            marker = "📎"
        rows.append({
            "": marker,
            "Group": p["group"],
            "Source": p["source"],
            "Placeholder": f"{{{{{key}}}}}",
            "Value": val if val else "—",
        })
    st.dataframe(rows, use_container_width=True, hide_index=True)

    total = len(PLACEHOLDER_MAP)
    filled = sum(1 for p in PLACEHOLDER_MAP if get_value_with_fallback(p["key"], p.get("default", "")))
    col1, col2, col3 = st.columns(3)
    col1.metric("Total", total)
    col2.metric("Filled", filled)
    col3.metric("Empty", total - filled)


# ============================================================
# TAB 6: GENERATE
# ============================================================
with tab_generate:
    st.subheader("Generate Filled MOP")
    ewp_bytes = st.session_state.get("ewp_image_bytes")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"❌ Template not found at `{TEMPLATE_PATH}`.")
    elif not ewp_bytes:
        st.warning("⚠️ Upload EWP image first.")
    else:
        st.success("✅ Ready to generate.")

        st.markdown("#### 📄 Output Filename")
        st.code(build_output_filename(), language="text")

        st.markdown("#### 📋 Generation Plan")
        placeholder_count = len([p for p in PLACEHOLDER_MAP if p["key"] != "FIO_REF"])
        plan = [
            f"1. Replace **{placeholder_count}** text placeholders (excludes `{{{{FIO_REF}}}}`)",
            "2. Insert EWP image into `{{ewp_image}}`",
        ]
        if st.session_state.get("fio_attachment_bytes") and PACKAGE_EMBED_AVAILABLE:
            plan.append(
                f"3. Embed **{st.session_state.fio_attachment_name}** as Package at `{{{{FIO_REF}}}}`"
            )
        elif st.session_state.get("fio_attachment_bytes") and OLE_EMBED_AVAILABLE:
            plan.append(
                f"3. Embed **{st.session_state.fio_attachment_name}** as OLE (fallback)"
            )
        elif st.session_state.get("fio_attachment_bytes"):
            plan.append("3. ⚠️ No embedder available")
        else:
            plan.append("3. ⚠️ No FIO attachment")
        st.markdown("\n".join(plan))

        with st.expander("🔍 Pre-Generation Diagnostic — Check Template", expanded=False):
            st.caption("Verify `{{FIO_REF}}` exists in the template before generating.")

            if st.button("🔍 Run Diagnostic on Template", key="diag_btn"):
                diag = check_placeholder_in_docx(TEMPLATE_PATH, "{{FIO_REF}}")
                st.json(diag)

            if st.button("🔍 Scan Template Placeholders", key="scan_btn"):
                scan = scan_docx_for_placeholders(TEMPLATE_PATH)

                col1, col2 = st.columns(2)
                with col1:
                    st.metric("✅ Valid", len(scan.get("valid", [])))
                with col2:
                    st.metric("⚠️ Broken", len(scan.get("broken", [])))

                if scan.get("broken"):
                    st.error("**Broken placeholders found:**")
                    for b in scan["broken"]:
                        st.code(f"{b['text']}  (open={b['open_braces']}, close={b['close_braces']})")

                with st.expander("📋 All valid placeholders"):
                    for v in scan.get("valid", []):
                        st.code(v["text"])

        missing_keys = get_missing_keys()
        if missing_keys:
            st.warning(f"⚠️ **{len(missing_keys)}** fields still empty.")

        embed_method = "package"
        if PACKAGE_EMBED_AVAILABLE and OLE_EMBED_AVAILABLE:
            with st.expander("⚙️ Embedding Method", expanded=False):
                embed_method = st.radio(
                    "Choose embedding method:",
                    options=["package", "ole"],
                    format_func=lambda x: {
                        "package": "📦 Package (Word 2016+ recommended)",
                        "ole": "🔗 Legacy OLE (older Word)",
                    }[x],
                    horizontal=True,
                    key="embed_method_radio",
                )

        if generate_btn:
            st.session_state.debug_log = []
            st.session_state.ole_debug_log = []
            with st.spinner("Generating MOP..."):
                step1_path = None
                step2_path = None
                step3_path = None
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp1:
                        step1_path = tmp1.name

                    mapping = {}
                    for p in PLACEHOLDER_MAP:
                        key = p["key"]
                        if key == "FIO_REF":
                            continue
                        val = get_value_with_fallback(key, p.get("default", ""))
                        if val:
                            mapping[key] = val

                    st.session_state.debug_log.append(
                        f"Mapping keys: **{len(mapping)}**"
                    )

                    replace_placeholders(TEMPLATE_PATH, mapping, step1_path)
                    st.session_state.debug_log.append(
                        f"✅ Step 1: Replaced {len(mapping)} placeholders → {step1_path}"
                    )

                    scan_result = scan_docx_for_placeholders(step1_path)
                    st.session_state.debug_log.append("--- Scan of Step 1 output ---")

                    if scan_result.get("valid"):
                        st.session_state.debug_log.append(
                            f"⚠️ Valid placeholders still present: **{len(scan_result['valid'])}**"
                        )
                        for v in scan_result["valid"][:20]:
                            in_map = v["key"] in mapping
                            st.session_state.debug_log.append(
                                f"  - `{v['text']}` (in map: {in_map})"
                            )

                    if scan_result.get("broken"):
                        st.session_state.debug_log.append(
                            f"⚠️ Broken placeholders: **{len(scan_result['broken'])}**"
                        )
                        for b in scan_result["broken"][:20]:
                            st.session_state.debug_log.append(
                                f"  - `{b['text']}` (open={b['open_braces']}, close={b['close_braces']})"
                            )

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp2:
                        step2_path = tmp2.name

                    image_stream = BytesIO(ewp_bytes)
                    replace_ewp_image(step1_path, image_stream, step2_path)
                    st.session_state.debug_log.append(
                        f"✅ Step 2: Inserted EWP image → {step2_path}"
                    )

                    final_path = step2_path
                    has_attachment = bool(st.session_state.get("fio_attachment_bytes"))

                    if has_attachment:
                        try:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp3:
                                step3_path = tmp3.name

                            if embed_method == "package" and PACKAGE_EMBED_AVAILABLE:
                                embed_excel_as_package(
                                    docx_path=step2_path,
                                    xlsx_bytes=st.session_state.fio_attachment_bytes,
                                    xlsx_filename=st.session_state.fio_attachment_name or "FIO.xlsx",
                                    placeholder="{{FIO_REF}}",
                                    output_path=step3_path,
                                )
                                st.session_state.debug_log.append(
                                    f"✅ Step 3: Package embed → {step3_path}"
                                )
                                st.info("📦 FIO Excel embedded as Package at `{{FIO_REF}}`.")

                            elif OLE_EMBED_AVAILABLE:
                                embed_excel_in_docx(
                                    docx_path=step2_path,
                                    xlsx_bytes=st.session_state.fio_attachment_bytes,
                                    xlsx_filename=st.session_state.fio_attachment_name or "FIO.xlsx",
                                    placeholder="{{FIO_REF}}",
                                    output_path=step3_path,
                                )
                                st.session_state.debug_log.append(
                                    f"✅ Step 3: OLE embed → {step3_path}"
                                )
                                st.info("🔗 FIO Excel embedded as OLE at `{{FIO_REF}}`.")
                            else:
                                raise ValueError("No embedder available")

                            final_path = step3_path

                        except ValueError as ve:
                            st.session_state.debug_log.append(f"❌ Step 3 ValueError: {ve}")
                            st.error(f"❌ Embed failed: {ve}")
                        except Exception as embed_err:
                            st.session_state.debug_log.append(
                                f"❌ Step 3 Exception: {embed_err}"
                            )
                            st.error(f"❌ Embedding failed: {embed_err}")
                            st.exception(embed_err)
                    else:
                        st.session_state.debug_log.append(
                            "⚠️ No FIO attachment — skipped embed"
                        )
                        st.warning("⚠️ No FIO attachment — `{{FIO_REF}}` stays as text.")

                    with open(final_path, "rb") as f:
                        output_bytes = f.read()

                    output_filename = build_output_filename()

                    st.session_state.generated_file = {
                        "bytes": output_bytes,
                        "filename": output_filename,
                    }
                    st.session_state.debug_log.append(
                        f"✅ Final: {output_filename} ({len(output_bytes)} bytes)"
                    )
                    st.success(f"✅ MOP generated: `{output_filename}`")

                except Exception as e:
                    st.session_state.debug_log.append(f"❌ Generation failed: {e}")
                    st.error(f"❌ Generation failed: {e}")
                    st.exception(e)
                finally:
                    for path in [step1_path, step2_path, step3_path]:
                        if path and os.path.exists(path):
                            try:
                                os.unlink(path)
                            except Exception:
                                pass

        if st.session_state.get("debug_log"):
            st.markdown("---")
            st.markdown("### 🔍 Debug — Generation Steps")
            for line in st.session_state.debug_log:
                st.markdown(f"- {line}")

        if st.session_state.get("generated_file"):
            gen = st.session_state.generated_file
            st.markdown("---")
            st.markdown("### ⬇️ Download")
            st.download_button(
                label="⬇️ Download Filled MOP (.docx)",
                data=gen["bytes"],
                file_name=gen["filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="download_btn",
            )
            st.caption(f"File: `{gen['filename']}` — {len(gen['bytes']) / 1024:.1f} KB")

            if st.button("🔍 Verify Embedding", key="verify_embed_btn"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    tmp.write(gen["bytes"])
                    tmp_path = tmp.name

                try:
                    with zipfile.ZipFile(tmp_path, "r") as z:
                        names = z.namelist()

                        col1, col2 = st.columns(2)

                        with col1:
                            st.markdown("**Files in DOCX:**")
                            has_bin = any("oleObject" in n and n.endswith(".bin") for n in names)
                            has_pkg = any("oleObject" in n and n.endswith(".xlsx") for n in names)
                            has_icon = any("excel_icon" in n for n in names)
                            st.markdown(f"- `oleObject*.bin`: {'✅' if has_bin else '❌'}")
                            st.markdown(f"- `oleObject*.xlsx`: {'✅' if has_pkg else '❌'}")
                            st.markdown(f"- `excel_icon*.png`: {'✅' if has_icon else '❌'}")

                        with col2:
                            st.markdown("**XML Checks:**")
                            doc_xml = z.read("word/document.xml").decode("utf-8")

                            has_ole = "<o:OLEObject" in doc_xml
                            has_ns_o = 'xmlns:o=' in doc_xml
                            has_ns_v = 'xmlns:v=' in doc_xml

                            st.markdown(f"- `<o:OLEObject>`: {'✅' if has_ole else '❌'}")
                            st.markdown(f"- `xmlns:o`: {'✅' if has_ns_o else '❌'}")
                            st.markdown(f"- `xmlns:v`: {'✅' if has_ns_v else '❌'}")

                            rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
                            has_ole_rel = "oleObject" in rels
                            has_pkg_rel = "package" in rels
                            st.markdown(f"- OLE relationship: {'✅' if has_ole_rel else '❌'}")
                            st.markdown(f"- Package relationship: {'✅' if has_pkg_rel else '❌'}")

                            wt_runs = re.findall(
                                r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", doc_xml, re.DOTALL
                            )
                            concatenated = "".join(wt_runs)
                            remaining = re.findall(
                                r"\{\{[A-Z_][A-Z0-9_]*\}\}", concatenated
                            )
                            if remaining:
                                st.warning(f"⚠️ {len(remaining)} placeholders remain")
                            else:
                                st.success("✅ No placeholders remaining")

                        with st.expander("📄 Full file list"):
                            for n in sorted(names):
                                st.code(n)

                finally:
                    os.unlink(tmp_path)


# ============================================================
# TAB 7: OCR DEBUG
# ============================================================
with tab_ocr:
    st.subheader("🔍 OCR Debug")
    if not EWP_OCR_AVAILABLE:
        st.warning("⚠️ Tesseract not installed.")
    elif not st.session_state.get("ewp_uploaded"):
        st.info("ℹ️ Upload EWP image first.")
    else:
        col1, col2 = st.columns([1, 1])
        with col1:
            st.markdown("### 🖼️ EWP Image")
            st.image(st.session_state.ewp_image_bytes, use_container_width=True)
        with col2:
            st.markdown("### 📝 Raw OCR Text")
            st.text_area(
                "text",
                value=st.session_state.get("ewp_ocr_text", ""),
                height=500,
                label_visibility="collapsed",
            )

        st.markdown("### 🎯 Candidate Matches (for dropdowns)")
        candidates = st.session_state.get("ewp_candidates", {})
        if candidates:
            rows = [
                {
                    "Placeholder": f"{{{{{k}}}}}",
                    "Candidates": ", ".join(v) if v else "—",
                }
                for k, v in candidates.items()
            ]
            st.dataframe(rows, use_container_width=True, hide_index=True)


# ============================================================
# TAB 8: ABOUT
# ============================================================
with tab_about:
    st.title("ℹ️ About MOP Automation")
    st.caption("Built for Globe Telecom · Nokia Lightspan MF-2 OLT Integration · Version 1.0")

    st.markdown("---")

    st.header("👨‍💻 Developer")
    st.markdown(
        """
        **Engr. John Carlo Rabanes, ECE**  
        Electronics and Communications Engineer  
        📧 [rabanes.johncarlo4@gmail.com](mailto:rabanes.johncarlo4@gmail.com)
        """
    )

    st.markdown("---")

    mission_col, vision_col = st.columns(2)

    with mission_col:
        st.subheader("🎯 Mission")
        st.markdown(
            """
            To **eliminate repetitive, error-prone manual work** in network
            implementation documentation by automating the generation of
            Method of Procedure (MOP) documents — allowing engineers to focus
            on **actual deployment and troubleshooting** instead of copy-paste
            and cross-referencing spreadsheets.
            """
        )

    with vision_col:
        st.subheader("🌟 Vision")
        st.markdown(
            """
            A future where **every MOP, FIO, and network integration document**
            is generated in seconds with **zero manual errors**, standardized
            formatting, and full traceability from source data (FIO + EWP)
            to final deliverable.
            """
        )

    st.markdown("---")

    st.header("🚀 Features")
    st.markdown(
        """
        - **📊 Auto-FIO Parsing** — Reads `fio.xlsx`, extracts OLT/AN sites, IPs, VLANs, VSI names, VC IDs, peer IPs, AGG/BNG nodes, and port mappings.
        - **🖼️ Auto-OCR from EWP** — Tesseract OCR reads the EWP topology image, extracts node names, IPs, and ports with dropdown selection for ambiguous values.
        - **🤖 AI-Assisted Extraction** — Paste Gemini's JSON response with a focused prompt for missing fields. Robust parser handles single quotes, markdown fences, trailing commas, and multiple JSON blocks.
        - **📎 Excel Attachment Embedding** — Embeds the FIO `.xlsx` as a Package (Word 2016+ native) with OLE fallback. Double-click the icon to open in Excel.
        - **📄 DOCX Generation** — Replaces 135+ placeholders, inserts EWP image, handles split-run placeholders and Word auto-wrap issues.
        - **🔍 Diagnostics & Verification** — Pre-flight template scan, post-generation verify button, OLE debug log, OCR raw text viewer, 🆕 AI-updated highlights.
        """
    )

    st.markdown("---")

    st.header("⚡ Advantages Over Manual MOP Creation")
    st.markdown(
        """
        | Aspect | Manual MOP | MOP Automation |
        |--------|-----------|----------------|
        | **Time per MOP** | 4–8 hours | **2–5 minutes** |
        | **Error rate** | High (typos, stale values) | **Near-zero** (single source of truth) |
        | **Consistency** | Varies per engineer | **100% standardized** |
        | **Placeholder errors** | Frequent | **Auto-validated** |
        | **FIO cross-reference** | Manual lookup | **Auto-parsed** |
        | **EWP data extraction** | Manual typing | **OCR + AI-assisted** |
        | **Excel attachment** | Manual re-embed | **Auto Package/OLE** |
        | **Traceability** | Lost across versions | **Full log per generation** |
        | **Review effort** | High | **Low** (only flagged items) |
        | **Revision cost** | Full redo | **Regenerate in seconds** |
        | **Knowledge transfer** | Person-dependent | **Template-driven** |
        | **Multi-site rollout** | New MOP from scratch | **Reuse with new FIO/EWP** |
        """
    )

    st.markdown("---")

    st.header("🔄 How It Works")
    st.code(
        """
1. Upload FIO Excel  ─────►  Auto-parse ~65 values
                             │
2. Upload EWP Image  ─────►  OCR extract ~35 values
                             │
3. (Optional) AI Paste ───►  Fill missing fields via Gemini
                             │
4. Review tabs       ─────►  FIO / EWP / Metadata / Preview
                             │
5. Click Generate    ─────►  135+ placeholders replaced
                             + EWP image inserted
                             + FIO Excel embedded
                             │
6. Download .docx    ─────►  Ready for SICE review
        """,
        language="text",
    )

    st.markdown("---")

    st.header("🛠️ Built With")

    tech_col1, tech_col2, tech_col3 = st.columns(3)

    with tech_col1:
        st.markdown("**Frontend**")
        st.markdown("- Streamlit\n- Custom CSS")

    with tech_col2:
        st.markdown("**Document Processing**")
        st.markdown("- `python-docx`\n- `openpyxl`\n- Raw OOXML")

    with tech_col3:
        st.markdown("**AI & OCR**")
        st.markdown("- Tesseract OCR\n- Gemini (via prompt)\n- `Pillow`")

    st.markdown("---")

    with st.expander("⚠️ Disclaimer", expanded=False):
        st.markdown(
            """
            This tool is **for internal use only** by Globe Telecom network engineers.

            - Automatically generated MOPs **must still be reviewed by SICE** before deployment.
            - The tool does **not** deploy any configuration — it only produces documentation.
            - Always verify generated values against the source FIO/EWP.
            - Credentials are never stored — use the vault for sensitive data.

            **© 2026 Engr. John Carlo Rabanes, ECE**
            """
        )

    st.markdown("---")

    st.markdown(
        """
        <div style="text-align: center; padding: 20px;">
            <p style="color: #666;">
                Questions, bug reports, or feature requests?<br>
                📧 <a href="mailto:rabanes.johncarlo4@gmail.com">rabanes.johncarlo4@gmail.com</a>
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption("MOP Automation • Nokia Lightspan MF-2 OLT Integration • Globe Telecom")
