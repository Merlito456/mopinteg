import streamlit as st
from docx import Document
from io import BytesIO
from pathlib import Path
import tempfile
import os
import json
import re
import zipfile

from utils.placeholder_map import (
    PLACEHOLDER_MAP, get_groups, get_by_group,
    get_fio_keys, get_ewp_keys, get_doc_keys, get_vault_keys,
)
from utils.fio_parser import parse_fio
from utils.docx_replacer import replace_placeholders
from utils.image_replacer import replace_ewp_image

try:
    from utils.ewp_parser import parse_ewp, get_ocr_text, extract_candidates
    EWP_OCR_AVAILABLE = True
except Exception:
    EWP_OCR_AVAILABLE = False
    parse_ewp = None
    get_ocr_text = None
    extract_candidates = None

# ⭐ Try Package approach first, fall back to legacy OLE
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

from utils.ai_prompt import GEMINI_PROMPT_TEMPLATE, parse_ai_response


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
if "placeholder_values" not in st.session_state:
    st.session_state.placeholder_values = {}
if "fio_uploaded" not in st.session_state:
    st.session_state.fio_uploaded = False
if "ewp_uploaded" not in st.session_state:
    st.session_state.ewp_uploaded = False
if "ewp_image_bytes" not in st.session_state:
    st.session_state.ewp_image_bytes = None
if "ewp_ocr_text" not in st.session_state:
    st.session_state.ewp_ocr_text = ""
if "ewp_candidates" not in st.session_state:
    st.session_state.ewp_candidates = {}
if "fio_attachment_bytes" not in st.session_state:
    st.session_state.fio_attachment_bytes = None
if "fio_attachment_name" not in st.session_state:
    st.session_state.fio_attachment_name = None
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
if "ai_updated_keys" not in st.session_state:
    st.session_state.ai_updated_keys = set()
if "debug_log" not in st.session_state:
    st.session_state.debug_log = []
if "ole_debug_log" not in st.session_state:
    st.session_state.ole_debug_log = []


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
        help="Used for both auto-parsing values AND embedding at {{FIO_REF}}.",
    )
    ewp_image = st.file_uploader(
        "Upload EWP Image (.jpg/.png)",
        type=["jpg", "jpeg", "png"],
        key="ewp_uploader",
        help="Used for OCR extraction and inserted into Section 9.",
    )

    if fio_file:
        if st.session_state.get("fio_attachment_name") != fio_file.name:
            try:
                fio_bytes = fio_file.read()
                st.session_state.fio_attachment_bytes = fio_bytes
                st.session_state.fio_attachment_name = fio_file.name
                fio_file.seek(0)
            except Exception as e:
                st.error(f"❌ Failed to read FIO: {e}")

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
    if st.button("🔄 Reset App", use_container_width=True, key="reset_btn"):
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
            for k, v in parsed.items():
                set_value(k, v)

            if not parsed.get("FIO_REF"):
                if st.session_state.get("fio_attachment_name"):
                    name = os.path.splitext(st.session_state.fio_attachment_name)[0]
                    set_value("FIO_REF", name)

            st.session_state.fio_uploaded = True
            st.success(f"✅ FIO parsed — {len(parsed)} values mapped")
        except Exception as e:
            st.error(f"❌ Failed to parse FIO: {e}")
            st.exception(e)


# ============================================================
# STORE EWP IMAGE + AUTO-PARSE VIA OCR
# ============================================================
if ewp_image and not st.session_state.ewp_uploaded:
    try:
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
                        if not get_value(k):
                            set_value(k, v)
                            applied += 1

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
tab_fio, tab_ewp, tab_ai, tab_doc, tab_preview, tab_generate, tab_ocr = st.tabs([
    "📊 FIO-Mapped",
    "🖼️ EWP-Only",
    "🤖 AI Paste",
    "📝 Document Metadata",
    "👁️ Preview",
    "📄 Generate",
    "🔍 OCR Debug",
])


# ============================================================
# TAB 1: FIO-MAPPED
# ============================================================
with tab_fio:
    st.subheader("FIO-Mapped Placeholders")
    st.caption("Auto-extracted from FIO. AI Paste and manual edits are reflected here.")

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
                        if val != current:
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

                    if new_val != current:
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

    st.markdown("### Step 1 — Copy this prompt to Gemini")

    if missing_keys:
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
            "Extract ONLY these fields and return them as a JSON object:\n\n"
            "{\n"
            + ",\n".join(f'  "{p["key"]}": ""' for p in missing_keys)
            + "\n}\n\n"
            "FIELD DESCRIPTIONS:\n"
            + "\n".join(f'- {p["key"]}: {p["label"]}' for p in missing_keys)
        )
        st.info(f"💡 **{len(missing_keys)} fields still empty.** Prompt focuses on those fields.")
        st.code(focused_prompt, language="markdown")
        st.markdown("#### 📋 Or copy this JSON template directly")
        st.code(json.dumps({p["key"]: "" for p in missing_keys}, indent=2), language="json")
    else:
        st.success("✅ All fields already filled! Full prompt shown for reference.")
        st.code(GEMINI_PROMPT_TEMPLATE, language="markdown")

    with st.expander("🔍 Show full prompt (all fields)", expanded=False):
        st.code(GEMINI_PROMPT_TEMPLATE, language="markdown")

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
                            for prefix in ("fio_", "ewp_", "doc_", "ewp_dd_", "ewp_ti_"):
                                ck = f"{prefix}{k}"
                                if ck in st.session_state:
                                    del st.session_state[ck]
                                lk = f"_last_{ck}"
                                if lk in st.session_state:
                                    del st.session_state[lk]
                        else:
                            unknown_keys.append(k)

                    st.session_state.ai_updated_keys = set(applied_keys)
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
            for k in list(st.session_state.keys()):
                if k.startswith(("fio_", "ewp_", "doc_", "_last_")):
                    del st.session_state[k]
            st.success("✅ Cleared all fields.")
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
                    if val != current:
                        set_value(key, val)
                        st.session_state[f"_last_{widget_key}"] = val


# ============================================================
# TAB 5: PREVIEW
# ============================================================
with tab_preview:
    st.subheader("Preview — All Placeholder Values")

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
        val = get_value(key, p.get("default", ""))
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
    filled = sum(1 for p in PLACEHOLDER_MAP if get_value(p["key"], p.get("default", "")))
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

        # Show generation plan
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

        # Pre-generation diagnostic
        with st.expander("🔍 Pre-Generation Diagnostic — Check Template", expanded=False):
            st.caption("Verify `{{FIO_REF}}` exists in the template before generating.")
            if st.button("🔍 Run Diagnostic on Template", key="diag_btn"):
                diag = check_placeholder_in_docx(TEMPLATE_PATH, "{{FIO_REF}}")
                st.json(diag)

        missing_keys = get_missing_keys()
        if missing_keys:
            st.warning(f"⚠️ **{len(missing_keys)}** fields still empty.")

        # Embedding method selector
        embed_method = "package"  # default
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
                    # --- Step 1: Replace text placeholders (EXCEPT FIO_REF) ---
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp1:
                        step1_path = tmp1.name

                    mapping = {}
                    for p in PLACEHOLDER_MAP:
                        key = p["key"]
                        if key == "FIO_REF":
                            continue
                        val = get_value(key, p.get("default", ""))
                        if val:
                            mapping[key] = val

                    replace_placeholders(TEMPLATE_PATH, mapping, step1_path)
                    st.session_state.debug_log.append(
                        f"✅ Step 1: Replaced {len(mapping)} placeholders → {step1_path}"
                    )

                    # --- Step 2: Insert EWP image ---
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp2:
                        step2_path = tmp2.name

                    image_stream = BytesIO(ewp_bytes)
                    replace_ewp_image(step1_path, image_stream, step2_path)
                    st.session_state.debug_log.append(
                        f"✅ Step 2: Inserted EWP image → {step2_path}"
                    )

                    # Diagnostic on step2 BEFORE embed
                    diag = check_placeholder_in_docx(step2_path, "{{FIO_REF}}")
                    st.session_state.debug_log.append("--- Diagnostic on Step 2 output ---")
                    st.session_state.debug_log.append(
                        f"Exists in raw XML: **{diag['exists_raw']}**"
                    )
                    st.session_state.debug_log.append(
                        f"Exists in concatenated: **{diag['exists_concatenated']}**"
                    )

                    # --- Step 3: Embed FIO at {{FIO_REF}} ---
                    final_path = step2_path
                    has_attachment = bool(st.session_state.get("fio_attachment_bytes"))

                    st.session_state.debug_log.append(
                        f"FIO attachment bytes: **{len(st.session_state.get('fio_attachment_bytes') or b'')}**"
                    )
                    st.session_state.debug_log.append(
                        f"PACKAGE_EMBED_AVAILABLE: **{PACKAGE_EMBED_AVAILABLE}**"
                    )
                    st.session_state.debug_log.append(
                        f"OLE_EMBED_AVAILABLE: **{OLE_EMBED_AVAILABLE}**"
                    )

                    if has_attachment:
                        try:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp3:
                                step3_path = tmp3.name

                            # Prefer Package, fall back to OLE
                            if embed_method == "package" and PACKAGE_EMBED_AVAILABLE:
                                st.session_state.debug_log.append(
                                    "🔧 Using **Package** embedding..."
                                )
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
                                st.session_state.debug_log.append(
                                    "🔧 Using **Legacy OLE** embedding..."
                                )
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
                            try:
                                diag2 = diagnose_placeholder(step2_path, "{{FIO_REF}}")
                                st.session_state.debug_log.append(
                                    f"Diagnostic: {json.dumps(diag2, indent=2)}"
                                )
                            except Exception as de:
                                st.session_state.debug_log.append(f"Diagnostic failed: {de}")
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

                    # --- Step 4: Read output ---
                    with open(final_path, "rb") as f:
                        output_bytes = f.read()

                    output_filename = (
                        f"MOP_INTEGRATION_{get_value('OLT_SITE', 'OUTPUT') or 'OUTPUT'}.docx"
                    )
                    st.session_state.generated_file = {
                        "bytes": output_bytes,
                        "filename": output_filename,
                    }
                    st.session_state.debug_log.append(
                        f"✅ Final: {output_filename} ({len(output_bytes)} bytes)"
                    )
                    st.success("✅ MOP generated successfully!")

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

        # Display debug log
        if st.session_state.get("debug_log"):
            st.markdown("---")
            st.markdown("### 🔍 Debug — Generation Steps")
            for line in st.session_state.debug_log:
                st.markdown(f"- {line}")

        # Display OLE debug log
        if st.session_state.get("ole_debug_log"):
            with st.expander("🔍 OLE/Embed Debug Log", expanded=False):
                for line in st.session_state.ole_debug_log:
                    st.code(line, language="text")

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

            # Verify button
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

                            # ID uniqueness
                            shape_ids = re.findall(r'ShapeID="([^"]+)"', doc_xml)
                            object_ids = re.findall(r'ObjectID="([^"]+)"', doc_xml)
                            shape_unique = len(shape_ids) == len(set(shape_ids))
                            object_unique = len(object_ids) == len(set(object_ids))
                            st.markdown(f"- ShapeIDs unique: {'✅' if shape_unique else '❌'}")
                            st.markdown(f"- ObjectIDs unique: {'✅' if object_unique else '❌'}")

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
# FOOTER
# ============================================================
st.markdown("---")
st.caption("MOP Automation • Nokia Lightspan MF-2 OLT Integration • Globe Telecom")
