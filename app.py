import streamlit as st
from docx import Document
from io import BytesIO
from pathlib import Path
import tempfile
import os
import json
import re

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

from utils.ai_prompt import (
    GEMINI_PROMPT_TEMPLATE, parse_ai_response
)

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
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
if "ai_paste_result" not in st.session_state:
    st.session_state.ai_paste_result = None

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

# ============================================================
# HEADER
# ============================================================
st.title("📄 MOP Automation — Nokia Lightspan MF-2 OLT Integration")
st.markdown(
    """
    Upload the **FIO Excel** and **EWP Image**, verify the mapped values,
    then generate the filled **MOP Integration Template**.

    **Three ways to fill EWP data:**
    1. 🖼️ **Auto-OCR** — extracts values from the EWP image
    2. 📋 **Dropdown** — pick from detected candidates
    3. 🤖 **AI Paste** — use Gemini to extract, paste JSON here
    """
)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("⚙️ Configuration")

    st.markdown("### Step 1 — Upload Files")
    fio_file = st.file_uploader("Upload FIO (.xlsx)", type=["xlsx"], key="fio_uploader")
    ewp_image = st.file_uploader("Upload EWP Image (.jpg/.png)", type=["jpg", "jpeg", "png"], key="ewp_uploader")

    st.markdown("---")
    st.markdown("### Step 2 — Review & Edit")
    st.markdown("FIO & EWP values are auto-filled. Verify and edit as needed.")

    st.markdown("---")
    st.markdown("### Step 3 — Generate MOP")
    generate_btn = st.button("🚀 Generate MOP", use_container_width=True, type="primary", key="generate_btn")

    st.markdown("---")
    if st.button("🔄 Reset App", use_container_width=True, key="reset_btn"):
        reset_app()

    st.markdown("---")
    st.markdown("### 📌 Status")
    st.markdown(
        f"- FIO: {'✅ Loaded' if st.session_state.fio_uploaded else '❌ Not uploaded'}\n"
        f"- EWP Image: {'✅ Loaded' if st.session_state.ewp_uploaded else '❌ Not uploaded'}\n"
        f"- Template: {'✅ Found' if os.path.exists(TEMPLATE_PATH) else '❌ Missing'}\n"
        f"- OCR Engine: {'✅ Available' if EWP_OCR_AVAILABLE else '⚠️ Unavailable'}"
    )

# ============================================================
# PARSE FIO
# ============================================================
if fio_file and not st.session_state.fio_uploaded:
    with st.spinner("Parsing FIO..."):
        try:
            parsed = parse_fio(fio_file)
            for k, v in parsed.items():
                set_value(k, v)
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
                    # Extract candidates (all matches, for dropdowns)
                    candidates = extract_candidates(ewp_bytes)
                    st.session_state.ewp_candidates = candidates

                    # Extract single best value
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
    "🖼️ EWP-Only (Dropdown)",
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
    st.caption("Auto-extracted from FIO. Editable.")

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
                    with cols[i % 2]:
                        val = st.text_input(label, value=get_value(key), key=f"fio_{key}")
                        if val != get_value(key):
                            set_value(key, val)

# ============================================================
# TAB 2: EWP-ONLY (DROPDOWN)
# ============================================================
with tab_ewp:
    st.subheader("EWP-Only Placeholders")
    st.caption(
        "Values are **auto-extracted via OCR**. Where multiple candidates were found, "
        "**pick from the dropdown**. You can also override manually."
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
                current = get_value(key)

                with cols[i % 2]:
                    if candidates and len(candidates) > 1:
                        # Dropdown mode
                        options = ["(none)"] + candidates
                        default_idx = options.index(current) if current in options else 0
                        pick = st.selectbox(
                            f"{label} ({len(candidates)} candidates)",
                            options=options,
                            index=default_idx,
                            key=f"ewp_dd_{key}",
                        )
                        new_val = "" if pick == "(none)" else pick
                    else:
                        # Text input with candidate hint
                        hint = candidates[0] if candidates else ""
                        new_val = st.text_input(
                            label,
                            value=current or hint,
                            key=f"ewp_ti_{key}",
                            help=f"OCR detected: {hint}" if hint else "No OCR match",
                        )
                    if new_val != current:
                        set_value(key, new_val)

# ============================================================
# TAB 3: AI PASTE
# ============================================================
with tab_ai:
    st.subheader("🤖 AI-Assisted Extraction (Gemini)")
    st.caption(
        "Use **Google Gemini** to read the FIO and EWP image and output a JSON block. "
        "Paste the JSON below — the app will auto-fill all placeholders."
    )

    st.markdown("### Step 1 — Copy this prompt to Gemini")
    st.code(GEMINI_PROMPT_TEMPLATE, language="markdown")

    st.markdown("### Step 2 — Upload FIO + EWP to Gemini and paste the response below")

    ai_response = st.text_area(
        "Paste Gemini JSON response here",
        height=300,
        key="ai_response_textarea",
        placeholder='{"OLT_SITE": "CDO_013_GPONA_02", ...}',
    )

    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button("✅ Apply AI Response", use_container_width=True, key="apply_ai_btn"):
            if not ai_response.strip():
                st.warning("⚠️ Paste a JSON response first.")
            else:
                try:
                    parsed = parse_ai_response(ai_response)
                    applied = 0
                    for k, v in parsed.items():
                        set_value(k, v)
                        applied += 1
                    st.success(f"✅ Applied {applied} values from AI response")
                except Exception as e:
                    st.error(f"❌ Failed to parse AI response: {e}")
                    st.exception(e)

    with col_b:
        if st.button("🔍 Preview AI Response", use_container_width=True, key="preview_ai_btn"):
            if not ai_response.strip():
                st.warning("⚠️ Paste a JSON response first.")
            else:
                try:
                    parsed = parse_ai_response(ai_response)
                    st.json(parsed)
                except Exception as e:
                    st.error(f"❌ Failed to parse: {e}")

    st.markdown("---")
    st.markdown("### 📋 Full List of Expected Data")
    st.caption("Ensure Gemini returns ALL these fields:")
    expected_keys = [p["key"] for p in PLACEHOLDER_MAP if p["source"] in ("FIO", "EWP")]
    st.code(json.dumps({k: "..." for k in expected_keys}, indent=2), language="json")

# ============================================================
# TAB 4: DOCUMENT METADATA
# ============================================================
with tab_doc:
    st.subheader("Document Metadata & Fixed Values")
    st.caption("Sensible defaults — editable.")

    groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "DOC"))
    for group in groups:
        items = [p for p in PLACEHOLDER_MAP if p["source"] == "DOC" and p["group"] == group]
        with st.expander(f"📁 {group}", expanded=True):
            cols = st.columns(2)
            for i, item in enumerate(items):
                key, label = item["key"], item["label"]
                default = item.get("default", "")
                with cols[i % 2]:
                    val = st.text_input(label, value=get_value(key, default), key=f"doc_{key}")
                    if val != get_value(key, default):
                        set_value(key, val)

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

    st.markdown("### 📋 Placeholder Values")
    rows = []
    for p in PLACEHOLDER_MAP:
        key = p["key"]
        val = get_value(key, p.get("default", ""))
        rows.append({
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

        if generate_btn:
            with st.spinner("Generating MOP..."):
                step1_path = None
                step2_path = None
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp1:
                        step1_path = tmp1.name
                    mapping = {}
                    for p in PLACEHOLDER_MAP:
                        key = p["key"]
                        val = get_value(key, p.get("default", ""))
                        if val:
                            mapping[key] = val
                    replace_placeholders(TEMPLATE_PATH, mapping, step1_path)

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp2:
                        step2_path = tmp2.name
                    image_stream = BytesIO(ewp_bytes)
                    replace_ewp_image(step1_path, image_stream, step2_path)

                    with open(step2_path, "rb") as f:
                        output_bytes = f.read()

                    output_filename = f"MOP_INTEGRATION_{get_value('OLT_SITE', 'OUTPUT') or 'OUTPUT'}.docx"
                    st.session_state.generated_file = {
                        "bytes": output_bytes,
                        "filename": output_filename,
                    }
                    st.success("✅ MOP generated!")
                except Exception as e:
                    st.error(f"❌ Generation failed: {e}")
                    st.exception(e)
                finally:
                    for path in [step1_path, step2_path]:
                        if path and os.path.exists(path):
                            try: os.unlink(path)
                            except Exception: pass

        if st.session_state.get("generated_file"):
            gen = st.session_state.generated_file
            st.download_button(
                label="⬇️ Download Filled MOP (.docx)",
                data=gen["bytes"],
                file_name=gen["filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="download_btn",
            )

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
            st.text_area("text", value=st.session_state.get("ewp_ocr_text", ""), height=500, label_visibility="collapsed")

        st.markdown("### 🎯 Candidate Matches (for dropdowns)")
        candidates = st.session_state.get("ewp_candidates", {})
        if candidates:
            rows = [{"Placeholder": f"{{{{{k}}}}}", "Candidates": ", ".join(v) if v else "—"} for k, v in candidates.items()]
            st.dataframe(rows, use_container_width=True, hide_index=True)

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption("MOP Automation • Nokia Lightspan MF-2 OLT Integration • Globe Telecom")
