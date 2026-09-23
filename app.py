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
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
if "ai_updated_keys" not in st.session_state:
    st.session_state.ai_updated_keys = set()
if "force_sync" not in st.session_state:
    st.session_state.force_sync = False


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


def sync_widget_state(widget_key, external_value):
    """
    Sync a widget's state with an external value from placeholder_values.
    Only overwrites the widget if the external value has changed
    since the last sync (tracked via `_last_{widget_key}`).
    """
    last_key = f"_last_{widget_key}"
    current_widget_val = st.session_state.get(widget_key, None)
    last_synced_val = st.session_state.get(last_key, None)

    # If widget doesn't exist yet → initialize it
    if widget_key not in st.session_state:
        st.session_state[widget_key] = external_value
        st.session_state[last_key] = external_value
        return

    # If external value changed since last sync → push it into widget
    if last_synced_val != external_value:
        st.session_state[widget_key] = external_value
        st.session_state[last_key] = external_value


# ============================================================
# HEADER
# ============================================================
st.title("📄 MOP Automation — Nokia Lightspan MF-2 OLT Integration")
st.markdown(
    """
    Upload the **FIO Excel** and **EWP Image**, verify the mapped values,
    then generate the filled **MOP Integration Template**.

    **Three ways to fill data:**
    1. 📊 **Auto-FIO** — extracts values from the FIO Excel
    2. 🖼️ **Auto-OCR** — extracts values from the EWP image
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
    ewp_image = st.file_uploader(
        "Upload EWP Image (.jpg/.png)",
        type=["jpg", "jpeg", "png"],
        key="ewp_uploader",
    )

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
        f"- Template: {'✅ Found' if os.path.exists(TEMPLATE_PATH) else '❌ Missing'}\n"
        f"- OCR Engine: {'✅ Available' if EWP_OCR_AVAILABLE else '⚠️ Unavailable'}"
    )

    if st.session_state.ai_updated_keys:
        st.markdown("---")
        st.markdown("### 🆕 AI-Updated Fields")
        st.markdown(f"**{len(st.session_state.ai_updated_keys)}** fields updated via AI Paste")


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
                    current = get_value(key)
                    widget_key = f"fio_{key}"

                    # Sync widget with external updates (AI Paste)
                    sync_widget_state(widget_key, current)

                    display_label = f"🆕 {label}" if key in st.session_state.ai_updated_keys else label

                    with cols[i % 2]:
                        val = st.text_input(display_label, key=widget_key)
                        if val != current:
                            set_value(key, val)
                            # Track manual edit as "last synced"
                            st.session_state[f"_last_{widget_key}"] = val


# ============================================================
# TAB 2: EWP-ONLY (DROPDOWN + TEXT)
# ============================================================
with tab_ewp:
    st.subheader("EWP-Only Placeholders")
    st.caption(
        "Values are **auto-extracted via OCR**. Where multiple candidates were found, "
        "**pick from the dropdown**. AI Paste also populates these fields."
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
                widget_key = f"ewp_{key}"

                # Sync widget with external updates
                sync_widget_state(widget_key, current)

                display_label = f"🆕 {label}" if key in st.session_state.ai_updated_keys else label

                with cols[i % 2]:
                    if candidates and len(candidates) > 1:
                        # Dropdown mode
                        options = ["(none)"] + candidates
                        try:
                            default_idx = options.index(current) if current in options else 0
                        except Exception:
                            default_idx = 0

                        # Ensure widget state is one of the options
                        if st.session_state.get(widget_key) not in options:
                            st.session_state[widget_key] = options[default_idx]

                        pick = st.selectbox(
                            f"{display_label} ({len(candidates)} candidates)",
                            options=options,
                            key=widget_key,
                        )
                        new_val = "" if pick == "(none)" else pick
                    else:
                        # Text input mode
                        hint = candidates[0] if candidates else ""

                        val = st.text_input(
                            display_label,
                            key=widget_key,
                            help=f"OCR detected: {hint}" if hint else "No OCR match",
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

    st.markdown("### Step 1 — Copy this prompt to Gemini")
    st.code(GEMINI_PROMPT_TEMPLATE, language="markdown")

    st.markdown("### Step 2 — Upload FIO + EWP to Gemini, then paste the response below")

    ai_response = st.text_area(
        "Paste Gemini JSON response here (must be a JSON object like `{\"key\": \"value\"}`)",
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

                    # Apply each value — only if key is in our placeholder map
                    applied_keys = []
                    unknown_keys = []
                    all_known_keys = {p["key"] for p in PLACEHOLDER_MAP}

                    for k, v in parsed.items():
                        if k in all_known_keys:
                            set_value(k, v)
                            applied_keys.append(k)

                            # Clear widget cache so all tabs re-read the new value
                            for prefix in ("fio_", "ewp_", "doc_", "ewp_dd_", "ewp_ti_"):
                                ck = f"{prefix}{k}"
                                if ck in st.session_state:
                                    del st.session_state[ck]
                                lk = f"_last_{ck}"
                                if lk in st.session_state:
                                    del st.session_state[lk]
                        else:
                            unknown_keys.append(k)

                    # Track AI-updated keys for visual highlight
                    st.session_state.ai_updated_keys = set(applied_keys)

                    st.success(f"✅ Applied {len(applied_keys)} values from AI response")

                    if unknown_keys:
                        st.warning(
                            f"⚠️ {len(unknown_keys)} keys from AI response were "
                            f"not recognized and were skipped."
                        )

                    with st.expander("📋 Applied values", expanded=False):
                        st.json({k: parsed[k] for k in applied_keys})

                    if unknown_keys:
                        with st.expander(f"❌ Unknown keys ({len(unknown_keys)})", expanded=False):
                            st.write(unknown_keys)

                    # Force a rerun so all tabs refresh with new values
                    st.rerun()

                except ValueError as e:
                    st.error(f"❌ Failed to parse AI response: {e}")
                    st.info(
                        "💡 **Tip:** Make sure Gemini returned a JSON block like:\n"
                        "```json\n"
                        '{\n  "OLT_SITE": "CDO_013_GPONA_02",\n  "OLT_OM_VLAN": "734"\n}\n'
                        "```\n"
                        "If it returned plain text, Python code, or a different format, "
                        "ask Gemini to re-run using the prompt from Step 1."
                    )
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
                except Exception as e:
                    st.error(f"❌ Unexpected error: {e}")

    if st.session_state.ai_updated_keys:
        if st.button("🧹 Clear AI Highlights", key="clear_ai_highlights"):
            st.session_state.ai_updated_keys = set()
            st.rerun()

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
    st.caption("Sensible defaults — editable. AI Paste can override these.")

    groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "DOC"))
    for group in groups:
        items = [p for p in PLACEHOLDER_MAP if p["source"] == "DOC" and p["group"] == group]
        with st.expander(f"📁 {group}", expanded=True):
            cols = st.columns(2)
            for i, item in enumerate(items):
                key, label = item["key"], item["label"]
                default = item.get("default", "")
                current = get_value(key, default)
                widget_key = f"doc_{key}"

                # Sync widget with external updates
                sync_widget_state(widget_key, current)

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

    st.markdown("### 📋 Placeholder Values")
    rows = []
    for p in PLACEHOLDER_MAP:
        key = p["key"]
        val = get_value(key, p.get("default", ""))
        marker = "🆕" if key in st.session_state.ai_updated_keys else ""
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
                            try:
                                os.unlink(path)
                            except Exception:
                                pass

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
