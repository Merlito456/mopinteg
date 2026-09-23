import streamlit as st
from docx import Document
from io import BytesIO
from pathlib import Path
import tempfile
import os

from utils.placeholder_map import (
    PLACEHOLDER_MAP, get_groups, get_by_group,
    get_fio_keys, get_ewp_keys, get_doc_keys
)
from utils.fio_parser import parse_fio
from utils.docx_replacer import replace_placeholders
from utils.image_replacer import replace_ewp_image

# Optional EWP OCR parser — graceful fallback if Tesseract not available
try:
    from utils.ewp_parser import parse_ewp, get_ocr_text
    EWP_OCR_AVAILABLE = True
except Exception:
    EWP_OCR_AVAILABLE = False
    parse_ewp = None
    get_ocr_text = None


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

if "generated_file" not in st.session_state:
    st.session_state.generated_file = None


# ============================================================
# HELPERS
# ============================================================
def set_value(key, value):
    st.session_state.placeholder_values[key] = value


def get_value(key, default=""):
    return st.session_state.placeholder_values.get(key, default)


def reset_app():
    """Reset all session state."""
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

    - **FIO values** → auto-extracted from the Excel file
    - **EWP values** → auto-extracted via OCR, editable for verification
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
    )
    ewp_image = st.file_uploader(
        "Upload EWP Image (.jpg/.png)",
        type=["jpg", "jpeg", "png"],
        key="ewp_uploader",
    )

    st.markdown("---")
    st.markdown("### Step 2 — Review & Edit")
    st.markdown(
        "FIO-derived values are **auto-filled**.\n\n"
        "EWP-only values are **auto-extracted via OCR** — please verify and edit."
    )

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

    # Status indicators
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

        # --- Auto-parse EWP via OCR ---
        if EWP_OCR_AVAILABLE:
            with st.spinner("🔍 Extracting values from EWP image (OCR)..."):
                try:
                    ewp_parsed = parse_ewp(ewp_bytes)
                    ocr_raw = get_ocr_text(ewp_bytes)
                    st.session_state.ewp_ocr_text = ocr_raw

                    # Only overwrite if empty (FIO takes precedence)
                    applied = 0
                    for k, v in ewp_parsed.items():
                        if not get_value(k):
                            set_value(k, v)
                            applied += 1

                    st.success(
                        f"✅ EWP image loaded — {len(ewp_parsed)} values extracted "
                        f"({applied} applied)"
                    )

                except Exception as ocr_err:
                    st.warning(
                        f"⚠️ OCR extraction failed: {ocr_err}. "
                        "You can still enter EWP values manually."
                    )
        else:
            st.info(
                "ℹ️ EWP image loaded. OCR engine unavailable — "
                "please enter EWP values manually."
            )

    except Exception as e:
        st.error(f"❌ Failed to load EWP image: {e}")
        st.exception(e)


# ============================================================
# TABS
# ============================================================
tab_fio, tab_ewp, tab_doc, tab_preview, tab_generate, tab_ocr = st.tabs([
    "📊 FIO-Mapped (Auto)",
    "🖼️ EWP-Only (OCR + Manual)",
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
    st.caption("These values were extracted from the FIO. You may edit them below.")

    if not st.session_state.fio_uploaded:
        st.warning("⚠️ Upload the FIO Excel file first in the sidebar.")
    else:
        groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "FIO"))
        for group in groups:
            items = [p for p in PLACEHOLDER_MAP if p["source"] == "FIO" and p["group"] == group]
            with st.expander(f"📁 {group}", expanded=True):
                cols = st.columns(2)
                for i, item in enumerate(items):
                    key = item["key"]
                    label = item["label"]
                    with cols[i % 2]:
                        val = st.text_input(
                            label,
                            value=get_value(key),
                            key=f"fio_{key}",
                        )
                        if val != get_value(key):
                            set_value(key, val)


# ============================================================
# TAB 2: EWP-ONLY (OCR + MANUAL)
# ============================================================
with tab_ewp:
    st.subheader("EWP-Only Placeholders (OCR + Manual)")
    st.caption(
        "These values are **auto-extracted via OCR** from the EWP image. "
        "Please verify and edit as needed."
    )

    if not st.session_state.ewp_uploaded:
        st.warning("⚠️ Upload the EWP image first (it will be inserted into the MOP).")
    elif not EWP_OCR_AVAILABLE:
        st.info("ℹ️ OCR engine unavailable — please enter EWP values manually.")
    else:
        st.info("💡 Tip: Check the **OCR Debug** tab to verify extraction accuracy.")

    groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "EWP"))
    for group in groups:
        items = [p for p in PLACEHOLDER_MAP if p["source"] == "EWP" and p["group"] == group]
        with st.expander(f"📁 {group}", expanded=True):
            cols = st.columns(2)
            for i, item in enumerate(items):
                key = item["key"]
                label = item["label"]
                with cols[i % 2]:
                    current_val = get_value(key)
                    val = st.text_input(
                        label,
                        value=current_val,
                        key=f"ewp_{key}",
                        help="✅ Auto-filled" if current_val else "⚠️ Empty — please input",
                    )
                    if val != current_val:
                        set_value(key, val)


# ============================================================
# TAB 3: DOCUMENT METADATA
# ============================================================
with tab_doc:
    st.subheader("Document Metadata & Fixed Values")
    st.caption("These have sensible defaults but can be edited.")

    groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "DOC"))
    for group in groups:
        items = [p for p in PLACEHOLDER_MAP if p["source"] == "DOC" and p["group"] == group]
        with st.expander(f"📁 {group}", expanded=True):
            cols = st.columns(2)
            for i, item in enumerate(items):
                key = item["key"]
                label = item["label"]
                default = item.get("default", "")
                with cols[i % 2]:
                    val = st.text_input(
                        label,
                        value=get_value(key, default),
                        key=f"doc_{key}",
                    )
                    if val != get_value(key, default):
                        set_value(key, val)


# ============================================================
# TAB 4: PREVIEW
# ============================================================
with tab_preview:
    st.subheader("Preview — All Placeholder Values")
    st.caption("This is what will be substituted into the template.")

    # --- EWP Image (SAFE rendering) ---
    ewp_bytes = st.session_state.get("ewp_image_bytes")
    if ewp_bytes:
        st.markdown("### 🖼️ EWP Image")
        try:
            st.image(ewp_bytes, use_container_width=True)
        except Exception as e:
            st.error(f"❌ Failed to render EWP image: {e}")
    else:
        st.info("ℹ️ No EWP image uploaded yet. Upload one in the sidebar.")

    # --- Placeholder table ---
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

    # --- Summary metrics ---
    st.markdown("### 📈 Summary")
    total = len(PLACEHOLDER_MAP)
    filled = sum(1 for p in PLACEHOLDER_MAP if get_value(p["key"], p.get("default", "")))
    empty = total - filled

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Placeholders", total)
    col2.metric("Filled", filled)
    col3.metric("Empty", empty)


# ============================================================
# TAB 5: GENERATE
# ============================================================
with tab_generate:
    st.subheader("Generate Filled MOP")

    ewp_bytes = st.session_state.get("ewp_image_bytes")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(
            f"❌ Template not found at `{TEMPLATE_PATH}`. "
            "Please add `MOP_INTEGRATION_TEMPLATE.docx` to the `template/` folder."
        )
    elif not ewp_bytes:
        st.warning("⚠️ Please upload the EWP image first in the sidebar.")
    else:
        st.success("✅ Template found. Ready to generate.")

        if generate_btn:
            with st.spinner("Generating MOP..."):
                step1_path = None
                step2_path = None
                try:
                    # --- Step 1: Text placeholder replacement ---
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp1:
                        step1_path = tmp1.name

                    mapping = {}
                    for p in PLACEHOLDER_MAP:
                        key = p["key"]
                        val = get_value(key, p.get("default", ""))
                        if val:
                            mapping[key] = val

                    replace_placeholders(TEMPLATE_PATH, mapping, step1_path)

                    # --- Step 2: EWP image insertion ---
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp2:
                        step2_path = tmp2.name

                    image_stream = BytesIO(ewp_bytes)
                    replace_ewp_image(step1_path, image_stream, step2_path)

                    # --- Step 3: Read output ---
                    with open(step2_path, "rb") as f:
                        output_bytes = f.read()

                    output_filename = (
                        f"MOP_INTEGRATION_"
                        f"{get_value('OLT_SITE', 'OUTPUT') or 'OUTPUT'}.docx"
                    )

                    st.session_state.generated_file = {
                        "bytes": output_bytes,
                        "filename": output_filename,
                    }

                    st.success("✅ MOP generated successfully!")

                except Exception as e:
                    st.error(f"❌ Generation failed: {e}")
                    st.exception(e)

                finally:
                    # Cleanup temp files
                    for path in [step1_path, step2_path]:
                        if path and os.path.exists(path):
                            try:
                                os.unlink(path)
                            except Exception:
                                pass

        # --- Download button (persists across reruns) ---
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
# TAB 6: OCR DEBUG
# ============================================================
with tab_ocr:
    st.subheader("🔍 OCR Extraction Debug")
    st.caption(
        "This tab shows what the OCR engine extracted from the EWP image. "
        "Use it to verify accuracy and diagnose extraction issues."
    )

    if not EWP_OCR_AVAILABLE:
        st.warning(
            "⚠️ OCR engine (Tesseract) is not installed. "
            "Add `tesseract-ocr` to `packages.txt` and redeploy."
        )
    elif not st.session_state.get("ewp_uploaded"):
        st.info("ℹ️ Upload the EWP image first.")
    else:
        col1, col2 = st.columns([1, 1])

        with col1:
            st.markdown("### 🖼️ EWP Image")
            try:
                st.image(st.session_state.ewp_image_bytes, use_container_width=True)
            except Exception as e:
                st.error(f"❌ Failed to render image: {e}")

        with col2:
            st.markdown("### 📝 Raw OCR Text")
            ocr_text = st.session_state.get("ewp_ocr_text", "")
            if ocr_text:
                st.text_area(
                    "Extracted text",
                    value=ocr_text,
                    height=500,
                    label_visibility="collapsed",
                )
            else:
                st.info("ℹ️ No OCR text available yet.")

        st.markdown("---")
        st.markdown("### 🎯 Auto-Extracted Values")

        try:
            ewp_parsed = parse_ewp(st.session_state.ewp_image_bytes)
            if ewp_parsed:
                rows = [
                    {
                        "Placeholder": f"{{{{{k}}}}}",
                        "Value": v,
                        "Currently Set": get_value(k) or "—",
                    }
                    for k, v in ewp_parsed.items()
                ]
                st.dataframe(rows, use_container_width=True, hide_index=True)
                st.success(f"✅ {len(ewp_parsed)} values extracted from OCR")
            else:
                st.warning(
                    "⚠️ No values were extracted. "
                    "Check the OCR text above — the image may be blurry or low resolution."
                )
        except Exception as e:
            st.error(f"❌ Extraction failed: {e}")
            st.exception(e)


# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption("MOP Automation • Nokia Lightspan MF-2 OLT Integration • Globe Telecom")
