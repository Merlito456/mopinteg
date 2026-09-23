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

# ---------------- CONFIG ----------------
st.set_page_config(
    page_title="MOP Automation",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

TEMPLATE_PATH = "template/MOP_INTEGRATION_TEMPLATE.docx"

# ---------------- SESSION STATE ----------------
if "placeholder_values" not in st.session_state:
    st.session_state.placeholder_values = {}

if "fio_uploaded" not in st.session_state:
    st.session_state.fio_uploaded = False

if "ewp_uploaded" not in st.session_state:
    st.session_state.ewp_uploaded = False

if "ewp_image_bytes" not in st.session_state:
    st.session_state.ewp_image_bytes = None


# ---------------- HELPERS ----------------
def set_value(key, value):
    st.session_state.placeholder_values[key] = value


def get_value(key, default=""):
    return st.session_state.placeholder_values.get(key, default)


# ---------------- HEADER ----------------
st.title("📄 MOP Automation — Nokia Lightspan MF-2 OLT Integration")
st.markdown(
    """
    Upload the **FIO Excel** and **EWP Image**, verify the mapped values, 
    then generate the filled **MOP Integration Template**.
    """
)

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.header("⚙️ Configuration")
    st.markdown("### Step 1 — Upload Files")
    fio_file = st.file_uploader("Upload FIO (.xlsx)", type=["xlsx"])
    ewp_image = st.file_uploader("Upload EWP Image (.jpg/.png)", type=["jpg", "jpeg", "png"])
    st.markdown("---")
    st.markdown("### Step 2 — Review & Edit")
    st.markdown(
        "FIO-derived values are **auto-filled** and editable.\n\n"
        "EWP-only values require **manual input**."
    )
    st.markdown("---")
    st.markdown("### Step 3 — Generate MOP")
    generate_btn = st.button("🚀 Generate MOP", use_container_width=True, type="primary")

# ---------------- PARSE FIO ----------------
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

# ---------------- STORE EWP IMAGE ----------------
if ewp_image and not st.session_state.ewp_uploaded:
    st.session_state.ewp_image_bytes = ewp_image.read()
    st.session_state.ewp_uploaded = True
    st.success("✅ EWP image loaded")

# ---------------- TABS ----------------
tab_fio, tab_ewp, tab_doc, tab_preview, tab_generate = st.tabs([
    "📊 FIO-Mapped (Auto)", "🖼️ EWP-Only (Manual)", "📝 Document Metadata", "👁️ Preview", "📄 Generate"
])

# ================= TAB 1: FIO-MAPPED =================
with tab_fio:
    st.subheader("FIO-Mapped Placeholders")
    st.caption("These values were extracted from the FIO. You may edit them below.")

    if not st.session_state.fio_uploaded:
        st.warning("⚠️ Upload the FIO Excel file first.")
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

# ================= TAB 2: EWP-ONLY =================
with tab_ewp:
    st.subheader("EWP-Only Placeholders (Manual Input)")
    st.caption(
        "These values cannot be extracted from the FIO — they must be read from the EWP image "
        "and entered manually."
    )

    if not st.session_state.ewp_uploaded:
        st.warning("⚠️ Upload the EWP image first (it will be inserted into the MOP).")

    groups = sorted(set(p["group"] for p in PLACEHOLDER_MAP if p["source"] == "EWP"))
    for group in groups:
        items = [p for p in PLACEHOLDER_MAP if p["source"] == "EWP" and p["group"] == group]
        with st.expander(f"📁 {group}", expanded=True):
            cols = st.columns(2)
            for i, item in enumerate(items):
                key = item["key"]
                label = item["label"]
                with cols[i % 2]:
                    val = st.text_input(
                        label,
                        value=get_value(key),
                        key=f"ewp_{key}",
                    )
                    if val != get_value(key):
                        set_value(key, val)

# ================= TAB 3: DOCUMENT METADATA =================
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

# ================= TAB 4: PREVIEW =================
with tab_preview:
    st.subheader("Preview — All Placeholder Values")
    st.caption("This is what will be substituted into the template.")

    if st.session_state.ewp_image_bytes:
        st.markdown("### 🖼️ EWP Image")
        st.image(st.session_state.ewp_image_bytes, use_column_width=True)

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

# ================= TAB 5: GENERATE =================
with tab_generate:
    st.subheader("Generate Filled MOP")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"❌ Template not found at `{TEMPLATE_PATH}`. Please add it to the repo.")
    elif not st.session_state.ewp_image_bytes:
        st.warning("⚠️ Please upload the EWP image first.")
    else:
        st.success("✅ Template found. Ready to generate.")

        if generate_btn:
            with st.spinner("Generating MOP..."):
                # Step 1: Replace text placeholders
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp1:
                    step1_path = tmp1.name

                mapping = {}
                for p in PLACEHOLDER_MAP:
                    key = p["key"]
                    val = get_value(key, p.get("default", ""))
                    if val:
                        mapping[key] = val

                replace_placeholders(TEMPLATE_PATH, mapping, step1_path)

                # Step 2: Replace EWP image
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp2:
                    step2_path = tmp2.name

                with open(step1_path, "rb") as f:
                    image_stream = BytesIO(st.session_state.ewp_image_bytes)

                replace_ewp_image(step1_path, image_stream, step2_path)

                # Step 3: Provide download
                with open(step2_path, "rb") as f:
                    output_bytes = f.read()

                output_filename = f"MOP_INTEGRATION_{get_value('OLT_SITE', 'OUTPUT')}.docx"

                st.success("✅ MOP generated successfully!")
                st.download_button(
                    label="⬇️ Download Filled MOP (.docx)",
                    data=output_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

                # Cleanup
                os.unlink(step1_path)
                os.unlink(step2_path)

# ---------------- FOOTER ----------------
st.markdown("---")
st.caption("MOP Automation • Nokia Lightspan MF-2 OLT Integration • Globe Telecom")
