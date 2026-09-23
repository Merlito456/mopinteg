"""
)

st.markdown("---")

# ---- Tech Stack ----
st.markdown("## 🛠️ Built With")

tech_col1, tech_col2, tech_col3 = st.columns(3)

with tech_col1:
st.markdown(
    """
    **Frontend**
    - Streamlit
    - Custom CSS
    """
)

with tech_col2:
st.markdown(
    """
    **Document Processing**
    - `python-docx`
    - `openpyxl`
    - Raw OOXML
    """
)

with tech_col3:
st.markdown(
    """
    **AI & OCR**
    - Tesseract OCR
    - Gemini (via prompt)
    - `Pillow`
    """
)

st.markdown("---")

# ---- Disclaimer ----
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

# ---- Contact ----
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
