# MOP Automation — Nokia Lightspan MF-2 OLT Integration

Streamlit app that generates a filled **MOP Integration Template** from:
- **FIO Excel** (`fio.xlsx`) — auto-mapped values
- **EWP Image** (`.jpg/.png`) — inserted into `{{ewp_image}}` placeholder
- **Manual EWP-only inputs** — for values not present in the FIO

---

## 🚀 Quick Start (Local)

```bash
git clone https://github.com/<your-org>/mop-automation.git
cd mop-automation
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

---

## 🌐 Deploy on Streamlit Cloud

1. Push this repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io).
3. Connect the repo, set **Main file path** to `app.py`.
4. Click **Deploy**.

---

## 📂 Project Structure

```
mop-automation/
├── app.py                          # Streamlit main app
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── .gitignore                      # Git ignores
├── .streamlit/
│   └── config.toml                 # Streamlit theme + upload limits
├── template/
│   └── MOP_INTEGRATION_TEMPLATE.docx  # Base template (with placeholders)
└── utils/
    ├── __init__.py
    ├── fio_parser.py               # FIO Excel parser
    ├── placeholder_map.py          # Master placeholder registry
    ├── docx_replacer.py            # Text replacement in DOCX
    └── image_replacer.py           # Image insertion for {{ewp_image}}
```

---

## 🧩 Placeholder Sources

| Source | Description | UI Tab |
|--------|-------------|--------|
| **FIO** | Auto-extracted from FIO Excel | 📊 FIO-Mapped (Auto) |
| **EWP** | Manual input from EWP image | 🖼️ EWP-Only (Manual) |
| **DOC** | Document metadata with defaults | 📝 Document Metadata |

---

## 🔄 Workflow

1. **Upload FIO Excel** → app parses and maps ~40 values automatically.
2. **Upload EWP Image** → app stores it for insertion.
3. **Review FIO-mapped values** → editable in Tab 1.
4. **Input EWP-only values** → Tab 2.
5. **Verify preview** → Tab 4.
6. **Generate MOP** → Tab 5, download `.docx`.

---

## 🛠️ Customization

### Add a new placeholder
1. Add an entry to `PLACEHOLDER_MAP` in `utils/placeholder_map.py`.
2. If it comes from the FIO, add parsing logic to `utils/fio_parser.py`.
3. The UI updates automatically.

### Update the template
Replace `template/MOP_INTEGRATION_TEMPLATE.docx` with a new version.
Ensure placeholders use the `{{PLACEHOLDER_NAME}}` format.

---

## 📝 License

Internal — Globe Telecom. Not for external distribution.
