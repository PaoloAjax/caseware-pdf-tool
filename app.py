import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import pdfplumber

st.set_page_config(page_title="CaseWare PDF tool", layout="wide")
st.title("CaseWare PDF -> Excel")
st.write("Upload je PDF-exports. Deze eerste versie zet alleen bestandsnamen in een Excel.")

uploaded_files = st.file_uploader(
    "Upload PDF-bestanden",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    rows = []
    for f in uploaded_files:
    tekst = ""

    with pdfplumber.open(f) as pdf:
        for page in pdf.pages:
            tekst += page.extract_text() or ""

    rows.append({
        "bestand": Path(f.name).name,
        "tekst_preview": tekst[:300]
    })

    df = pd.DataFrame(rows)
    st.subheader("Preview")
    st.dataframe(df, use_container_width=True)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="overzicht")

    st.download_button(
        "Download Excel",
        data=buffer.getvalue(),
        file_name="caseware_overzicht.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Nog geen PDF's geüpload.")
