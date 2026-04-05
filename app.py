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
                tekst += (page.extract_text() or "") + "\n"

        regels = tekst.split("\n")

regels = tekst.split("\n")

current_text = ""
blocks = []

for regel in regels:
    regel = regel.strip()

    if not regel:
        continue

    if any(x in regel.lower() for x in [
        "cliëntnaam", "tabblad", "volgnr", "naam datum", "opgesteld"
    ]):
        continue

    if regel.startswith("1 ") or regel.startswith("2 ") or regel.startswith("3 "):
        if current_text:
            blocks.append(current_text.strip())
        current_text = regel
    else:
        current_text += " " + regel

if current_text:
    blocks.append(current_text.strip())

for block in blocks:
    rows.append({
        "bestand": Path(f.name).name,
        "instructie": block
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
