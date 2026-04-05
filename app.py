import re
from collections import Counter
from io import BytesIO
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="CaseWare PDF tool", layout="wide")
st.title("CaseWare PDF -> Excel")
st.write("Upload je PDF-exports. Deze versie filtert ruis, herkent instructies en kan AI-antwoorden voorbereiden.")

uploaded_files = st.file_uploader(
    "Upload PDF-bestanden",
    type=["pdf"],
    accept_multiple_files=True
)

# -----------------------------
# Config / woordenlijsten
# -----------------------------
NOISE_KEYWORDS = [
    "cliëntnaam",
    "clientnaam",
    "tabblad",
    "volgnr",
    "naam datum",
    "opgesteld",
    "aangemaakt",
    "gewijzigd",
    "dossier",
    "werkprogramma",
    "controleprogramma",
]

INSTRUCTION_VERBS = [
    "controleer",
    "beoordeel",
    "vul",
    "licht toe",
    "geef aan",
    "documenteer",
    "neem op",
    "stem af",
    "verwerk",
    "onderbouw",
    "beschrijf",
    "bevestig",
    "vergelijk",
    "specificeer",
    "verklaar",
    "voeg toe",
    "ga na",
    "sluit aan",
    "bereken",
]

# Start van instructie, zoals:
# 1 Controleer...
# 1. Controleer...
# 1) Controleer...
# 1.1 Controleer...
INSTRUCTION_START_RE = re.compile(r"^\d+(?:[.)]|\.\d+)*\s+")

PAGE_NUMBER_RE = re.compile(r"^(pagina\s+\d+(\s+van\s+\d+)?)$", re.IGNORECASE)
ONLY_NUMBERS_RE = re.compile(r"^\d+$")
DATE_RE = re.compile(r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b")
TIME_RE = re.compile(r"\b\d{1,2}:\d{2}\b")


# -----------------------------
# Helpers
# -----------------------------
def normalize_line(line: str) -> str:
    line = line.replace("\xa0", " ")
    line = re.sub(r"\s+", " ", line).strip()
    return line


def looks_like_noise(line: str, repeated_lines: set[str]) -> bool:
    low = line.lower().strip()

    if not low:
        return True

    if line in repeated_lines:
        return True

    if PAGE_NUMBER_RE.match(low):
        return True

    if ONLY_NUMBERS_RE.match(low):
        return True

    # Hele korte "rommel"-regels
    if len(low) <= 2:
        return True

    # Veel metadata-regels
    if any(k in low for k in NOISE_KEYWORDS):
        return True

    # Bestands-/systeemachtige regels
    if "\\" in line or "/" in line and len(line) > 60:
        return True

    # Alleen datum/tijd of bijna alleen administratie
    if DATE_RE.search(line) and len(line) < 25:
        return True

    if TIME_RE.search(line) and len(line) < 15:
        return True

    # Regels die vooral uit symbolen bestaan
    cleaned = re.sub(r"[A-Za-zÀ-ÿ0-9]", "", line)
    if len(cleaned) > 0 and len(cleaned) / max(len(line), 1) > 0.6:
        return True

    return False


def is_instruction_start(line: str) -> bool:
    low = line.lower()

    # Nummering aan het begin
    if INSTRUCTION_START_RE.match(line):
        return True

    # Soms begint een instructie direct met werkwoord
    if any(low.startswith(v + " ") for v in INSTRUCTION_VERBS):
        return True

    return False


def clean_instruction_text(text: str) -> str:
    text = normalize_line(text)

    # spaties voor leestekens opschonen
    text = re.sub(r"\s+([,.;:])", r"\1", text)

    # dubbele punten/spaties herstellen
    text = re.sub(r"\s{2,}", " ", text)

    return text.strip()


def classify_instruction(text: str) -> str:
    low = text.lower()

    if any(w in low for w in ["licht toe", "verklaar", "beschrijf"]):
        return "toelichting"
    if any(w in low for w in ["controleer", "beoordeel", "vergelijk", "ga na"]):
        return "controle"
    if any(w in low for w in ["vul", "verwerk", "neem op", "documenteer", "voeg toe"]):
        return "actie"

    return "overig"


def extract_pdf_lines(file_obj):
    """
    Leest alle pagina's en geeft een lijst terug:
    [(page_number, line), ...]
    """
    page_lines = []

    with pdfplumber.open(file_obj) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            lines = text.split("\n")

            for line in lines:
                norm = normalize_line(line)
                if norm:
                    page_lines.append((page_num, norm))

    return page_lines


def find_repeated_lines(page_lines):
    """
    Regels die op meerdere pagina's exact terugkomen zijn vaak header/footer.
    """
    line_pages = {}

    for page_num, line in page_lines:
        line_pages.setdefault(line, set()).add(page_num)

    repeated = {
        line for line, pages in line_pages.items()
        if len(pages) >= 2 and len(line) < 120
    }

    return repeated


def build_instruction_blocks(page_lines):
    repeated_lines = find_repeated_lines(page_lines)

    filtered_lines = []
    for page_num, line in page_lines:
        if not looks_like_noise(line, repeated_lines):
            filtered_lines.append(line)

    blocks = []
    current = ""

    for line in filtered_lines:
        if is_instruction_start(line):
            if current:
                blocks.append(clean_instruction_text(current))
            current = line
        else:
            if current:
                # doorlopende tekst toevoegen aan bestaand blok
                if current.endswith((".", ":", ";")):
                    current += " " + line
                else:
                    current += " " + line
            else:
                # tekst vóór eerste instructie negeren
                continue

    if current:
        blocks.append(clean_instruction_text(current))

    # Extra opschoning: weg met hele korte of verdachte blokken
    cleaned_blocks = []
    for b in blocks:
        if len(b) < 10:
            continue
        if len(b.split()) < 3:
            continue
        cleaned_blocks.append(b)

    return cleaned_blocks


def generate_ai_answer_placeholder(instruction_text: str) -> str:
    """
    Vervang dit later door echte AI-aanroep.
    """
    return (
        "Voorbeeld AI-antwoord: vat deze instructie samen en geef een zakelijk controleantwoord. "
        f"Instructie: {instruction_text[:250]}"
    )


# -----------------------------
# UI
# -----------------------------
generate_ai = st.checkbox("Genereer AI-antwoorden (placeholder)", value=False)

if uploaded_files:
    rows = []

    for f in uploaded_files:
        page_lines = extract_pdf_lines(f)
        blocks = build_instruction_blocks(page_lines)

        for i, block in enumerate(blocks, start=1):
            cleaned = clean_instruction_text(block)
            instruction_type = classify_instruction(cleaned)

            ai_answer = ""
            if generate_ai:
                ai_answer = generate_ai_answer_placeholder(cleaned)

            rows.append({
                "bestand": Path(f.name).name,
                "blok_nr": i,
                "type": instruction_type,
                "instructie": cleaned,
                "ai_antwoord": ai_answer
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
