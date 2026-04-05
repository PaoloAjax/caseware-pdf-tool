import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from openai import OpenAI

# =========================================================
# Streamlit setup
# =========================================================
st.set_page_config(page_title="CaseWare PDF -> Excel", layout="wide")
st.title("CaseWare PDF -> Excel")
st.write(
    "Upload CaseWare PDF-exports, herken vragen, genereer AI-antwoorden "
    "en exporteer alles naar Excel."
)

# =========================================================
# OpenAI helper
# =========================================================
def get_openai_client():
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return None
    return OpenAI(api_key=api_key)

# =========================================================
# Text helpers
# =========================================================
def normalize_line(line: str) -> str:
    if not line:
        return ""
    line = line.replace("\xa0", " ")
    line = line.replace("’", "'").replace("‘", "'")
    line = line.replace("“", '"').replace("”", '"')
    line = re.sub(r"\s+", " ", line).strip()
    return line


def clean_multiline_text(lines: list[str]) -> str:
    """
    Bouw nette vraagtekst op.
    Bullets blijven op aparte regels.
    Gewone tekstregels worden samengevoegd.
    """
    output = []
    paragraph = []

    def flush_paragraph():
        nonlocal paragraph, output
        if paragraph:
            text = " ".join(paragraph)
            text = re.sub(r"\s+([,.;:])", r"\1", text)
            text = re.sub(r"\(\s+", "(", text)
            text = re.sub(r"\s+\)", ")", text)
            output.append(text.strip())
            paragraph = []

    for line in lines:
        line = normalize_line(line)
        if not line:
            continue

        if is_bullet_line(line):
            flush_paragraph()
            output.append(normalize_bullet(line))
        else:
            paragraph.append(line)

    flush_paragraph()
    return "\n".join(output).strip()


def is_bullet_line(line: str) -> bool:
    return bool(re.match(r"^[-•*]\s+", line))


def normalize_bullet(line: str) -> str:
    line = re.sub(r"^[-•*]\s*", "- ", line)
    line = re.sub(r"\s+", " ", line).strip()
    return line


# =========================================================
# Noise detection
# =========================================================
NOISE_EXACT = {
    "...",
    "antwoord",
    "onderbouwing",
    "naam datum",
    "opgesteld",
}

NOISE_STARTS = [
    "cliëntnaam",
    "clientnaam",
    "cliënt-code",
    "client-code",
    "jaar",
    "dossier",
    "tabblad",
    "volgnr",
    "nr. instructie",
    "! er zijn verplichte velden",
    "naam / datum",
]

QUESTION_END_MARKERS = {
    "onderbouwing (verplicht)",
    "onderbouwing",
}

def is_noise_line(line: str) -> bool:
    low = normalize_line(line).lower()

    if not low:
        return True

    if low in NOISE_EXACT:
        return True

    if any(low.startswith(prefix) for prefix in NOISE_STARTS):
        return True

    # pagina-indicatie zoals 1/1
    if re.match(r"^\d+\s*/\s*\d+$", low):
        return True

    # losse administratieve regel
    if re.match(r"^[a-z]+\s+[a-z]+$", low) and low in {"naam datum", "opgesteld door"}:
        return True

    return False


def is_question_end_marker(line: str) -> bool:
    return normalize_line(line).lower() in QUESTION_END_MARKERS


# =========================================================
# Question start detection
# =========================================================
def detect_question_start(line: str):
    """
    Herkent regels zoals:
    1 KICK-OFF
    2 VOORRAAD
    10 LIQUIDE MIDDELEN
    3 Debiteuren
    """
    line = normalize_line(line)
    match = re.match(r"^(\d+)\s+(.+)$", line)
    if not match:
        return None

    nr = match.group(1).strip()
    title = match.group(2).strip()

    # Bescherm tegen rommel zoals "2025 4. 1. 1 1/1"
    if len(title) < 2:
        return None

    return nr, title


# =========================================================
# PDF extractie
# =========================================================
def extract_lines_from_pdf(file_obj) -> list[str]:
    """
    Leest alle pagina's uit pdfplumber en geeft genormaliseerde regels terug.
    """
    extracted_lines = []

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            raw_lines = text.split("\n")

            for raw_line in raw_lines:
                line = normalize_line(raw_line)
                if line:
                    extracted_lines.append(line)

    return extracted_lines


# =========================================================
# Parser
# =========================================================
def build_question_text(title: str, body_lines: list[str]) -> str:
    body = clean_multiline_text(body_lines)

    if title and body:
        return f"{title}\n\n{body}"
    if title:
        return title
    return body


def parse_caseware_questions(lines: list[str]) -> list[dict]:
    """
    Parseert CaseWare-achtige PDF regels naar vragen.
    """
    questions = []

    current_nr = None
    current_title = None
    current_body = []
    in_question = False

    for line in lines:
        if is_noise_line(line):
            continue

        if is_question_end_marker(line):
            if in_question:
                vraag = build_question_text(current_title, current_body)
                questions.append({
                    "nr": current_nr,
                    "titel": current_title,
                    "vraag": vraag
                })

                current_nr = None
                current_title = None
                current_body = []
                in_question = False
            continue

        question_start = detect_question_start(line)
        if question_start:
            # vorige vraag eerst opslaan
            if in_question:
                vraag = build_question_text(current_title, current_body)
                questions.append({
                    "nr": current_nr,
                    "titel": current_title,
                    "vraag": vraag
                })

            current_nr, current_title = question_start
            current_body = []
            in_question = True
            continue

        if in_question:
            current_body.append(line)

    # laatste open vraag nog opslaan
    if in_question:
        vraag = build_question_text(current_title, current_body)
        questions.append({
            "nr": current_nr,
            "titel": current_title,
            "vraag": vraag
        })

    # laatste opschoning
    cleaned_questions = []
    for q in questions:
        vraag = normalize_line(q["vraag"].replace("\n", " \n "))
        vraag = re.sub(r" ?\n ?", "\n", vraag).strip()

        if len(vraag) < 5:
            continue

        cleaned_questions.append({
            "nr": q["nr"],
            "titel": q["titel"],
            "vraag": q["vraag"].strip()
        })

    return cleaned_questions


# =========================================================
# AI prompt & generatie
# =========================================================
def build_caseware_prompt(vraag_text: str) -> str:
    return f"""
Je helpt bij het opstellen van een professioneel dossierantwoord voor een CaseWare werkprogramma.

Schrijf in het Nederlands.
Schrijf zakelijk, concreet en compact.
Gebruik geen markdown.
Gebruik geen titel, geen inleiding en geen afsluiting.
Schrijf direct de tekst die in 'Onderbouwing (verplicht)' geplakt kan worden.

Belangrijke regels:
- Verwerk expliciet de onderdelen uit de instructie.
- Als de instructie meerdere punten bevat, geef dan een nette, zakelijke puntsgewijze uitwerking.
- Verzin geen feitelijke details die niet bekend zijn.
- Als informatie ontbreekt, benoem dan professioneel dat dit nog moet worden afgestemd, onderbouwd of aangevuld.
- Schrijf alsof dit dossierdocumentatie is van een accountant.
- Houd de toon professioneel en bruikbaar voor direct copy-paste gebruik.

Instructie:
{vraag_text}
""".strip()


def generate_ai_answer(client: OpenAI, vraag_text: str, model_name: str) -> str:
    prompt = build_caseware_prompt(vraag_text)

    try:
        response = client.responses.create(
            model=model_name,
            input=prompt
        )
        return response.output_text.strip()
    except Exception as e:
        return f"AI fout: {e}"


# =========================================================
# Excel export
# =========================================================
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="overzicht")
    return buffer.getvalue()


# =========================================================
# UI
# =========================================================
with st.sidebar:
    st.subheader("Instellingen")
    generate_ai = st.checkbox("Genereer AI-antwoorden", value=False)
    model_name = st.text_input("OpenAI model", value="gpt-5.4")
    show_debug = st.checkbox("Toon debug-info", value=False)

uploaded_files = st.file_uploader(
    "Upload PDF-bestanden",
    type=["pdf"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Nog geen PDF's geüpload.")
    st.stop()

client = get_openai_client()
if generate_ai and client is None:
    st.error("OPENAI_API_KEY ontbreekt in st.secrets.")
    st.stop()

all_rows = []
debug_data = []

for f in uploaded_files:
    lines = extract_lines_from_pdf(f)
    questions = parse_caseware_questions(lines)

    if show_debug:
        debug_data.append({
            "bestand": f.name,
            "regels": lines,
            "vragen": questions
        })

    for q in questions:
        ai_answer = ""
        if generate_ai:
            ai_answer = generate_ai_answer(client, q["vraag"], model_name)

        all_rows.append({
            "bestand": Path(f.name).name,
            "nr": q["nr"],
            "titel": q["titel"],
            "vraag": q["vraag"],
            "ai_antwoord": ai_answer,
            "kopieer_antwoord": ai_answer
        })

df = pd.DataFrame(all_rows)

st.subheader("Preview")
st.dataframe(df, use_container_width=True)

excel_bytes = to_excel_bytes(df)

st.download_button(
    "Download Excel",
    data=excel_bytes,
    file_name="caseware_overzicht.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =========================================================
# Debug sectie
# =========================================================
if show_debug:
    st.subheader("Debug")

    for item in debug_data:
        with st.expander(f"Debug: {item['bestand']}"):
            st.write("Ruwe/genormaliseerde regels")
            st.write(item["regels"])

            st.write("Gevonden vragen")
            st.json(item["vragen"])
