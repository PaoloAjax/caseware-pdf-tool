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
    "Upload CaseWare PDF-exports, herken vragen, genereer optioneel AI-antwoorden "
    "en exporteer alles naar Excel."
)

# =========================================================
# OpenAI helpers
# =========================================================
def get_openai_client():
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return None
    return OpenAI(api_key=api_key)


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

Instructie:
{vraag_text}
""".strip()


def generate_ai_answer(client, vraag_text: str, model_name: str) -> str:
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


def clean_title(title: str) -> str:
    title = normalize_line(title)
    title = re.sub(r"\s*\.\.\.\s*$", "", title)
    return title.strip()


def is_bullet_start(line: str) -> bool:
    return bool(re.match(r"^[-•*]\s+", line))


def normalize_bullet(line: str) -> str:
    line = re.sub(r"^[-•*]\s*", "", line)
    line = normalize_line(line)
    return f"- {line}"


# =========================================================
# Noise filter
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

    if re.match(r"^\d+\s*/\s*\d+$", low):
        return True

    if "b.v." in low and re.search(r"\b20\d{2}\b", low):
        return True

    return False


def is_question_end_marker(line: str) -> bool:
    return normalize_line(line).lower() in QUESTION_END_MARKERS


# =========================================================
# Question start detection
# =========================================================
def detect_question_start(line: str):
    line = normalize_line(line)

    match = re.match(r"^(\d+)\s+(.+)$", line)
    if not match:
        return None

    nr = match.group(1).strip()
    title = clean_title(match.group(2).strip())

    if len(title) < 2:
        return None

    if len(nr) == 4 and nr.startswith(("19", "20")):
        return None

    return nr, title


# =========================================================
# PDF extraction
# =========================================================
def extract_lines_from_pdf(file_obj):
    lines = []

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.split("\n"):
                line = normalize_line(raw_line)
                if line:
                    lines.append(line)

    return lines


# =========================================================
# Body builder
# =========================================================
def build_question_body(body_lines):
    result = []
    current_paragraph = []
    current_bullet = None

    def flush_paragraph():
        nonlocal current_paragraph
        if current_paragraph:
            text = " ".join(current_paragraph)
            text = re.sub(r"\s+([,.;:])", r"\1", text)
            text = re.sub(r"\(\s+", "(", text)
            text = re.sub(r"\s+\)", ")", text)
            result.append(text.strip())
            current_paragraph = []

    def flush_bullet():
        nonlocal current_bullet
        if current_bullet:
            text = re.sub(r"\s+([,.;:])", r"\1", current_bullet)
            text = re.sub(r"\(\s+", "(", text)
            text = re.sub(r"\s+\)", ")", text)
            result.append(text.strip())
            current_bullet = None

    for raw_line in body_lines:
        line = normalize_line(raw_line)

        if not line:
            continue

        if is_bullet_start(line):
            flush_paragraph()
            flush_bullet()
            current_bullet = normalize_bullet(line)
        else:
            if current_bullet is not None:
                current_bullet += " " + line
            else:
                current_paragraph.append(line)

    flush_paragraph()
    flush_bullet()

    return "\n".join(result).strip()


def build_question_text(title, body_lines):
    body = build_question_body(body_lines)

    if title and body:
        return f"{title}\n\n{body}"
    if title:
        return title
    return body


# =========================================================
# Parser
# =========================================================
def parse_caseware_questions(lines):
    questions = []

    current_nr = None
    current_title = None
    current_body = []
    in_question = False

    for line in lines:
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

        if is_noise_line(line):
            continue

        start = detect_question_start(line)
        if start:
            if in_question:
                vraag = build_question_text(current_title, current_body)
                questions.append({
                    "nr": current_nr,
                    "titel": current_title,
                    "vraag": vraag
                })

            current_nr, current_title = start
            current_body = []
            in_question = True
            continue

        if in_question:
            current_body.append(line)

    if in_question:
        vraag = build_question_text(current_title, current_body)
        questions.append({
            "nr": current_nr,
            "titel": current_title,
            "vraag": vraag
        })

    cleaned_questions = []
    for q in questions:
        vraag = q["vraag"].strip()
        if len(vraag) < 5:
            continue
        cleaned_questions.append(q)

    return cleaned_questions


# =========================================================
# Excel export
# =========================================================
def to_excel_bytes(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="overzicht")
    return buffer.getvalue()


# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.subheader("Instellingen")
    show_debug = st.checkbox("Toon debug-info", value=False)
    generate_ai = st.checkbox("Genereer AI-antwoorden", value=False)
    model_name = st.text_input("Model", value="gpt-5.4")


# =========================================================
# File upload
# =========================================================
uploaded_files = st.file_uploader(
    "Upload PDF-bestanden",
    type=["pdf"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Upload eerst een PDF.")
    st.stop()


# =========================================================
# OpenAI client
# =========================================================
client = None
if generate_ai:
    client = get_openai_client()
    if client is None:
        st.error("OPENAI_API_KEY ontbreekt in Streamlit secrets.")
        st.stop()


# =========================================================
# Main processing
# =========================================================
rows = []
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

        rows.append({
            "bestand": Path(f.name).name,
            "nr": q["nr"],
            "titel": q["titel"],
            "vraag": q["vraag"],
            "ai_antwoord": ai_answer
        })

df = pd.DataFrame(rows)

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
# Debug
# =========================================================
if show_debug:
    st.subheader("Debug")

    for item in debug_data:
        with st.expander(f"Debug: {item['bestand']}"):
            st.write("Ruwe/genormaliseerde regels")
            st.write(item["regels"])

            st.write("Gevonden vragen")
            st.json(item["vragen"])
