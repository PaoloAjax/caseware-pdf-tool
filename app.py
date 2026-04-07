import hashlib
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
    "Upload CaseWare PDF-exports, verwerk eerst de vragen en genereer daarna "
    "optioneel AI-antwoorden."
)

PROMPT_VERSION = "v2_short_no_bullets"

# =========================================================
# Session state init
# =========================================================
if "parsed_rows" not in st.session_state:
    st.session_state.parsed_rows = []

if "ai_rows" not in st.session_state:
    st.session_state.ai_rows = []

if "last_file_signature" not in st.session_state:
    st.session_state.last_file_signature = None

if "debug_data" not in st.session_state:
    st.session_state.debug_data = []


# =========================================================
# OpenAI helpers
# =========================================================
def get_openai_client():
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return None
    return OpenAI(api_key=api_key)


def build_caseware_prompt(vraag_text: str, dossier_context: str) -> str:
    return f"""
Je helpt bij het opstellen van een professioneel dossierantwoord voor een CaseWare werkprogramma.

Schrijf in het Nederlands.
Schrijf zakelijk, concreet en compact.
Gebruik GEEN markdown.
Gebruik GEEN bullets.
Gebruik GEEN verbindingsstreepjes (-).
Gebruik GEEN opsommingstekens.
Schrijf alleen in korte lopende tekst of een korte alinea.

Schrijf zo kort mogelijk, maar wel volledig genoeg om direct in 'Onderbouwing (verplicht)' te plakken.
Vermijd herhaling, uitleg en inleidende tekst.
Gebruik bij voorkeur maximaal 3 tot 4 zinnen.

Belangrijke regels:
- Verwerk de kern van de instructie.
- Gebruik de meegeleverde dossiercontext als basis.
- Verzin geen feiten.
- Als informatie ontbreekt, benoem dit kort en zakelijk.
- Schrijf alsof dit direct in een accountantsdossier wordt opgenomen.

Vraag / instructie:
{vraag_text}

Dossiercontext:
{dossier_context}
""".strip()


@st.cache_data(show_spinner=False)
def cached_generate_ai_answer(
    vraag_text: str,
    dossier_context: str,
    model_name: str,
    prompt_version: str
) -> str:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return "AI fout: OPENAI_API_KEY ontbreekt in Streamlit secrets."

    client = OpenAI(api_key=api_key)
    prompt = build_caseware_prompt(vraag_text, dossier_context)

    try:
        response = client.responses.create(
            model=model_name,
            input=prompt
        )
        text = response.output_text.strip()

        if not text:
            return "AI fout: leeg antwoord ontvangen."

        # extra schoonmaak om bullets/verbindingsstreepjes alsnog weg te halen
        text = re.sub(r"(?m)^\s*[-•*]\s*", "", text)
        text = re.sub(r"\n\s*[-•*]\s*", " ", text)
        text = re.sub(r"\s{2,}", " ", text).strip()

        return text
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
def is_likely_uppercase_title(line: str) -> bool:
    line = normalize_line(line)

    if not line:
        return False

    if is_noise_line(line):
        return False

    if is_question_end_marker(line):
        return False

    if is_bullet_start(line):
        return False

    if len(line) > 120:
        return False

    letters = [c for c in line if c.isalpha()]
    if not letters:
        return False

    uppercase_letters = [c for c in letters if c.isupper()]
    ratio = len(uppercase_letters) / len(letters)

    return ratio >= 0.7


def detect_question_start(line: str):
    line = normalize_line(line)

    # Genummerde titel
    match = re.match(r"^(\d+)\s+(.+)$", line)
    if match:
        nr = match.group(1).strip()
        title = clean_title(match.group(2).strip())

        if len(title) >= 2 and not (len(nr) == 4 and nr.startswith(("19", "20"))):
            return nr, title

    # Ongenummerde titel in hoofdletters
    if is_likely_uppercase_title(line):
        return "", clean_title(line)

    return None


# =========================================================
# PDF extraction and parsing
# =========================================================
def extract_lines_from_pdf_bytes(file_bytes: bytes):
    lines = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.split("\n"):
                line = normalize_line(raw_line)
                if line:
                    lines.append(line)

    return lines


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


@st.cache_data(show_spinner=False)
def parse_pdf_to_questions(file_name: str, file_bytes: bytes):
    lines = extract_lines_from_pdf_bytes(file_bytes)
    questions = parse_caseware_questions(lines)

    return {
        "bestand": file_name,
        "lines": lines,
        "questions": questions,
        "question_count": len(questions)
    }


# =========================================================
# Context helpers
# =========================================================
def build_file_signature(uploaded_files):
    hasher = hashlib.sha256()

    for f in uploaded_files:
        file_bytes = f.getvalue()
        hasher.update(f.name.encode("utf-8"))
        hasher.update(file_bytes)

    return hasher.hexdigest()


def build_global_context(parsed_rows):
    parts = []

    for row in parsed_rows:
        parts.append(
            f"Bestand: {row['bestand']}\n"
            f"Titel: {row['titel']}\n"
            f"Vraag:\n{row['vraag']}\n"
        )

    return "\n---\n".join(parts).strip()


@st.cache_data(show_spinner=False)
def dataframe_to_excel_bytes(records):
    df = pd.DataFrame(records)
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
    model_name = st.text_input("Model", value="gpt-4.1-mini")

    if st.button("Reset sessie"):
        st.session_state.parsed_rows = []
        st.session_state.ai_rows = []
        st.session_state.last_file_signature = None
        st.session_state.debug_data = []
        st.rerun()


# =========================================================
# Upload
# =========================================================
uploaded_files = st.file_uploader(
    "Upload PDF-bestanden",
    type=["pdf"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Upload eerst je PDF-bestanden.")
    st.stop()

current_signature = build_file_signature(uploaded_files)

if st.session_state.last_file_signature != current_signature:
    st.session_state.parsed_rows = []
    st.session_state.ai_rows = []
    st.session_state.debug_data = []
    st.session_state.last_file_signature = current_signature

col1, col2 = st.columns(2)

with col1:
    process_clicked = st.button("Verwerk PDF's", use_container_width=True)

with col2:
    ai_clicked = st.button("Genereer AI-antwoorden", use_container_width=True)


# =========================================================
# Stap 1: Verwerk PDF's
# =========================================================
if process_clicked:
    parsed_rows = []
    debug_data = []
    summary_rows = []

    for f in uploaded_files:
        parsed = parse_pdf_to_questions(f.name, f.getvalue())

        debug_data.append(parsed)
        summary_rows.append({
            "bestand": f.name,
            "gevonden_vragen": parsed["question_count"]
        })

        for q in parsed["questions"]:
            parsed_rows.append({
                "bestand": Path(f.name).name,
                "nr": q["nr"],
                "titel": q["titel"],
                "vraag": q["vraag"],
                "ai_antwoord": ""
            })

    st.session_state.parsed_rows = parsed_rows
    st.session_state.ai_rows = parsed_rows.copy()
    st.session_state.debug_data = debug_data

    st.success("PDF's verwerkt. Controleer eerst de vragen en klik daarna op 'Genereer AI-antwoorden'.")

    st.subheader("Samenvatting per bestand")
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)


# =========================================================
# Preview van parse-resultaat
# =========================================================
if st.session_state.parsed_rows:
    st.subheader("Preview vragen")
    preview_df = pd.DataFrame(st.session_state.ai_rows)
    st.dataframe(preview_df, use_container_width=True)
else:
    st.info("Klik op 'Verwerk PDF's' om eerst de vragen uit de PDF's te halen.")


# =========================================================
# Stap 2: Genereer AI-antwoorden
# =========================================================
if ai_clicked:
    if not st.session_state.parsed_rows:
        st.error("Verwerk eerst de PDF's voordat je AI-antwoorden genereert.")
        st.stop()

    client = get_openai_client()
    if client is None:
        st.error("OPENAI_API_KEY ontbreekt in Streamlit secrets.")
        st.stop()

    dossier_context = build_global_context(st.session_state.parsed_rows)

    updated_rows = []
    progress = st.progress(0)
    total = len(st.session_state.parsed_rows)

    for i, row in enumerate(st.session_state.parsed_rows, start=1):
        with st.spinner(f"AI antwoord genereren voor {row['titel']}..."):
            ai_answer = cached_generate_ai_answer(
                vraag_text=row["vraag"],
                dossier_context=dossier_context,
                model_name=model_name,
                prompt_version=PROMPT_VERSION
            )

        new_row = row.copy()
        new_row["ai_antwoord"] = ai_answer
        updated_rows.append(new_row)

        progress.progress(i / total)

    st.session_state.ai_rows = updated_rows
    st.success("AI-antwoorden zijn gegenereerd.")


# =========================================================
# Definitieve preview + download
# =========================================================
if st.session_state.ai_rows:
    st.subheader("Output")
    output_df = pd.DataFrame(st.session_state.ai_rows)
    st.dataframe(output_df, use_container_width=True)

    excel_bytes = dataframe_to_excel_bytes(st.session_state.ai_rows)

    st.download_button(
        "Download Excel",
        data=excel_bytes,
        file_name="caseware_overzicht.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================================================
# Debug
# =========================================================
if show_debug and st.session_state.debug_data:
    st.subheader("Debug parsing")

    for parsed in st.session_state.debug_data:
        with st.expander(f"Debug: {parsed['bestand']}"):
            st.write(f"Aantal gevonden vragen: {parsed['question_count']}")
            st.write("Ruwe/genormaliseerde regels")
            st.write(parsed["lines"])

            st.write("Gevonden vragen")
            st.json(parsed["questions"])
