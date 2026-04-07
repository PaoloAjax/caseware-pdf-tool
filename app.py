import hashlib
import json
import re
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document as WordDocument
from openai import OpenAI

# =========================================================
# Streamlit setup
# =========================================================
st.set_page_config(page_title="AI Accountancy Workspace", layout="wide")
st.title("AI Accountancy Workspace")
st.write("Gebruik modules voor dossierwerk, jaarrekeningchecks en later periodebalanschecks.")

DOSSIER_PROMPT_VERSION = "v2_short_no_bullets"
CHECKER_PROMPT_VERSION = "v1_checker_minimal"

# =========================================================
# Session state init
# =========================================================
DEFAULT_SESSION_KEYS = {
    # Dossier module
    "dossier_parsed_rows": [],
    "dossier_ai_rows": [],
    "dossier_last_file_signature": None,
    "dossier_debug_data": [],
    # Jaarrekening checker module
    "checker_last_signature": None,
    "checker_processed": False,
    "checker_jr_text": "",
    "checker_jr_lines": [],
    "checker_jr_name": "",
    "checker_model_text": "",
    "checker_model_lines": [],
    "checker_model_name": "",
    "checker_issues": [],
    "checker_process_info": {},
}

for key, value in DEFAULT_SESSION_KEYS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# OpenAI helpers
# =========================================================
def get_openai_client():
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return None
    return OpenAI(api_key=api_key)


def safe_json_extract(text: str) -> Dict:
    text = text.strip()
    text = re.sub(r"^```json", "", text, flags=re.IGNORECASE).strip()
    text = re.sub(r"^```", "", text).strip()
    text = re.sub(r"```$", "", text).strip()

    match = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if match:
        text = match.group(0)

    try:
        return json.loads(text)
    except Exception:
        return {
            "status": "attention",
            "issue_text": text[:500].strip() if text else "AI kon output niet goed structureren."
        }


# =========================================================
# Generic text helpers
# =========================================================
def normalize_line(line: str) -> str:
    if not line:
        return ""
    line = line.replace("\xa0", " ")
    line = line.replace("’", "'").replace("‘", "'")
    line = line.replace("“", '"').replace("”", '"')
    line = line.replace("–", "-").replace("—", "-")
    line = re.sub(r"\s+", " ", line).strip()
    return line


def normalize_text(text: str) -> str:
    lines = [normalize_line(x) for x in text.splitlines()]
    lines = [x for x in lines if x]
    return "\n".join(lines)


def clean_title(title: str) -> str:
    title = normalize_line(title)
    title = re.sub(r"\s*\.\.\.\s*$", "", title)
    return title.strip()


def build_files_signature(uploaded_files: List) -> str:
    hasher = hashlib.sha256()
    for f in uploaded_files:
        file_bytes = f.getvalue()
        hasher.update(f.name.encode("utf-8"))
        hasher.update(file_bytes)
    return hasher.hexdigest()


def build_two_file_signature(file_a, file_b) -> str:
    hasher = hashlib.sha256()

    for f in [file_a, file_b]:
        hasher.update(f.name.encode("utf-8"))
        hasher.update(f.getvalue())

    return hasher.hexdigest()


# =========================================================
# DOCX / PDF extraction
# =========================================================
def extract_text_from_pdf_bytes(file_bytes: bytes) -> Tuple[str, List[str]]:
    lines = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.split("\n"):
                line = normalize_line(raw_line)
                if line:
                    lines.append(line)

    full_text = "\n".join(lines)
    return full_text, lines


def extract_text_from_docx_bytes(file_bytes: bytes) -> Tuple[str, List[str]]:
    doc = WordDocument(BytesIO(file_bytes))
    lines = []

    for para in doc.paragraphs:
        line = normalize_line(para.text)
        if line:
            lines.append(line)

    for table in doc.tables:
        for row in table.rows:
            cell_texts = []
            for cell in row.cells:
                cell_text = normalize_line(cell.text)
                if cell_text:
                    cell_texts.append(cell_text)
            if cell_texts:
                lines.append(" | ".join(cell_texts))

    full_text = "\n".join(lines)
    return full_text, lines


@st.cache_data(show_spinner=False)
def extract_text_from_file(file_name: str, file_bytes: bytes) -> Dict:
    suffix = Path(file_name).suffix.lower()

    if suffix == ".pdf":
        full_text, lines = extract_text_from_pdf_bytes(file_bytes)
    elif suffix == ".docx":
        full_text, lines = extract_text_from_docx_bytes(file_bytes)
    else:
        full_text, lines = "", []

    return {
        "file_name": file_name,
        "full_text": full_text,
        "lines": lines,
        "line_count": len(lines),
        "char_count": len(full_text),
    }


# =========================================================
# Dossier module helpers
# =========================================================
def is_bullet_start(line: str) -> bool:
    return bool(re.match(r"^[-•*]\s+", line))


def normalize_bullet(line: str) -> str:
    line = re.sub(r"^[-•*]\s*", "", line)
    line = normalize_line(line)
    return f"- {line}"


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

    match = re.match(r"^(\d+)\s+(.+)$", line)
    if match:
        nr = match.group(1).strip()
        title = clean_title(match.group(2).strip())

        if len(title) >= 2 and not (len(nr) == 4 and nr.startswith(("19", "20"))):
            return nr, title

    if is_likely_uppercase_title(line):
        return "", clean_title(line)

    return None


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
    extracted = extract_text_from_file(file_name, file_bytes)
    lines = extracted["lines"]
    questions = parse_caseware_questions(lines)

    return {
        "bestand": file_name,
        "lines": lines,
        "questions": questions,
        "question_count": len(questions)
    }


def build_global_context(parsed_rows):
    parts = []

    for row in parsed_rows:
        parts.append(
            f"Bestand: {row['bestand']}\n"
            f"Titel: {row['titel']}\n"
            f"Vraag:\n{row['vraag']}\n"
        )

    return "\n---\n".join(parts).strip()


def build_dossier_prompt(vraag_text: str, dossier_context: str) -> str:
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
def cached_generate_dossier_ai_answer(
    vraag_text: str,
    dossier_context: str,
    model_name: str,
    prompt_version: str
) -> str:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return "AI fout: OPENAI_API_KEY ontbreekt in Streamlit secrets."

    client = OpenAI(api_key=api_key)
    prompt = build_dossier_prompt(vraag_text, dossier_context)

    try:
        response = client.responses.create(
            model=model_name,
            input=prompt
        )
        text = response.output_text.strip()

        if not text:
            return "AI fout: leeg antwoord ontvangen."

        text = re.sub(r"(?m)^\s*[-•*]\s*", "", text)
        text = re.sub(r"\n\s*[-•*]\s*", " ", text)
        text = re.sub(r"\s{2,}", " ", text).strip()
        return text
    except Exception as e:
        return f"AI fout: {e}"


# =========================================================
# Jaarrekening checker helpers
# =========================================================
CHECK_DEFINITIONS = [
    {
        "name": "Oprichting ontbreekt",
        "type": "rule",
        "keywords": ["oprichting", "opgericht", "akte"],
    },
    {
        "name": "Eerste of verlengd boekjaar ontbreekt",
        "type": "rule",
        "keywords": ["eerste boekjaar", "verlengd boekjaar"],
    },
    {
        "name": "Activiteiten ontbreken",
        "type": "rule",
        "keywords": ["activiteiten"],
    },
    {
        "name": "Fiscale positie ontbreekt",
        "type": "rule",
        "keywords": ["fiscale positie", "belastbaar bedrag", "vennootschapsbelasting"],
    },
    {
        "name": "Lonen en sociale lasten niet apart opgenomen",
        "type": "rule",
        "keywords": ["lonen", "sociale lasten"],
    },
    {
        "name": "Werknemerstoelichting ontbreekt terwijl loonkosten aanwezig zijn",
        "type": "rule_special",
    },
    {
        "name": "Continuïteit mist bij risicosituatie",
        "type": "ai",
        "jr_keywords": ["continuïteit", "eigen vermogen", "negatief", "verlies"],
        "model_keywords": ["continuïteit", "materiële onzekerheid", "negatief"],
        "instruction": (
            "Controleer of er in de jaarrekening aanleiding is voor een continuïteitstoelichting "
            "en of die toelichting ontbreekt of te beperkt is."
        ),
    },
    {
        "name": "Leningen zijn onvoldoende toegelicht",
        "type": "ai",
        "jr_keywords": ["lening", "leningen", "rente", "looptijd", "aflossing", "schuld"],
        "model_keywords": ["langlopende schulden", "rente", "looptijd", "aflossing"],
        "instruction": (
            "Beoordeel of leningen voldoende zijn toegelicht. Let op rente, looptijd, aflossing, "
            "voorwaarden en andere relevante bepalingen."
        ),
    },
    {
        "name": "Grondslagen zijn te generiek of missen relevante posten",
        "type": "ai",
        "jr_keywords": ["grondslagen", "waardering", "resultaatbepaling", "balans", "winst-en-verliesrekening"],
        "model_keywords": ["grondslagen", "waardering", "resultaatbepaling"],
        "instruction": (
            "Beoordeel of de grondslagen voldoende aansluiten op de posten in balans en "
            "winst-en-verliesrekening en niet te generiek zijn."
        ),
    },
    {
        "name": "Toelichtingen missen onderdelen ten opzichte van het modelrapport",
        "type": "ai",
        "jr_keywords": ["toelichting", "balans", "winst-en-verliesrekening", "werknemers", "lening", "activiteiten"],
        "model_keywords": ["toelichting", "werknemers", "fiscale positie", "continuïteit", "activiteiten"],
        "instruction": (
            "Vergelijk de jaarrekening op hoofdlijnen met het modelrapport en signaleer "
            "een concreet ontbrekend of zwak toegelicht onderdeel."
        ),
    },
]


def contains_any(text: str, keywords: List[str]) -> bool:
    lower = text.lower()
    return any(k.lower() in lower for k in keywords)


def extract_relevant_snippets(text: str, keywords: List[str], window: int = 3, max_chars: int = 4500) -> str:
    lines = [normalize_line(x) for x in text.splitlines()]
    lines = [x for x in lines if x]

    selected_indexes = set()

    for idx, line in enumerate(lines):
        line_lower = line.lower()
        if any(k.lower() in line_lower for k in keywords):
            start = max(0, idx - window)
            end = min(len(lines), idx + window + 1)
            for j in range(start, end):
                selected_indexes.add(j)

    if not selected_indexes:
        snippet = "\n".join(lines[:80])
        return snippet[:max_chars]

    selected_lines = [lines[i] for i in sorted(selected_indexes)]
    snippet = "\n".join(selected_lines)
    return snippet[:max_chars]


def build_checker_ai_prompt(
    check_name: str,
    instruction: str,
    jr_snippet: str,
    model_snippet: str
) -> str:
    return f"""
Je controleert een jaarrekening voor een accountant.

Beoordeel uitsluitend deze check:
{check_name}

Instructie:
{instruction}

Geef output ALLEEN als JSON in exact dit formaat:
{{
  "status": "ok" of "attention" of "missing",
  "issue_text": "korte Nederlandse omschrijving van het verbeterpunt, leeg als status ok is"
}}

Regels:
- Gebruik "ok" als er geen verbeterpunt is.
- Gebruik "attention" als er een aandachtspunt of beperkte toelichting is.
- Gebruik "missing" als het onderdeel ontbreekt.
- Wees kort en concreet.
- Geen markdown.
- Geen extra tekst buiten JSON.

Jaarrekening snippet:
{jr_snippet}

Modelrapport snippet:
{model_snippet}
""".strip()


@st.cache_data(show_spinner=False)
def cached_checker_ai_assessment(
    check_name: str,
    instruction: str,
    jr_snippet: str,
    model_snippet: str,
    model_name: str,
    prompt_version: str
) -> Dict:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return {
            "status": "attention",
            "issue_text": "OPENAI_API_KEY ontbreekt in Streamlit secrets."
        }

    client = OpenAI(api_key=api_key)
    prompt = build_checker_ai_prompt(check_name, instruction, jr_snippet, model_snippet)

    try:
        response = client.responses.create(
            model=model_name,
            input=prompt
        )
        raw_text = response.output_text.strip()
        parsed = safe_json_extract(raw_text)

        status = str(parsed.get("status", "attention")).strip().lower()
        issue_text = str(parsed.get("issue_text", "")).strip()

        if status not in {"ok", "attention", "missing"}:
            status = "attention"

        return {
            "status": status,
            "issue_text": issue_text
        }
    except Exception as e:
        return {
            "status": "attention",
            "issue_text": f"AI fout bij check '{check_name}': {e}"
        }


def run_rule_check_issues(jr_text: str) -> List[Dict]:
    issues = []
    lower = jr_text.lower()

    # 1-5 eenvoudige checks
    for check in CHECK_DEFINITIONS:
        if check["type"] != "rule":
            continue

        if check["name"] == "Lonen en sociale lasten niet apart opgenomen":
            has_lonen = "lonen" in lower
            has_sociale_lasten = "sociale lasten" in lower
            if has_lonen and not has_sociale_lasten:
                issues.append({
                    "check": check["name"],
                    "status": "❌",
                    "opmerking": "Lonen zijn gevonden, maar sociale lasten niet apart."
                })
            elif not has_lonen:
                issues.append({
                    "check": check["name"],
                    "status": "❌",
                    "opmerking": "Lonen zijn niet als aparte post aangetroffen."
                })
            continue

        if not contains_any(jr_text, check["keywords"]):
            issues.append({
                "check": check["name"],
                "status": "❌",
                "opmerking": f"Onderdeel lijkt niet aanwezig in de jaarrekening."
            })

    # Werknemers-check
    has_lonen = "lonen" in lower or "personeelskosten" in lower
    has_werknemers = "werknemers" in lower or "gemiddeld aantal werknemers" in lower or "personeelsomvang" in lower

    if has_lonen and not has_werknemers:
        issues.append({
            "check": "Werknemerstoelichting ontbreekt terwijl loonkosten aanwezig zijn",
            "status": "⚠️",
            "opmerking": "Er zijn loonkosten aangetroffen, maar geen duidelijke werknemers- of personeelsinformatie."
        })

    return issues


def run_ai_check_issues(jr_text: str, model_text: str, model_name: str) -> List[Dict]:
    issues = []

    for check in CHECK_DEFINITIONS:
        if check["type"] != "ai":
            continue

        jr_snippet = extract_relevant_snippets(jr_text, check["jr_keywords"])
        model_snippet = extract_relevant_snippets(model_text, check["model_keywords"])

        result = cached_checker_ai_assessment(
            check_name=check["name"],
            instruction=check["instruction"],
            jr_snippet=jr_snippet,
            model_snippet=model_snippet,
            model_name=model_name,
            prompt_version=CHECKER_PROMPT_VERSION
        )

        if result["status"] == "ok":
            continue

        status_symbol = "⚠️" if result["status"] == "attention" else "❌"
        issue_text = result["issue_text"] or "Verbeterpunt geconstateerd."

        issues.append({
            "check": check["name"],
            "status": status_symbol,
            "opmerking": issue_text
        })

    return issues


def deduplicate_issues(issues: List[Dict]) -> List[Dict]:
    seen = set()
    deduped = []

    for issue in issues:
        key = (
            issue.get("check", "").strip().lower(),
            issue.get("status", "").strip(),
            issue.get("opmerking", "").strip().lower(),
        )
        if key not in seen:
            seen.add(key)
            deduped.append(issue)

    return deduped


def build_checker_word_report(
    jr_name: str,
    model_name: str,
    issues: List[Dict]
) -> bytes:
    doc = WordDocument()
    doc.add_heading("Jaarrekening checker - verbeterpunten", level=1)
    doc.add_paragraph(f"Jaarrekening: {jr_name}")
    doc.add_paragraph(f"Modelrapport: {model_name}")

    doc.add_paragraph("")

    if not issues:
        doc.add_paragraph("Er zijn geen verbeterpunten gevonden.")
    else:
        for idx, issue in enumerate(issues, start=1):
            doc.add_heading(f"{idx}. {issue['check']}", level=2)
            doc.add_paragraph(f"Status: {issue['status']}")
            doc.add_paragraph(issue["opmerking"])

    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# =========================================================
# Generic export helpers
# =========================================================
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
    openai_model_name = st.text_input("OpenAI model", value="gpt-4.1-mini")
    show_debug = st.checkbox("Toon debug-info", value=False)

    if st.button("Reset hele sessie"):
        for key, value in DEFAULT_SESSION_KEYS.items():
            st.session_state[key] = value
        st.rerun()


# =========================================================
# Tabs
# =========================================================
tab_dossier, tab_checker, tab_period = st.tabs(
    ["Dossier", "Jaarrekening checker", "Periodebalans checker"]
)

# =========================================================
# TAB 1 - DOSSIER
# =========================================================
with tab_dossier:
    st.subheader("Dossier")
    st.write("Upload meerdere CaseWare-PDF's, verwerk vragen en genereer daarna AI-antwoorden.")

    dossier_uploaded_files = st.file_uploader(
        "Upload PDF-bestanden voor dossiermodule",
        type=["pdf"],
        accept_multiple_files=True,
        key="dossier_uploader"
    )

    if dossier_uploaded_files:
        current_signature = build_files_signature(dossier_uploaded_files)

        if st.session_state.dossier_last_file_signature != current_signature:
            st.session_state.dossier_parsed_rows = []
            st.session_state.dossier_ai_rows = []
            st.session_state.dossier_debug_data = []
            st.session_state.dossier_last_file_signature = current_signature

        col1, col2 = st.columns(2)

        with col1:
            process_clicked = st.button("Verwerk PDF's", use_container_width=True, key="dossier_process")

        with col2:
            ai_clicked = st.button("Genereer AI-antwoorden", use_container_width=True, key="dossier_ai")

        if process_clicked:
            parsed_rows = []
            debug_data = []
            summary_rows = []

            for f in dossier_uploaded_files:
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

            st.session_state.dossier_parsed_rows = parsed_rows
            st.session_state.dossier_ai_rows = parsed_rows.copy()
            st.session_state.dossier_debug_data = debug_data

            st.success("PDF's verwerkt. Controleer eerst de vragen en klik daarna op 'Genereer AI-antwoorden'.")
            st.subheader("Samenvatting per bestand")
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

        if st.session_state.dossier_parsed_rows:
            st.subheader("Preview vragen")
            preview_df = pd.DataFrame(st.session_state.dossier_ai_rows)
            st.dataframe(preview_df, use_container_width=True)
        else:
            st.info("Klik op 'Verwerk PDF's' om eerst de vragen uit de PDF's te halen.")

        if ai_clicked:
            if not st.session_state.dossier_parsed_rows:
                st.error("Verwerk eerst de PDF's voordat je AI-antwoorden genereert.")
                st.stop()

            client = get_openai_client()
            if client is None:
                st.error("OPENAI_API_KEY ontbreekt in Streamlit secrets.")
                st.stop()

            dossier_context = build_global_context(st.session_state.dossier_parsed_rows)

            updated_rows = []
            progress = st.progress(0)
            total = len(st.session_state.dossier_parsed_rows)

            for i, row in enumerate(st.session_state.dossier_parsed_rows, start=1):
                with st.spinner(f"AI antwoord genereren voor {row['titel']}..."):
                    ai_answer = cached_generate_dossier_ai_answer(
                        vraag_text=row["vraag"],
                        dossier_context=dossier_context,
                        model_name=openai_model_name,
                        prompt_version=DOSSIER_PROMPT_VERSION
                    )

                new_row = row.copy()
                new_row["ai_antwoord"] = ai_answer
                updated_rows.append(new_row)
                progress.progress(i / total)

            st.session_state.dossier_ai_rows = updated_rows
            st.success("AI-antwoorden zijn gegenereerd.")

        if st.session_state.dossier_ai_rows:
            st.subheader("Output")
            output_df = pd.DataFrame(st.session_state.dossier_ai_rows)
            st.dataframe(output_df, use_container_width=True)

            excel_bytes = dataframe_to_excel_bytes(st.session_state.dossier_ai_rows)

            st.download_button(
                "Download Excel",
                data=excel_bytes,
                file_name="caseware_overzicht.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dossier_download"
            )

        if show_debug and st.session_state.dossier_debug_data:
            st.subheader("Debug parsing")
            for parsed in st.session_state.dossier_debug_data:
                with st.expander(f"Debug: {parsed['bestand']}"):
                    st.write(f"Aantal gevonden vragen: {parsed['question_count']}")
                    st.write("Ruwe/genormaliseerde regels")
                    st.write(parsed["lines"])
                    st.write("Gevonden vragen")
                    st.json(parsed["questions"])
    else:
        st.info("Upload hier je CaseWare-PDF's voor de dossiermodule.")


# =========================================================
# TAB 2 - JAARREKENING CHECKER
# =========================================================
with tab_checker:
    st.subheader("Jaarrekening checker")
    st.write(
        "Upload een jaarrekening en een modelrapport. De checker toont alleen verbeterpunten "
        "en maakt daarna een Word-rapport."
    )

    col_a, col_b = st.columns(2)

    with col_a:
        jr_file = st.file_uploader(
            "Upload jaarrekening (PDF of Word)",
            type=["pdf", "docx"],
            accept_multiple_files=False,
            key="checker_jr_uploader"
        )

    with col_b:
        model_file = st.file_uploader(
            "Upload modelrapport (Word of PDF)",
            type=["docx", "pdf"],
            accept_multiple_files=False,
            key="checker_model_uploader"
        )

    if jr_file and model_file:
        checker_signature = build_two_file_signature(jr_file, model_file)

        if st.session_state.checker_last_signature != checker_signature:
            st.session_state.checker_last_signature = checker_signature
            st.session_state.checker_processed = False
            st.session_state.checker_jr_text = ""
            st.session_state.checker_jr_lines = []
            st.session_state.checker_jr_name = ""
            st.session_state.checker_model_text = ""
            st.session_state.checker_model_lines = []
            st.session_state.checker_model_name = ""
            st.session_state.checker_issues = []
            st.session_state.checker_process_info = {}

        col1, col2 = st.columns(2)

        with col1:
            checker_process_clicked = st.button(
                "Verwerk documenten",
                use_container_width=True,
                key="checker_process"
            )

        with col2:
            checker_run_clicked = st.button(
                "Voer checks uit",
                use_container_width=True,
                key="checker_run"
            )

        if checker_process_clicked:
            jr_extracted = extract_text_from_file(jr_file.name, jr_file.getvalue())
            model_extracted = extract_text_from_file(model_file.name, model_file.getvalue())

            st.session_state.checker_processed = True
            st.session_state.checker_jr_text = jr_extracted["full_text"]
            st.session_state.checker_jr_lines = jr_extracted["lines"]
            st.session_state.checker_jr_name = jr_file.name

            st.session_state.checker_model_text = model_extracted["full_text"]
            st.session_state.checker_model_lines = model_extracted["lines"]
            st.session_state.checker_model_name = model_file.name

            st.session_state.checker_process_info = {
                "jaarrekening_regels": jr_extracted["line_count"],
                "jaarrekening_tekens": jr_extracted["char_count"],
                "modelrapport_regels": model_extracted["line_count"],
                "modelrapport_tekens": model_extracted["char_count"],
            }

            st.success("Documenten verwerkt. Klik nu op 'Voer checks uit'.")

        if st.session_state.checker_processed:
            info = st.session_state.checker_process_info
            st.subheader("Documentinfo")
            st.write(
                f"Jaarrekening: {st.session_state.checker_jr_name} "
                f"({info.get('jaarrekening_regels', 0)} regels, {info.get('jaarrekening_tekens', 0)} tekens)"
            )
            st.write(
                f"Modelrapport: {st.session_state.checker_model_name} "
                f"({info.get('modelrapport_regels', 0)} regels, {info.get('modelrapport_tekens', 0)} tekens)"
            )

        if checker_run_clicked:
            if not st.session_state.checker_processed:
                st.error("Verwerk eerst de documenten voordat je de checks uitvoert.")
                st.stop()

            client = get_openai_client()
            if client is None:
                st.error("OPENAI_API_KEY ontbreekt in Streamlit secrets.")
                st.stop()

            jr_text = st.session_state.checker_jr_text
            model_text = st.session_state.checker_model_text

            with st.spinner("Rule-based checks uitvoeren..."):
                issues = run_rule_check_issues(jr_text)

            with st.spinner("AI-checks uitvoeren..."):
                ai_issues = run_ai_check_issues(jr_text, model_text, openai_model_name)

            all_issues = deduplicate_issues(issues + ai_issues)
            st.session_state.checker_issues = all_issues

            if all_issues:
                st.success(f"{len(all_issues)} verbeterpunten gevonden.")
            else:
                st.success("Geen verbeterpunten gevonden op basis van de huidige checks.")

        if st.session_state.checker_issues:
            st.subheader("Verbeterpunten")
            issues_df = pd.DataFrame(st.session_state.checker_issues)
            st.dataframe(issues_df, use_container_width=True)

            word_bytes = build_checker_word_report(
                jr_name=st.session_state.checker_jr_name,
                model_name=st.session_state.checker_model_name,
                issues=st.session_state.checker_issues
            )

            st.download_button(
                "Download Word-rapport",
                data=word_bytes,
                file_name="jaarrekening_checker_verbeterpunten.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="checker_word_download"
            )
        elif st.session_state.checker_processed:
            st.info("Nog geen verbeterpunten zichtbaar. Klik op 'Voer checks uit'.")

        if show_debug and st.session_state.checker_processed:
            with st.expander("Debug jaarrekening tekst"):
                st.write(st.session_state.checker_jr_lines)

            with st.expander("Debug modelrapport tekst"):
                st.write(st.session_state.checker_model_lines)

    else:
        st.info("Upload zowel een jaarrekening als een modelrapport om de checker te gebruiken.")


# =========================================================
# TAB 3 - PERIODEBALANS CHECKER (placeholder)
# =========================================================
with tab_period:
    st.subheader("Periodebalans checker")
    st.info("Deze module bouwen we hierna.")
