import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="CaseWare PDF -> Excel", layout="wide")
st.title("CaseWare PDF -> Excel")
st.write("Upload CaseWare PDF-exports, herken vragen en exporteer naar Excel.")

# =========================================================
# Helpers
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
    title = re.sub(r"\s*\.\.\.\s*$", "", title)   # haal trailing ... weg
    title = re.sub(r"\s+", " ", title).strip()
    return title


def is_bullet_start(line: str) -> bool:
    return bool(re.match(r"^[-•*]\s+", line))


def normalize_bullet(line: str) -> str:
    line = re.sub(r"^[-•*]\s*", "", line)
    line = normalize_line(line)
    return f"- {line}"


# =========================================================
# Ruisfilter
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

    # pagina zoals 1/1 of 2/5
    if re.match(r"^\d+\s*/\s*\d+$", low):
        return True

    return False


def is_question_end_marker(line: str) -> bool:
    return normalize_line(line).lower() in QUESTION_END_MARKERS


# =========================================================
# Vraagstart detectie
# =========================================================
def detect_question_start(line: str):
    """
    Herkent regels zoals:
    1 KICK-OFF
    2 VOORRAAD
    10 LIQUIDE MIDDELEN
    """
    line = normalize_line(line)
    match = re.match(r"^(\d+)\s+(.+)$", line)
    if not match:
        return None

    nr = match.group(1).strip()
    title = clean_title(match.group(2).strip())

    if len(title) < 2:
        return None

    # vermijd jaartallen als false positive
    if len(nr) == 4 and nr.startswith(("19", "20")):
        return None

    return nr, title


# =========================================================
# PDF extractie
# =========================================================
def extract_lines_from_pdf(file_obj) -> list[str]:
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
# Body opbouwen
# =========================================================
def build_question_body(body_lines: list[str]) -> str:
    """
    Maak van losse PDF-regels nette alinea's + bullets.
    """
    result = []
    current_paragraph = []
    current_bullet = None

    def flush_paragraph():
        nonlocal current_paragraph, result
        if current_paragraph:
            text = " ".join(current_paragraph)
            text = re.sub(r"\s+([,.;:])", r"\1", text)
            text = re.sub(r"\(\s+", "(", text)
            text = re.sub(r"\s+\)", ")", text)
            result.append(text.strip())
            current_paragraph = []

    def flush_bullet():
        nonlocal current_bullet, result
        if current_bullet:
            text = current_bullet
            text = re.sub(r"\s+([,.;:])", r"\1", text)
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
    # Belangrijk: check of dit waarschijnlijk doorlopende bullet is
    if current_bullet is not None:
        # als regel niet eindigt op punt → waarschijnlijk doorlopend
        if not current_bullet.endswith("."):
            current_bullet += " " + line
        else:
            flush_bullet()
            current_paragraph.append(line)
    else:
        current_paragraph.append(line)

    flush_paragraph()
    flush_bullet()

    return "\n".join(result).strip()


def build_question_text(title: str, body_lines: list[str]) -> str:
    body = build_question_body(body_lines)

    if title and body:
        return f"{title}\n\n{body}"
    if title:
        return title
    return body


# =========================================================
# Parser
# =========================================================
def parse_caseware_questions(lines: list[str]) -> list[dict]:
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

        question_start = detect_question_start(line)
        if question_start:
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

        cleaned_questions.append({
            "nr": q["nr"],
            "titel": q["titel"],
            "vraag": vraag
        })

    return cleaned_questions


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
    show_debug = st.checkbox("Toon debug-info", value=False)

uploaded_files = st.file_uploader(
    "Upload PDF-bestanden",
    type=["pdf"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Nog geen PDF's geüpload.")
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
        all_rows.append({
            "bestand": Path(f.name).name,
            "nr": q["nr"],
            "titel": q["titel"],
            "vraag": q["vraag"]
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

if show_debug:
    st.subheader("Debug")
    for item in debug_data:
        with st.expander(f"Debug: {item['bestand']}"):
            st.write("Ruwe/genormaliseerde regels")
            st.write(item["regels"])

            st.write("Gevonden vragen")
            st.json(item["vragen"])
