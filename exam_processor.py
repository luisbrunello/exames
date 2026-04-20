import copy
import datetime as dt
import io
import json
import re
from collections import Counter
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
import openpyxl
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process
from unidecode import unidecode

SHEET_NAME = "EXAMES"


def generic_result_pattern(title_pattern: str, unit_pattern: str) -> str:
    return rf"(?:{title_pattern}).*?RESULTADO[^\n]*\n\s*([\d.,]+)\s*{unit_pattern}"


EXAM_PATTERNS = {
    "GLICEMIA DE JEJUM": [
        generic_result_pattern(r"GLICOSE,\s*SORO|GLICEMIA(?:\s+DE\s+JEJUM)?", r"MG/DL"),
    ],
    "INSULINA DE JEJUM": [
        generic_result_pattern(r"INSULINA,\s*SORO|INSULINA(?:\s+DE\s+JEJUM)?", r"MU/L"),
    ],
    "HOMA IR": [
        r"HOMA[-\s]?IR\s*:\s*([\d.,]+)",
    ],
    "HEMOGLOBINA GLICADA A1C": [
        generic_result_pattern(r"HEMOGLOBINA GLICADA(?:\s*\(A1C\))?,\s*SANGUE TOTAL|HEMOGLOBINA GLICADA|A1C", r"%"),
    ],
    "GLICEMIA MÉDIA ESTIMADA": [
        r"GLICEMIA MEDIA ESTIMADA\s+DE\s+([\d.,]+)\s*MG/DL",
    ],
    "COLESTEROL TOTAL": [
        generic_result_pattern(r"COLESTEROL TOTAL,\s*SORO", r"MG/DL"),
    ],
    "COLESTEROL HDL": [
        generic_result_pattern(r"HDL[-\s]?COLESTEROL,\s*SORO", r"MG/DL"),
    ],
    "COLESTEROL LDL": [
        generic_result_pattern(r"LDL[-\s]?COLESTEROL,\s*SORO", r"MG/DL"),
    ],
    "TRIGLICERÍDEOS": [
        generic_result_pattern(r"TRIGLICERIDES,\s*SORO|TRIGLICERIDEOS,\s*SORO", r"MG/DL"),
    ],
    "COLESTEROL VLDL": [
        generic_result_pattern(r"VLDL[-\s]?COLESTEROL,\s*SORO", r"MG/DL"),
    ],
    "COLESTEROL NÃO HDL": [
        generic_result_pattern(r"NAO[-\s]?HDL[-\s]?COLESTEROL,\s*SORO", r"MG/DL"),
    ],
    "LIPOPROTEÍNA A": [
        generic_result_pattern(r"LIPOPROTEINA\s+L?P?\(A\),\s*SORO", r"NMOL/L"),
    ],
    "APOLIPOPROTEINA A1": [
        generic_result_pattern(r"APOLIPOPROTEINA\s+A[-\s]?1,\s*SORO", r"MG/DL"),
    ],
    "APOLIPOPROTEÍNA B": [
        generic_result_pattern(r"APOLIPOPROTEINA\s+B,\s*SORO", r"MG/DL"),
    ],
    "RELAÇÃO APO B / APO A": [
        r"RAZAO\s+APO\s+B\s*/\s*APO\s+A1\s*:\s*([\d.,]+)",
        r"RAZAO\s+APO\s+B\s*/\s*APO\s+A\s*:\s*([\d.,]+)",
    ],
    "PROTEÍNA C REATIVA ULTRASSENSÍVEL (PCR-US)": [
        generic_result_pattern(r"PROTEINA C[-\s]?REATIVA,\s*SORO", r"MG/DL"),
    ],
    "VELOCIDADE DE HEMOSSEDIMENTAÇÃO": [
        r"HEMOSSEDIMENTACAO.*?PRIMEIRA HORA\s*:\s*([\d.,]+)\s*MM",
        r"VEL(?:OCIDADE)?\s+DE\s+HEMOSSEDIMENTACAO.*?([\d.,]+)\s*MM",
    ],
    "UREIA": [
        generic_result_pattern(r"UREIA,\s*SORO", r"MG/DL"),
    ],
    "CREATININA": [
        generic_result_pattern(r"CREATININA,\s*SORO", r"MG/DL"),
    ],
    "TFG": [
        r"CKD[-\s]?EPI 2021\s*:\s*([\d.,]+)",
        r"CKDEPI 2021\s*:\s*([\d.,]+)",
    ],
    "CÁLCIO": [
        generic_result_pattern(r"CALCIO,\s*SORO", r"MG/DL"),
    ],
    "PARATORMÔNIO MOLÉCULA INTACTA": [
        generic_result_pattern(r"PARATORMONIO.*?SORO|\bPTH\b.*?SORO", r"PG/ML"),
    ],
    "FERRITINA": [
        generic_result_pattern(r"FERRITINA,\s*SORO", r"MICROG/L"),
    ],
    "FERRO": [
        generic_result_pattern(r"FERRO,\s*SORO", r"MCG/DL"),
    ],
    "SATURAÇÃO DA TRANSFERRINA": [
        r"SATURACAO DA TRANSFERRINA.*?SATURACAO DA TRANSFERRINA\s*:\s*([\d.,]+)\s*%",
    ],
    "HEMOGLOBINA": [
        r"HEMOGLOBINA\s*:\s*([\d.,]+)\s*G/DL",
    ],
    "HEMATÓCRITO": [
        r"HEMATOCRITO\s*:\s*([\d.,]+)\s*%",
    ],
    "HCM": [
        r"HEMOGLOBINA CORPUSCULAR MEDIA\s*:\s*([\d.,]+)\s*PG",
    ],
    "VCM": [
        r"VOLUME CORPUSCULAR MEDIO\s*:\s*([\d.,]+)\s*FL",
    ],
    "CHCM": [
        r"CONCENTRACAO\s+DE\s+HEMOGLOBINA\s+CORPUSCULAR\s+MEDIA\s*:\s*([\d.,]+)\s*G/DL",
    ],
    "RDW": [
        r"RDW\)\s*:\s*([\d.,]+)\s*%",
        r"RDW\s*:\s*([\d.,]+)\s*%",
    ],
    "LEUCÓCITOS TOTAIS": [
        r"LEUCOCITOS\s+([\d.,]+)\s+\d[\d.,]*\s+A\s+\d[\d.,]*",
        r"LEUCOCITOS\s*:\s*([\d.,]+)",
    ],
    "EOSINÓFILOS": [
        r"EOSINOFILOS\s*:\s*[\d.,]+\s+([\d.,]+)",
        r"EOSINOFILOS\s+[\d.,]+\s+([\d.,]+)",
    ],
    "MONÓCITOS": [
        r"MONOCITOS\s*:\s*[\d.,]+\s+([\d.,]+)",
        r"MONOCITOS\s+[\d.,]+\s+([\d.,]+)",
    ],
    "PLAQUETAS": [
        r"TOTAL DE PLAQUETAS:\s*([\d.,]+)\s*/MM3",
    ],
    "VPM PLAQUETAS": [
        r"VOLUME PLAQUETARIO MEDIO:\s*([\d.,]+)\s*FL",
    ],
    "HEMOGLOBINA A1": [
        r"HEMOGLOBINA A\s*:\s*([\d.,]+)\s*%",
    ],
    "HEMOGLOBINA A2": [
        r"HEMOGLOBINA A2\s*:\s*([\d.,]+)\s*%",
    ],
    "HEMOGLOBINA F": [
        r"HEMOGLOBINA F\s*:\s*([\d.,]+)\s*%",
    ],
}


def normalize_text(text: str) -> str:
    text = unidecode(text or "").upper()
    text = text.replace("\r", "")
    text = text.replace("–", "-").replace("—", "-").replace("‑", "-")
    text = re.sub(r"[ \t]+", " ", text)
    return text



def normalize_key(text: str) -> str:
    text = normalize_text(text)
    text = re.sub(r"[^A-Z0-9/% ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text



def convert_value(raw: str):
    raw = str(raw).strip().replace(" ", "")
    if not raw:
        return raw

    if re.fullmatch(r"\d{1,3}(\.\d{3})+", raw):
        return int(raw.replace(".", ""))

    if "," in raw and "." not in raw:
        raw = raw.replace(",", ".")
        try:
            value = float(raw)
            return int(value) if value.is_integer() else value
        except ValueError:
            return raw

    if "." in raw and "," in raw and raw.rfind(",") > raw.rfind("."):
        raw = raw.replace(".", "").replace(",", ".")
        try:
            value = float(raw)
            return int(value) if value.is_integer() else value
        except ValueError:
            return raw

    if "." in raw and "," in raw and raw.rfind(".") > raw.rfind(","):
        raw = raw.replace(",", "")
        try:
            value = float(raw)
            return int(value) if value.is_integer() else value
        except ValueError:
            return raw

    try:
        value = float(raw)
        return int(value) if value.is_integer() else value
    except ValueError:
        return raw



def extract_pdf_text(pdf_stream) -> str:
    parts: List[str] = []
    data = pdf_stream.read() if hasattr(pdf_stream, "read") else pdf_stream
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            parts.append(page.get_text("text"))
    return "\n".join(parts)



def parse_exams_from_text(text: str):
    normalized = normalize_text(text)
    extracted = {}
    not_found = []

    for canonical_name, patterns in EXAM_PATTERNS.items():
        value_found = None
        matched_pattern = None

        for pattern in patterns:
            match = re.search(pattern, normalized, flags=re.S)
            if match:
                value_found = convert_value(match.group(1))
                matched_pattern = pattern
                break

        if value_found is not None:
            extracted[canonical_name] = {
                "value": value_found,
                "pattern_used": matched_pattern,
            }
        else:
            not_found.append(canonical_name)

    return extracted, not_found



def extract_collection_date(text: str) -> dt.date:
    normalized = normalize_text(text)
    matches = re.findall(r"RECEBIDO/COLETADO\s+EM:\s*(\d{2}/\d{2}/\d{4})", normalized)
    if not matches:
        matches = re.findall(r"DATA DA FICHA\s*(\d{2}/\d{2}/\d{4})", normalized)
    if matches:
        chosen = Counter(matches).most_common(1)[0][0]
        return dt.datetime.strptime(chosen, "%d/%m/%Y").date()
    return dt.date.today()



def build_row_map(ws):
    row_map = {}
    available_labels = []

    for row in range(1, ws.max_row + 1):
        label = ws.cell(row, 2).value  # coluna B
        if label:
            key = normalize_key(str(label))
            row_map[key] = row
            available_labels.append((key, row, str(label)))

    return row_map, available_labels



def find_row_for_exam(canonical_name: str, row_map, available_labels):
    key = normalize_key(canonical_name)

    if key in row_map:
        return row_map[key], 100, canonical_name

    choices = {label: row for label, row, _ in available_labels}
    best = process.extractOne(key, choices.keys(), scorer=fuzz.ratio)

    if best and best[1] >= 85:
        matched_key = best[0]
        return choices[matched_key], best[1], matched_key

    return None, 0, None



def find_data_row(ws) -> int:
    for row in range(1, ws.max_row + 1):
        label = ws.cell(row, 2).value
        if normalize_key(str(label or "")) == "DATA":
            return row
    raise ValueError("A linha 'DATA' não foi encontrada na coluna B da planilha.")



def next_available_data_column(ws, data_row: int, start_col: int = 3) -> int:
    last_used = start_col - 1
    for col in range(start_col, ws.max_column + 1):
        value = ws.cell(data_row, col).value
        if value not in (None, ""):
            last_used = col
    return max(start_col, last_used + 1)



def clone_column_format(ws, source_col: int, target_col: int):
    src_letter = get_column_letter(source_col)
    dst_letter = get_column_letter(target_col)
    ws.column_dimensions[dst_letter].width = ws.column_dimensions[src_letter].width
    ws.column_dimensions[dst_letter].hidden = ws.column_dimensions[src_letter].hidden

    for row in range(1, ws.max_row + 1):
        src = ws.cell(row, source_col)
        dst = ws.cell(row, target_col)
        if src.has_style:
            dst._style = copy.copy(src._style)
        if src.font:
            dst.font = copy.copy(src.font)
        if src.fill:
            dst.fill = copy.copy(src.fill)
        if src.border:
            dst.border = copy.copy(src.border)
        if src.alignment:
            dst.alignment = copy.copy(src.alignment)
        if src.protection:
            dst.protection = copy.copy(src.protection)
        dst.number_format = src.number_format



def write_to_workbook(xlsx_stream, extracted_data: dict, collection_date: dt.date):
    wb = openpyxl.load_workbook(xlsx_stream)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"A planilha precisa conter a aba '{SHEET_NAME}'.")

    ws = wb[SHEET_NAME]
    row_map, available_labels = build_row_map(ws)
    data_row = find_data_row(ws)
    target_col = next_available_data_column(ws, data_row)
    reference_col = max(3, target_col - 1)

    clone_column_format(ws, reference_col, target_col)

    ws.cell(data_row, target_col).value = collection_date
    ws.cell(data_row, target_col).number_format = "dd/mm/yyyy"

    written = []
    unmatched_in_sheet = []

    for canonical_name, payload in extracted_data.items():
        row, score, matched_label = find_row_for_exam(canonical_name, row_map, available_labels)

        if row is None:
            unmatched_in_sheet.append({
                "exam": canonical_name,
                "value": payload["value"],
            })
            continue

        ws.cell(row, target_col).value = payload["value"]
        written.append({
            "exam": canonical_name,
            "value": payload["value"],
            "row": row,
            "column": get_column_letter(target_col),
            "match_score": score,
            "sheet_match": matched_label,
        })

    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)

    return output_stream, {
        "target_column": get_column_letter(target_col),
        "data_row": data_row,
        "collection_date": collection_date.strftime("%d/%m/%Y"),
        "written_to_excel": written,
        "found_in_pdf_but_not_matched_in_sheet": unmatched_in_sheet,
    }



def process_files(pdf_stream, xlsx_stream, pdf_filename: str = "exames.pdf", xlsx_filename: str = "painel.xlsx"):
    text = extract_pdf_text(pdf_stream)
    extracted, not_found_in_pdf = parse_exams_from_text(text)
    collection_date = extract_collection_date(text)
    workbook_stream, write_info = write_to_workbook(xlsx_stream, extracted, collection_date)

    report = {
        "pdf": pdf_filename,
        "xlsx_template": xlsx_filename,
        "target_column": write_info["target_column"],
        "collection_date": write_info["collection_date"],
        "total_exams_extracted": len(extracted),
        "written_to_excel": write_info["written_to_excel"],
        "not_found_in_pdf": not_found_in_pdf,
        "found_in_pdf_but_not_matched_in_sheet": write_info["found_in_pdf_but_not_matched_in_sheet"],
    }

    return workbook_stream, report



def report_to_json_bytes(report: dict) -> bytes:
    return json.dumps(report, ensure_ascii=False, indent=2).encode("utf-8")
