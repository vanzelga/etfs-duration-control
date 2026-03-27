import os
import re

import pandas as pd
import requests

from src.utils import parse_iso_date, br_date, ver_date, br_number_to_python

ANBIMA_BASE_URL = "https://www.anbima.com.br/informacoes/ima/ima-carteira-down.asp"
ANBIMA_LANDING_URLS = [
    "https://www.anbima.com.br/informacoes/ima/ima-carteira.asp",
    "https://www.anbima.com.br/informacoes/ima/",
    "https://www.anbima.com.br/",
]

ANBIMA_IDIOMA = "PT"
ANBIMA_CONSULTA = "Carteira"
ANBIMA_DT_REF_VER_FIXO = "20260318"
TIMEOUT = 60
SALVAR_DEBUG = True


def save_text_debug(debug_dir: str, file_name: str, text: str) -> None:
    if not SALVAR_DEBUG:
        return

    os.makedirs(debug_dir, exist_ok=True)

    file_path = os.path.join(debug_dir, file_name)
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(text)


def create_anbima_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Origin": "https://www.anbima.com.br",
            "Referer": "https://www.anbima.com.br/informacoes/ima/ima-carteira.asp",
        }
    )
    return session


def warmup_anbima_session(session: requests.Session) -> None:
    for url in ANBIMA_LANDING_URLS:
        try:
            session.get(url, timeout=TIMEOUT)
        except Exception:
            pass


def build_anbima_payload(ref_date, dt_ref_ver_value: str, indice_value: str) -> dict:
    return {
        "Tipo": "",
        "DataRef": "",
        "Pai": "ima",
        "escolha": "2",
        "Idioma": ANBIMA_IDIOMA,
        "saida": "xml",
        "indice": indice_value,
        "consulta": ANBIMA_CONSULTA,
        "Dt_Ref_Ver": dt_ref_ver_value,
        "Dt_Ref": br_date(ref_date),
    }


def looks_like_valid_anbima_response(text: str) -> bool:
    if not text:
        return False

    text_lower = text.lower()
    return "<html" not in text_lower and "<carteiras" in text_lower and "c_cod_isin" in text_lower


def fetch_anbima_raw(ref_date, debug_dir: str) -> str:
    session = create_anbima_session()
    warmup_anbima_session(session)

    payload_attempts = [
        ("fixed_ver_space_index", build_anbima_payload(ref_date, ANBIMA_DT_REF_VER_FIXO, "ima-b 5+")),
        ("date_ver_space_index", build_anbima_payload(ref_date, ver_date(ref_date), "ima-b 5+")),
        ("fixed_ver_plus_index", build_anbima_payload(ref_date, ANBIMA_DT_REF_VER_FIXO, "ima-b+5+")),
        ("date_ver_plus_index", build_anbima_payload(ref_date, ver_date(ref_date), "ima-b+5+")),
    ]

    last_text = ""

    for attempt_name, payload in payload_attempts:
        response = session.post(ANBIMA_BASE_URL, data=payload, timeout=TIMEOUT)
        response.raise_for_status()

        text = response.text.strip()
        last_text = text

        save_text_debug(debug_dir, f"anbima_{ref_date.strftime('%Y%m%d')}_{attempt_name}.txt", text)

        if looks_like_valid_anbima_response(text):
            return text

    return last_text


def parse_anbima_carteira_text(raw_text: str) -> pd.DataFrame:
    familia_match = re.search(r"<IMA[^>]*Familia=['\"]([^'\"]+)['\"]", raw_text, re.IGNORECASE)
    dt_ref_match = re.search(r"<CARTEIRAS[^>]*DT_REF=['\"]([^'\"]+)['\"]", raw_text, re.IGNORECASE)

    familia = familia_match.group(1).strip() if familia_match else None
    dt_ref = dt_ref_match.group(1).strip() if dt_ref_match else None

    if not dt_ref:
        return pd.DataFrame()

    inner_match = re.search(r"<CARTEIRAS[^>]*>(.*)</CARTEIRAS>", raw_text, re.IGNORECASE | re.DOTALL)
    inner_text = inner_match.group(1) if inner_match else raw_text

    rows = []
    for chunk in inner_text.split("/>"):
        chunk = chunk.strip()
        if not chunk or ("C_Cod_ISIN" not in chunk and "C_Titulo" not in chunk):
            continue

        pairs = re.findall(r"(C_[A-Za-z0-9_]+)\s*=\s*['\"]([^'\"]*)['\"]", chunk)
        if not pairs:
            continue

        row = {k: br_number_to_python(v) for k, v in pairs}
        row["familia"] = familia
        row["dt_ref"] = dt_ref
        rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df["C_Cod_ISIN"] = df["C_Cod_ISIN"].astype(str).str.strip().str.upper()
    return df


def fetch_anbima_data(data_carteira: str, debug_dir: str):
    ref_date = parse_iso_date(data_carteira)
    raw_text = fetch_anbima_raw(ref_date, debug_dir)

    if not looks_like_valid_anbima_response(raw_text):
        raise ValueError("ANBIMA did not return a valid carteira response.")

    df = parse_anbima_carteira_text(raw_text)

    if df.empty:
        raise ValueError("ANBIMA response was received, but parsing returned no rows.")

    return df, raw_text


def save_anbima_snapshots(df: pd.DataFrame, raw_text: str, output_root: str, data_carteira: str) -> None:
    raw_dir = os.path.join(output_root, "anbima_raw")
    parsed_dir = os.path.join(output_root, "anbima_parsed")

    os.makedirs(raw_dir, exist_ok=True)
    os.makedirs(parsed_dir, exist_ok=True)

    date_tag = data_carteira.replace("-", "")

    raw_path = os.path.join(raw_dir, f"anbima_{date_tag}.txt")
    parsed_path = os.path.join(parsed_dir, f"anbima_{date_tag}.xlsx")

    with open(raw_path, "w", encoding="utf-8") as f:
        f.write(raw_text)

    df.to_excel(parsed_path, index=False)