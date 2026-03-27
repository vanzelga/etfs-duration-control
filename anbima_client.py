import os
import re
from datetime import date, datetime, timedelta
from typing import Optional

import pandas as pd
import requests


# ============================================================
# CONFIG
# ============================================================
OUTPUT_DIR = r"C:\composicao"
DEBUG_DIR = os.path.join(OUTPUT_DIR, "debug")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "ima_b_5_plus_historico.xlsx")

BASE_URL = "https://www.anbima.com.br/informacoes/ima/ima-carteira-down.asp"
LANDING_URLS = [
    "https://www.anbima.com.br/informacoes/ima/ima-carteira.asp",
    "https://www.anbima.com.br/informacoes/ima/",
    "https://www.anbima.com.br/",
]

IDIOMA = "PT"
CONSULTA = "Carteira"
INDICE = "ima-b 5+"

TIMEOUT = 60
DT_REF_VER_FIXO = "20260318"
SALVAR_DEBUG = True

# Date input examples:
# Single day:
# START_DATE = "2026-03-25"
# END_DATE   = "2026-03-25"
#
# Date range:
# START_DATE = "2026-03-20"
# END_DATE   = "2026-03-25"
#
# If both are None, the script fetches the last N business days.
START_DATE = "2026-03-25"
END_DATE = "2026-03-25"

# Fallback when START_DATE and END_DATE are not informed.
DIAS_UTEIS_PARA_BUSCAR = 5


# ============================================================
# FILESYSTEM
# ============================================================
def ensure_dirs() -> None:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(DEBUG_DIR, exist_ok=True)


# ============================================================
# DATE HELPERS
# ============================================================
def br_date(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def ver_date(d: date) -> str:
    return d.strftime("%Y%m%d")


def parse_iso_date(date_str: str) -> date:
    return datetime.strptime(date_str, "%Y-%m-%d").date()


def get_last_business_days(qtd: int, end_date: Optional[date] = None) -> list[date]:
    if end_date is None:
        end_date = date.today()

    dates = []
    current = end_date

    while len(dates) < qtd:
        if current.weekday() < 5:
            dates.append(current)
        current -= timedelta(days=1)

    return sorted(dates)


def get_target_dates() -> list[date]:
    """
    Date selection logic:
    - If START_DATE and END_DATE are informed:
      - same day -> fetch exactly that day
      - range     -> fetch all business days in the inclusive interval
    - If both are None:
      - fetch the last N business days automatically
    """
    if START_DATE and END_DATE:
        start = parse_iso_date(START_DATE)
        end = parse_iso_date(END_DATE)

        if start > end:
            raise ValueError("START_DATE cannot be greater than END_DATE.")

        if start == end:
            return [start]

        dates = []
        current = start
        while current <= end:
            if current.weekday() < 5:
                dates.append(current)
            current += timedelta(days=1)

        return dates

    if START_DATE or END_DATE:
        raise ValueError("Inform both START_DATE and END_DATE, or leave both as None.")

    return get_last_business_days(DIAS_UTEIS_PARA_BUSCAR)


# ============================================================
# REQUEST / SESSION
# ============================================================
def build_payload(ref_date: date, dt_ref_ver_value: str, indice_value: str) -> dict:
    return {
        "Tipo": "",
        "DataRef": "",
        "Pai": "ima",
        "escolha": "2",
        "Idioma": IDIOMA,
        "saida": "xml",
        "indice": indice_value,
        "consulta": CONSULTA,
        "Dt_Ref_Ver": dt_ref_ver_value,
        "Dt_Ref": br_date(ref_date),
    }


def create_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;q=0.9,"
                "image/avif,image/webp,image/apng,*/*;q=0.8"
            ),
            "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Origin": "https://www.anbima.com.br",
            "Referer": "https://www.anbima.com.br/informacoes/ima/ima-carteira.asp",
        }
    )
    return session


def warmup_session(session: requests.Session) -> None:
    for url in LANDING_URLS:
        try:
            session.get(url, timeout=TIMEOUT)
        except Exception:
            pass


def save_debug_response(ref_date: date, attempt_name: str, text: str) -> None:
    if not SALVAR_DEBUG:
        return

    file_name = f"debug_{ref_date.strftime('%Y%m%d')}_{attempt_name}.txt"
    file_path = os.path.join(DEBUG_DIR, file_name)

    with open(file_path, "w", encoding="utf-8") as f:
        f.write(text)


def looks_like_valid_carteira_response(text: str) -> bool:
    if not text:
        return False

    text_lower = text.lower()

    if "<html" in text_lower:
        return False

    if "<carteiras" not in text_lower:
        return False

    if "c_cod_isin" not in text_lower:
        return False

    return True


def fetch_one_day(session: requests.Session, ref_date: date) -> str:
    """
    Tries multiple payload variants because ANBIMA may behave inconsistently
    with Dt_Ref_Ver and index encoding expectations.
    """
    payload_attempts = [
        ("fixed_ver_space_index", build_payload(ref_date, DT_REF_VER_FIXO, "ima-b 5+")),
        ("date_ver_space_index", build_payload(ref_date, ver_date(ref_date), "ima-b 5+")),
        ("fixed_ver_plus_index", build_payload(ref_date, DT_REF_VER_FIXO, "ima-b+5+")),
        ("date_ver_plus_index", build_payload(ref_date, ver_date(ref_date), "ima-b+5+")),
    ]

    last_text = ""

    for attempt_name, payload in payload_attempts:
        response = session.post(BASE_URL, data=payload, timeout=TIMEOUT)
        response.raise_for_status()

        text = response.text.strip()
        last_text = text

        save_debug_response(ref_date, attempt_name, text)

        if looks_like_valid_carteira_response(text):
            return text

    return last_text


# ============================================================
# PARSING
# ============================================================
def br_number_to_python(value: str):
    value = value.strip()

    if value == "":
        return None

    # Date-like string remains text
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", value):
        return value

    # Brazilian number format -> Python number
    if re.fullmatch(r"-?\d{1,3}(?:\.\d{3})*(?:,\d+)?", value) or re.fullmatch(r"-?\d+(?:,\d+)?", value):
        normalized = value.replace(".", "").replace(",", ".")
        num = float(normalized)
        return int(num) if num.is_integer() else num

    return value


def parse_anbima_carteira_text(raw_text: str) -> pd.DataFrame:
    familia_match = re.search(r"<IMA[^>]*Familia=['\"]([^'\"]+)['\"]", raw_text, re.IGNORECASE)
    dt_ref_match = re.search(r"<CARTEIRAS[^>]*DT_REF=['\"]([^'\"]+)['\"]", raw_text, re.IGNORECASE)

    familia = familia_match.group(1).strip() if familia_match else None
    dt_ref = dt_ref_match.group(1).strip() if dt_ref_match else None

    if not dt_ref:
        return pd.DataFrame()

    # Extract only the inner content of the CARTEIRAS block
    inner_match = re.search(r"<CARTEIRAS[^>]*>(.*)</CARTEIRAS>", raw_text, re.IGNORECASE | re.DOTALL)
    inner_text = inner_match.group(1) if inner_match else raw_text

    # Each asset ends with "/>"
    chunks = inner_text.split("/>")

    rows = []
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue

        if "C_Cod_ISIN" not in chunk and "C_Titulo" not in chunk:
            continue

        pairs = re.findall(r"(C_[A-Za-z0-9_]+)\s*=\s*['\"]([^'\"]*)['\"]", chunk)

        if not pairs:
            continue

        row = {k: br_number_to_python(v) for k, v in pairs}
        row["familia"] = familia
        row["dt_ref"] = dt_ref
        row["dt_ref_date"] = datetime.strptime(dt_ref, "%d/%m/%Y").date()

        rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    first_cols = ["familia", "dt_ref", "dt_ref_date"]
    other_cols = [c for c in df.columns if c not in first_cols]
    return df[first_cols + other_cols]


# ============================================================
# HISTORY / MERGE
# ============================================================
def read_existing_history() -> pd.DataFrame:
    if not os.path.exists(OUTPUT_FILE):
        return pd.DataFrame()

    try:
        return pd.read_excel(OUTPUT_FILE, sheet_name="carteira")
    except Exception:
        return pd.DataFrame()


def merge_history(existing: pd.DataFrame, new_data: pd.DataFrame) -> pd.DataFrame:
    if existing.empty and new_data.empty:
        return pd.DataFrame()

    frames = [df for df in [existing, new_data] if not df.empty]
    merged = pd.concat(frames, ignore_index=True, sort=False)

    if "dt_ref_date" in merged.columns:
        merged["dt_ref_date"] = pd.to_datetime(merged["dt_ref_date"], errors="coerce").dt.date

    dedup_keys = [c for c in ["dt_ref", "C_Cod_ISIN"] if c in merged.columns]
    if dedup_keys:
        merged = merged.drop_duplicates(subset=dedup_keys, keep="last")

    sort_cols = [c for c in ["dt_ref_date", "C_Data_Vencimento", "C_Cod_ISIN"] if c in merged.columns]
    if sort_cols:
        merged = merged.sort_values(sort_cols).reset_index(drop=True)

    return merged


# ============================================================
# SAVE EXCEL
# ============================================================
def save_excel(df: pd.DataFrame) -> None:
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="carteira", index=False)

    try:
        from openpyxl import load_workbook

        wb = load_workbook(OUTPUT_FILE)
        ws = wb["carteira"]
        ws.freeze_panes = "A2"

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter

            for cell in col:
                value = "" if cell.value is None else str(cell.value)
                if len(value) > max_len:
                    max_len = len(value)

            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 40)

        wb.save(OUTPUT_FILE)
    except Exception:
        pass


# ============================================================
# MAIN
# ============================================================
def main():
    ensure_dirs()

    session = create_session()
    warmup_session(session)

    dates_to_fetch = get_target_dates()

    frames = []

    print(f"Fetching {len(dates_to_fetch)} date(s)...")
    for ref_date in dates_to_fetch:
        print(f" - {ref_date.isoformat()}")

        try:
            raw_text = fetch_one_day(session, ref_date)

            if not looks_like_valid_carteira_response(raw_text):
                print("   no valid data")
                snippet = raw_text[:300].replace("\n", " ")
                print(f"   response preview: {snippet}")
                continue

            df_day = parse_anbima_carteira_text(raw_text)

            if df_day.empty:
                print("   response received, but parsing returned no rows")
                continue

            frames.append(df_day)
            print(f"   ok | rows: {len(df_day)}")

        except Exception as e:
            print(f"   error: {e}")

    new_data = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
    existing = read_existing_history()
    final_df = merge_history(existing, new_data)

    if final_df.empty:
        print("No data to save.")
        print(f"Raw debug saved to: {DEBUG_DIR}")
        return

    save_excel(final_df)

    print("\nDone.")
    print(f"File: {OUTPUT_FILE}")
    print(f"Total rows: {len(final_df)}")
    print(f"Raw debug: {DEBUG_DIR}")


if __name__ == "__main__":
    main()