import json
import os
import re
from datetime import datetime, date

import pandas as pd
import requests


# ============================================================
# CONFIG
# ============================================================
OUTPUT_DIR = r"C:\composicao"
DEBUG_DIR = os.path.join(OUTPUT_DIR, "debug")

# Example: "2026-03-25"
DATA_CARTEIRA = "2026-03-25"

META_BASE_URL = "http://metabase-funds/api/public/card/353113d3-65c5-4912-aafa-6aed1ef5cac3/query/json"
META_PARAMETER_ID = "08263a25-a495-9bb4-ec3c-d27f3be5a69b"

ANBIMA_BASE_URL = "https://www.anbima.com.br/informacoes/ima/ima-carteira-down.asp"
ANBIMA_LANDING_URLS = [
    "https://www.anbima.com.br/informacoes/ima/ima-carteira.asp",
    "https://www.anbima.com.br/informacoes/ima/",
    "https://www.anbima.com.br/",
]

ANBIMA_IDIOMA = "PT"
ANBIMA_CONSULTA = "Carteira"
ANBIMA_INDICE = "ima-b 5+"
ANBIMA_DT_REF_VER_FIXO = "20260318"

DESENQUADRAMENTO_LIMITE = 720
TIMEOUT = 60
SALVAR_DEBUG = True


# ============================================================
# HELPERS
# ============================================================
def ensure_dirs() -> None:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(DEBUG_DIR, exist_ok=True)


def parse_iso_date(date_str: str) -> date:
    return datetime.strptime(date_str, "%Y-%m-%d").date()


def br_date(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def ver_date(d: date) -> str:
    return d.strftime("%Y%m%d")


def save_text_debug(file_name: str, text: str) -> None:
    if not SALVAR_DEBUG:
        return

    file_path = os.path.join(DEBUG_DIR, file_name)
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(text)


def br_number_to_python(value: str):
    value = value.strip()

    if value == "":
        return None

    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", value):
        return value

    if re.fullmatch(r"-?\d{1,3}(?:\.\d{3})*(?:,\d+)?", value) or re.fullmatch(r"-?\d+(?:,\d+)?", value):
        normalized = value.replace(".", "").replace(",", ".")
        num = float(normalized)
        return int(num) if num.is_integer() else num

    return value


# ============================================================
# METABASE
# ============================================================
def build_metabase_parameters(data_carteira: str) -> str:
    payload = [
        {
            "type": "date/single",
            "value": data_carteira,
            "target": ["variable", ["template-tag", "DataCarteira"]],
            "id": META_PARAMETER_ID,
        }
    ]
    return json.dumps(payload, separators=(",", ":"))


def fetch_metabase_data(data_carteira: str) -> pd.DataFrame:
    params = {"parameters": build_metabase_parameters(data_carteira)}
    response = requests.get(META_BASE_URL, params=params, timeout=TIMEOUT)
    response.raise_for_status()

    data = response.json()
    df = pd.DataFrame(data)

    if df.empty:
        return df

    df["DataCarteira"] = data_carteira
    df["NuIsin"] = df["NuIsin"].astype(str).str.strip().str.upper()

    return df


# ============================================================
# ANBIMA
# ============================================================
def create_anbima_session() -> requests.Session:
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


def warmup_anbima_session(session: requests.Session) -> None:
    for url in ANBIMA_LANDING_URLS:
        try:
            session.get(url, timeout=TIMEOUT)
        except Exception:
            pass


def build_anbima_payload(ref_date: date, dt_ref_ver_value: str, indice_value: str) -> dict:
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

    if "<html" in text_lower:
        return False

    if "<carteiras" not in text_lower:
        return False

    if "c_cod_isin" not in text_lower:
        return False

    return True


def fetch_anbima_raw(ref_date: date) -> str:
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

        save_text_debug(f"anbima_{ref_date.strftime('%Y%m%d')}_{attempt_name}.txt", text)

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
    df["C_Cod_ISIN"] = df["C_Cod_ISIN"].astype(str).str.strip().str.upper()
    return df


def fetch_anbima_data(data_carteira: str) -> pd.DataFrame:
    ref_date = parse_iso_date(data_carteira)
    raw_text = fetch_anbima_raw(ref_date)

    if not looks_like_valid_anbima_response(raw_text):
        raise ValueError("ANBIMA did not return a valid carteira response.")

    df = parse_anbima_carteira_text(raw_text)

    if df.empty:
        raise ValueError("ANBIMA response was received, but parsing returned no rows.")

    return df


# ============================================================
# BUSINESS RULES
# ============================================================
def build_detalhe_ativos(meta_df: pd.DataFrame, anbima_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    detalhe = meta_df.copy()

    # Eligible assets: NTNB with original duration different from 1
    detalhe["eh_ntnb"] = detalhe["CodTipoAgrupamento"].astype(str).str.upper().eq("NTNB")
    detalhe["duration_original"] = pd.to_numeric(detalhe["Duration"], errors="coerce")
    detalhe["vl_posicao"] = pd.to_numeric(detalhe["VlPosicao"], errors="coerce")

    eligible_mask = detalhe["eh_ntnb"] & detalhe["duration_original"].ne(1)

    anbima_aux = anbima_df[["C_Cod_ISIN", "C_PMR", "C_Duration", "C_Data_Vencimento", "dt_ref"]].copy()
    anbima_aux = anbima_aux.rename(
        columns={
            "C_Cod_ISIN": "NuIsin",
            "C_PMR": "pmr_anbima",
            "C_Duration": "duration_anbima",
            "C_Data_Vencimento": "data_vencimento_anbima",
            "dt_ref": "dt_ref_anbima",
        }
    )

    detalhe = detalhe.merge(anbima_aux, on="NuIsin", how="left")

    detalhe["pmr_anbima"] = pd.to_numeric(detalhe["pmr_anbima"], errors="coerce")
    detalhe["duration_anbima"] = pd.to_numeric(detalhe["duration_anbima"], errors="coerce")

    detalhe["duration_final"] = detalhe["duration_original"]
    detalhe.loc[eligible_mask, "duration_final"] = detalhe.loc[eligible_mask, "pmr_anbima"]

    detalhe["regra_aplicada"] = "mantido_meta"
    detalhe.loc[eligible_mask, "regra_aplicada"] = "substituido_por_pmr_anbima"

    erro_mask = eligible_mask & detalhe["pmr_anbima"].isna()

    erros = detalhe.loc[erro_mask, [
        "DataCarteira",
        "CgePortfolio",
        "CodAtivo",
        "NuIsin",
        "Nickname",
        "CodTipoAgrupamento",
        "duration_original",
        "pmr_anbima",
        "regra_aplicada",
    ]].copy()

    erros["erro"] = "NTNB elegivel sem correspondencia na ANBIMA para a data informada"

    detalhe["weighted_value"] = detalhe["vl_posicao"] * detalhe["duration_final"]

    return detalhe, erros


def build_resultado_fundos(detalhe_df: pd.DataFrame) -> pd.DataFrame:
    base = detalhe_df.copy()

    grouped = (
        base.groupby(["DataCarteira", "CgePortfolio"], dropna=False)
        .agg(
            vl_posicao_total=("vl_posicao", "sum"),
            weighted_value_total=("weighted_value", "sum"),
            qtd_ativos=("CodAtivo", "count"),
            qtd_ntnb=("eh_ntnb", "sum"),
        )
        .reset_index()
    )

    grouped["duration_ponderado"] = grouped["weighted_value_total"] / grouped["vl_posicao_total"]
    grouped["desenquadrado_720"] = grouped["duration_ponderado"] > DESENQUADRAMENTO_LIMITE

    return grouped


# ============================================================
# EXPORT
# ============================================================
def autosize_excel_columns(file_path: str) -> None:
    try:
        from openpyxl import load_workbook

        wb = load_workbook(file_path)

        for ws in wb.worksheets:
            ws.freeze_panes = "A2"

            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter

                for cell in col:
                    value = "" if cell.value is None else str(cell.value)
                    if len(value) > max_len:
                        max_len = len(value)

                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 40)

        wb.save(file_path)
    except Exception:
        pass


def save_output_excel(
    meta_df: pd.DataFrame,
    anbima_df: pd.DataFrame,
    detalhe_df: pd.DataFrame,
    resultado_df: pd.DataFrame,
    erros_df: pd.DataFrame,
    data_carteira: str,
) -> str:
    file_name = f"duration_final_{data_carteira.replace('-', '')}.xlsx"
    file_path = os.path.join(OUTPUT_DIR, file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        meta_df.to_excel(writer, sheet_name="base_meta", index=False)
        anbima_df.to_excel(writer, sheet_name="base_anbima", index=False)
        detalhe_df.to_excel(writer, sheet_name="detalhe_ativos", index=False)
        resultado_df.to_excel(writer, sheet_name="resultado_fundos", index=False)
        erros_df.to_excel(writer, sheet_name="erros", index=False)

    autosize_excel_columns(file_path)
    return file_path


# ============================================================
# MAIN
# ============================================================
def main():
    ensure_dirs()

    print(f"Buscando Meta para {DATA_CARTEIRA}...")
    meta_df = fetch_metabase_data(DATA_CARTEIRA)
    if meta_df.empty:
        print("Meta sem dados.")
        return

    print(f"Meta ok | linhas: {len(meta_df)}")

    print(f"Buscando ANBIMA para {DATA_CARTEIRA}...")
    anbima_df = fetch_anbima_data(DATA_CARTEIRA)
    print(f"ANBIMA ok | linhas: {len(anbima_df)}")

    detalhe_df, erros_df = build_detalhe_ativos(meta_df, anbima_df)
    resultado_df = build_resultado_fundos(detalhe_df)

    output_file = save_output_excel(
        meta_df=meta_df,
        anbima_df=anbima_df,
        detalhe_df=detalhe_df,
        resultado_df=resultado_df,
        erros_df=erros_df,
        data_carteira=DATA_CARTEIRA,
    )

    print("\nConcluido.")
    print(f"Arquivo: {output_file}")
    print(f"Detalhe ativos: {len(detalhe_df)}")
    print(f"Resultado fundos: {len(resultado_df)}")
    print(f"Erros: {len(erros_df)}")


if __name__ == "__main__":
    main()