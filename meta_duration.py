import json
import os
from datetime import datetime

import pandas as pd
import requests


OUTPUT_DIR = r"C:\composicao"
META_BASE_URL = "http://metabase-funds/api/public/card/353113d3-65c5-4912-aafa-6aed1ef5cac3/query/json"

# Example: "2026-03-25"
DATA_CARTEIRA = "2026-03-25"


def ensure_output_dir() -> None:
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def build_parameters(data_carteira: str) -> str:
    payload = [
        {
            "type": "date/single",
            "value": data_carteira,
            "target": ["variable", ["template-tag", "DataCarteira"]],
            "id": "08263a25-a495-9bb4-ec3c-d27f3be5a69b",
        }
    ]
    return json.dumps(payload, separators=(",", ":"))


def fetch_metabase_data(data_carteira: str) -> pd.DataFrame:
    params = {"parameters": build_parameters(data_carteira)}

    response = requests.get(META_BASE_URL, params=params, timeout=60)
    response.raise_for_status()

    data = response.json()
    return pd.DataFrame(data)


def save_outputs(df: pd.DataFrame, data_carteira: str) -> None:
    date_tag = datetime.strptime(data_carteira, "%Y-%m-%d").strftime("%Y%m%d")

    excel_path = os.path.join(OUTPUT_DIR, f"meta_posicao_{date_tag}.xlsx")
    json_path = os.path.join(OUTPUT_DIR, f"meta_posicao_{date_tag}.json")

    df.to_excel(excel_path, index=False)
    df.to_json(json_path, orient="records", force_ascii=False, indent=2)

    print(f"Excel salvo em: {excel_path}")
    print(f"JSON salvo em: {json_path}")


def main():
    ensure_output_dir()

    print(f"Buscando Meta para {DATA_CARTEIRA}...")
    df = fetch_metabase_data(DATA_CARTEIRA)

    if df.empty:
        print("Nenhum dado retornado.")
        return

    print(f"Linhas: {len(df)}")
    print("Colunas:")
    print(list(df.columns))

    save_outputs(df, DATA_CARTEIRA)


if __name__ == "__main__":
    main()