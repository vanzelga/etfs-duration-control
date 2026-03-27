import json
import os

import pandas as pd
import requests

META_BASE_URL = "http://metabase-funds/api/public/card/353113d3-65c5-4912-aafa-6aed1ef5cac3/query/json"
META_PARAMETER_ID = "08263a25-a495-9bb4-ec3c-d27f3be5a69b"
TIMEOUT = 60


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


def fetch_metabase_data(data_carteira: str):
    params = {"parameters": build_metabase_parameters(data_carteira)}
    response = requests.get(META_BASE_URL, params=params, timeout=TIMEOUT)
    response.raise_for_status()

    raw_text = response.text
    data = response.json()
    df = pd.DataFrame(data)

    if df.empty:
        return df, raw_text

    df["DataCarteira"] = data_carteira
    df["NuIsin"] = df["NuIsin"].astype(str).str.strip().str.upper()

    return df, raw_text


def save_metabase_snapshots(df: pd.DataFrame, raw_text: str, output_root: str, data_carteira: str) -> None:
    raw_dir = os.path.join(output_root, "meta_raw")
    parsed_dir = os.path.join(output_root, "meta_parsed")

    os.makedirs(raw_dir, exist_ok=True)
    os.makedirs(parsed_dir, exist_ok=True)

    date_tag = data_carteira.replace("-", "")

    raw_path = os.path.join(raw_dir, f"meta_{date_tag}.json")
    parsed_path = os.path.join(parsed_dir, f"meta_{date_tag}.xlsx")

    with open(raw_path, "w", encoding="utf-8") as f:
        f.write(raw_text)

    df.to_excel(parsed_path, index=False)