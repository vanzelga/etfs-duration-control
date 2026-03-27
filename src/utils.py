import json
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