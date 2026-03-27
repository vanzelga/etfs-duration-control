import os
import re
from datetime import datetime


def ensure_dirs(*dirs: str) -> None:
    for directory in dirs:
        os.makedirs(directory, exist_ok=True)


def parse_iso_date(date_str: str):
    return datetime.strptime(date_str, "%Y-%m-%d").date()


def br_date(d) -> str:
    return d.strftime("%d/%m/%Y")


def ver_date(d) -> str:
    return d.strftime("%Y%m%d")


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