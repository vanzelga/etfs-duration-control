import os
import sys
from datetime import datetime, timedelta

from src.metabase_client import fetch_metabase_data, save_metabase_snapshots
from src.anbima_client import fetch_anbima_data, save_anbima_snapshots
from src.calculator import build_detalhe_ativos, build_resultado_fundos
from src.excel_exporter import save_output_excel
from src.utils import ensure_dirs


# ============================================================
# CONFIG
# ============================================================
OUTPUT_DIR = r"C:\composicao"
DEBUG_DIR = os.path.join(OUTPUT_DIR, "debug")


def get_data_carteira() -> str:
    if len(sys.argv) >= 2:
        data_carteira = sys.argv[1].strip()

        try:
            datetime.strptime(data_carteira, "%Y-%m-%d")
            return data_carteira
        except ValueError:
            raise ValueError("Invalid date format. Use YYYY-MM-DD.")

    return (datetime.today().date() - timedelta(days=1)).strftime("%Y-%m-%d")


def main():
    ensure_dirs(OUTPUT_DIR, DEBUG_DIR)

    data_carteira = get_data_carteira()

    print(f"Buscando Meta para {data_carteira}...")
    meta_df, meta_raw_text = fetch_metabase_data(data_carteira)
    meta_df["DataCarteira"] = data_carteira

    if meta_df.empty:
        print("Meta sem dados.")
        return

    save_metabase_snapshots(meta_df, meta_raw_text, OUTPUT_DIR, data_carteira)
    print(f"Meta ok | linhas: {len(meta_df)}")

    print(f"Buscando ANBIMA para {data_carteira}...")
    anbima_df, anbima_raw_text = fetch_anbima_data(data_carteira, DEBUG_DIR)

    if anbima_df.empty:
        print("ANBIMA sem dados.")
        return

    save_anbima_snapshots(anbima_df, anbima_raw_text, OUTPUT_DIR, data_carteira)
    print(f"ANBIMA ok | linhas: {len(anbima_df)}")

    detalhe_df, erros_df = build_detalhe_ativos(meta_df, anbima_df)

    if not erros_df.empty:
        raise ValueError(
            f"Foram encontrados {len(erros_df)} casos de NTNB elegivel sem correspondencia na ANBIMA. Verifique a base e rode novamente."
        )

    resultado_df = build_resultado_fundos(detalhe_df)

    output_file = save_output_excel(
        output_dir=OUTPUT_DIR,
        data_carteira=data_carteira,
        meta_df=meta_df,
        anbima_df=anbima_df,
        detalhe_df=detalhe_df,
        resultado_df=resultado_df,
        erros_df=erros_df,
    )

    print("\nConcluído.")
    print(f"Arquivo: {output_file}")
    print(f"Detalhe ativos: {len(detalhe_df)}")
    print(f"Resultado fundos: {len(resultado_df)}")
    print(f"Erros: {len(erros_df)}")


if __name__ == "__main__":
    main()