import sys
from datetime import datetime, timedelta

from src.metabase_client import fetch_metabase_data
from src.anbima_client import fetch_anbima_data
from src.calculator import build_detalhe_ativos, build_resultado_fundos
from src.excel_exporter import save_output_excel
from src.utils import ensure_dirs


# ============================================================
# CONFIG
# ============================================================
OUTPUT_DIR = r"C:\composicao"
DEBUG_DIR = rf"{OUTPUT_DIR}\debug"


def get_data_carteira() -> str:
    """
    Reads the portfolio date from command-line argument.

    Expected format:
    python main.py 2026-03-25
    """
    if len(sys.argv) >= 2:
        data_carteira = sys.argv[1].strip()

        try:
            datetime.strptime(data_carteira, "%Y-%m-%d")
            return data_carteira
        except ValueError:
            raise ValueError("Invalid date format. Use YYYY-MM-DD.")

    # Fallback: previous calendar day
    return (datetime.today().date() - timedelta(days=1)).strftime("%Y-%m-%d")


def main():
    ensure_dirs(OUTPUT_DIR, DEBUG_DIR)

    data_carteira = get_data_carteira()

    print(f"Buscando Meta para {data_carteira}...")
    meta_df = fetch_metabase_data(data_carteira)
    meta_df["DataCarteira"] = data_carteira
    if meta_df.empty:
        print("Meta sem dados.")
        return
    print(f"Meta ok | linhas: {len(meta_df)}")

    print(f"Buscando ANBIMA para {data_carteira}...")
    anbima_df = fetch_anbima_data(data_carteira, DEBUG_DIR)
    if anbima_df.empty:
        print("ANBIMA sem dados.")
        return
    print(f"ANBIMA ok | linhas: {len(anbima_df)}")

    detalhe_df, erros_df = build_detalhe_ativos(meta_df, anbima_df)
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