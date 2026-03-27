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
DATA_CARTEIRA = "2026-03-25"


def main():
    ensure_dirs(OUTPUT_DIR, DEBUG_DIR)

    print(f"Buscando Meta para {DATA_CARTEIRA}...")
    meta_df = fetch_metabase_data(DATA_CARTEIRA)
    if meta_df.empty:
        print("Meta sem dados.")
        return
    print(f"Meta ok | linhas: {len(meta_df)}")

    print(f"Buscando ANBIMA para {DATA_CARTEIRA}...")
    anbima_df = fetch_anbima_data(DATA_CARTEIRA, DEBUG_DIR)
    if anbima_df.empty:
        print("ANBIMA sem dados.")
        return
    print(f"ANBIMA ok | linhas: {len(anbima_df)}")

    detalhe_df, erros_df = build_detalhe_ativos(meta_df, anbima_df)
    resultado_df = build_resultado_fundos(detalhe_df)

    output_file = save_output_excel(
        output_dir=OUTPUT_DIR,
        data_carteira=DATA_CARTEIRA,
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