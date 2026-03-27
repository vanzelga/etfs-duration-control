import pandas as pd

DESENQUADRAMENTO_LIMITE = 720


def build_detalhe_ativos(meta_df: pd.DataFrame, anbima_df: pd.DataFrame):
    detalhe = meta_df.copy()

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
    grouped = (
        detalhe_df.groupby(["DataCarteira", "CgePortfolio"], dropna=False)
        .agg(
            vl_posicao_total=("vl_posicao", "sum"),
            weighted_value_total=("weighted_value", "sum"),
            qtd_ativos=("CodAtivo", "count"),
            qtd_ntnb=("eh_ntnb", "sum"),
        )
        .reset_index()
    )

    grouped["duration_ponderado"] = grouped["weighted_value_total"] / grouped["vl_posicao_total"]
    grouped["desenquadrado_720"] = grouped["duration_ponderado"] < DESENQUADRAMENTO_LIMITE

    return grouped