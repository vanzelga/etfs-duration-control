import os
import pandas as pd


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
    output_dir: str,
    data_carteira: str,
    meta_df: pd.DataFrame,
    anbima_df: pd.DataFrame,
    detalhe_df: pd.DataFrame,
    resultado_df: pd.DataFrame,
    erros_df: pd.DataFrame,
) -> str:
    file_name = f"duration_final_{data_carteira.replace('-', '')}.xlsx"
    file_path = os.path.join(output_dir, file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        meta_df.to_excel(writer, sheet_name="base_meta", index=False)
        anbima_df.to_excel(writer, sheet_name="base_anbima", index=False)
        detalhe_df.to_excel(writer, sheet_name="detalhe_ativos", index=False)
        resultado_df.to_excel(writer, sheet_name="resultado_fundos", index=False)
        erros_df.to_excel(writer, sheet_name="erros", index=False)

    autosize_excel_columns(file_path)
    return file_path