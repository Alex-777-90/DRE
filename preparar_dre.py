
import pandas as pd
import numpy as np
import sys

def preparar_dre(input_xlsx, sheet_name="DRE - Mensal ", output_xlsx="DRE_transformado.xlsx", output_csv="DRE_long.csv"):
    # Ler com header correto (linha 4 visual, índice 3)
    df = pd.read_excel(input_xlsx, sheet_name=sheet_name, header=3)

    # Meses PT-BR
    meses_pt = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    mes_cols = [m for m in meses_pt if m in df.columns]

    # Padronizar primeiras colunas
    colunas = list(df.columns)
    if len(colunas) >= 2:
        colunas[0] = "Conta"
        colunas[1] = "Descrição"
        df.columns = colunas

    # Filtrar e limpar
    df_mes = df[["Conta","Descrição"] + mes_cols].copy()
    df_mes = df_mes.dropna(how="all", subset=mes_cols)
    for c in mes_cols:
        df_mes[c] = pd.to_numeric(df_mes[c], errors="coerce")

    # LONG
    dre_long = df_mes.melt(id_vars=["Conta","Descrição"],
                           value_vars=mes_cols,
                           var_name="Mês",
                           value_name="Valor")

    map_mesnum = {"Janeiro":1,"Fevereiro":2,"Março":3,"Abril":4,"Maio":5,"Junho":6,
                  "Julho":7,"Agosto":8,"Setembro":9,"Outubro":10,"Novembro":11,"Dezembro":12}
    dre_long["MesNum"] = dre_long["Mês"].map(map_mesnum)
    dre_long = dre_long[~dre_long["Valor"].isna()].copy()

    # Comparação mês a mês por conta
    dre_long_sorted = dre_long.sort_values(["Conta","MesNum"])
    dre_long_sorted["Valor_Mês_Anterior"] = dre_long_sorted.groupby("Conta")["Valor"].shift(1)
    dre_long_sorted["Dif_Abs"] = dre_long_sorted["Valor"] - dre_long_sorted["Valor_Mês_Anterior"]
    dre_long_sorted["Dif_%"] = np.where(
        dre_long_sorted["Valor_Mês_Anterior"].isna() | (dre_long_sorted["Valor_Mês_Anterior"]==0),
        np.nan,
        dre_long_sorted["Dif_Abs"] / dre_long_sorted["Valor_Mês_Anterior"]
    )

    # Salvar
    dre_long.to_csv(output_csv, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
        df_mes.to_excel(writer, sheet_name="wide_clean", index=False)
        dre_long.to_excel(writer, sheet_name="long", index=False)
        dre_long_sorted.to_excel(writer, sheet_name="comparacao_mensal", index=False)

    print(f"Arquivos gerados:\n- {output_csv}\n- {output_xlsx}")

if __name__ == "__main__":
    # Uso:
    # python prepara_dre.py "caminho/para/arquivo.xlsx" "DRE - Mensal " "saida.xlsx" "saida.csv"
    input_xlsx = sys.argv[1] if len(sys.argv) > 1 else "DRE 2025 agosto 2025.xlsx"
    sheet_name = sys.argv[2] if len(sys.argv) > 2 else "DRE - Mensal "
    output_xlsx = sys.argv[3] if len(sys.argv) > 3 else "DRE_transformado.xlsx"
    output_csv = sys.argv[4] if len(sys.argv) > 4 else "DRE_long.csv"

    preparar_dre(input_xlsx, sheet_name, output_xlsx, output_csv)
