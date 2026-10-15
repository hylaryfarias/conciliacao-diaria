#!/usr/bin/env python3
"""Extrai as vendas PIX do relatório de cupons do PDV e gera o resumo por loja e dia.

Uso:
    python3 scripts/pdv_pix.py <relatorio_cupons.xlsx> [pasta_saida]

Lê com python-calamine (o openpyxl quebra nos xlsx do portal Sicredi; o do PDV
abre, mas usamos o mesmo leitor para padronizar). Soma em centavos, nunca em float.
"""
import sys
from pathlib import Path

import pandas as pd
from python_calamine import CalamineWorkbook

FORMAS_PIX = {"PIX MAQUININHA", "PIX - TEF", "TEF - PIX"}


def ler_pdv(caminho: str) -> pd.DataFrame:
    wb = CalamineWorkbook.from_path(caminho)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    # cabeçalho achado pelo conteúdo, nunca por índice fixo
    idx = next(i for i, r in enumerate(rows) if "Desc. pagam." in r)
    df = pd.DataFrame(rows[idx + 1 :], columns=rows[idx])
    df = df[df["Filial"].astype(str).str.strip() != ""]
    df["Vl. pagamento"] = pd.to_numeric(df["Vl. pagamento"], errors="coerce").fillna(0)
    df["cent"] = (df["Vl. pagamento"] * 100).round().astype(int)
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True).dt.date
    df["Hora"] = df["Hora"].astype(str).str[:5]
    return df


def pix_do_pdv(df: pd.DataFrame) -> pd.DataFrame:
    pix = df[df["Desc. pagam."].str.upper().isin(FORMAS_PIX)].copy()
    cols = ["Filial", "Caixa", "Data", "Hora", "Cupom", "Desc. pagam.", "Vl. pagamento", "cent", "Chave"]
    cols += [c for c in ("dt_pdv", "Dia movimento") if c in pix.columns]
    return pix[cols].sort_values(["Filial", "Data", "Hora", "Cupom"]).reset_index(drop=True)


def resumo(pix: pd.DataFrame) -> pd.DataFrame:
    g = pix.groupby(["Filial", "Data"])["cent"].agg(qtd="count", cent="sum").reset_index()
    g["Valor PIX PDV"] = g["cent"] / 100
    return g.drop(columns="cent")


def main() -> None:
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    entrada = sys.argv[1]
    saida = Path(sys.argv[2] if len(sys.argv) > 2 else "saida")
    saida.mkdir(parents=True, exist_ok=True)

    df = ler_pdv(entrada)
    pix = pix_do_pdv(df)
    res = resumo(pix)
    d0, d1 = pix["Data"].min(), pix["Data"].max()
    nome = saida / f"pdv_pix_{d0:%Y-%m-%d}_a_{d1:%Y-%m-%d}.xlsx"
    with pd.ExcelWriter(nome) as xw:
        res.to_excel(xw, sheet_name="Resumo loja x dia", index=False)
        pix.drop(columns="cent").to_excel(xw, sheet_name="PIX por cupom", index=False)
        pd.crosstab(pix["Filial"], pix["Desc. pagam."]).to_excel(xw, sheet_name="Formas por loja")

    print(f"Cupons lidos: {len(df):,} | linhas PIX: {len(pix):,} | "
          f"total PIX: R$ {pix['cent'].sum() / 100:,.2f} | período {d0:%d/%m} a {d1:%d/%m}")
    print(f"Lojas: {pix['Filial'].nunique()} | gravado em {nome}")


if __name__ == "__main__":
    main()
