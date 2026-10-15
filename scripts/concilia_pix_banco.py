#!/usr/bin/env python3
"""Etapa 2 · PIX: cruza os PIX da maquininha (e do PDV, onde o portal não cobre) com o extrato.

Uso:
    python3 scripts/concilia_pix_banco.py <extrato.xlsx> <pdv_cupons.xlsx> [portal1.xlsx ...]

Extrato: export do painel de extratos (colunas Data, Descrição, Valor, Loja, Ag./Conta ...).
Entram só as linhas "RECEBIMENTO PIX". Regra medida no ROBS (01–13/10/2026, 54 de 54):
PIX cai D0 em dia útil; sábado, domingo e feriado caem no próximo dia útil. Sem taxa.

Casamento: mesma loja, mesmo valor, crédito na data da venda ou na primeira data
posterior em que há crédito. Venda posterior à última data do extrato = "AGUARDA EXTRATO",
que não é divergência.
"""
import sys
from pathlib import Path

import pandas as pd
from python_calamine import CalamineWorkbook

sys.path.insert(0, str(Path(__file__).parent))
from concilia_pix import HORA_VIRADA, carrega_lojas  # noqa: E402
from pdv_pix import ler_pdv, pix_do_pdv  # noqa: E402
from portal_sicredi import ler_portal, pix_do_portal  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent


def ler_extrato(caminho: str) -> pd.DataFrame:
    wb = CalamineWorkbook.from_path(caminho)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    h = next(i for i, r in enumerate(rows) if "Descrição" in r)
    ex = pd.DataFrame(rows[h + 1 :], columns=rows[h])
    ex = ex[ex["Descrição"].astype(str).str.upper().str.startswith("RECEBIMENTO PIX")].copy()
    ex["Valor"] = pd.to_numeric(ex["Valor"], errors="coerce").fillna(0)
    ex["cent"] = (ex["Valor"] * 100).round().astype(int)
    ex["Data"] = pd.to_datetime(ex["Data"]).dt.date
    ex["Pagador"] = ex["Descrição"].str.replace("RECEBIMENTO PIX", "", regex=False).str.strip()
    return ex.reset_index(drop=True)


def vendas_pix(pdv_arq: str, portais: list[str], lojas: dict) -> pd.DataFrame:
    """Uma linha por PIX vendido. Fonte = maquininha onde houver relatório; senão PDV."""
    b = ler_pdv(pdv_arq)
    b["dt_pdv"] = pd.to_datetime(b["Data"].astype(str) + " " + b["Hora"].astype(str).str[:5])
    b.loc[b["dt_pdv"].dt.hour < HORA_VIRADA, "dt_pdv"] += pd.Timedelta(days=1)
    b["Data"] = b["dt_pdv"].dt.date
    p = pix_do_pdv(b).rename(columns={"dt_pdv": "dt"})[["Filial", "Data", "dt", "Cupom", "cent"]].assign(Fonte="PDV", Ref=lambda x: "cupom " + x["Cupom"].astype(str))
    partes = []
    for arq in portais:
        df, meta = ler_portal(arq)
        filial = lojas.get(meta["EC"])
        if not filial:
            sys.exit(f"EC {meta['EC']} ({arq}) não está em lojas.csv")
        m = pix_do_portal(df)
        m = m[m["Status"].str.upper().isin({"AUTORIZADA", "APROVADA"})]
        partes.append(m[["Data", "dt", "cent"]].assign(Filial=filial, Fonte="Maquininha", Ref="aut. " + m["Cód. de autorização"].astype(str)))
        # dias cobertos pelo portal saem do PDV para não contar duas vezes
        p = p[~((p["Filial"] == filial) & (p["Data"] >= df["Data"].min()) & (p["Data"] <= df["Data"].max()))]
    return pd.concat(partes + [p], ignore_index=True).sort_values(["Filial", "dt"]).reset_index(drop=True)


def casa_banco(v: pd.DataFrame, ex: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    fim = ex["Data"].max()
    usados: set = set()
    out = []
    for _, a in v.iterrows():
        c = ex[(ex["cent"] == a["cent"]) & (ex["Data"] >= a["Data"]) & (~ex.index.isin(usados))]
        if len(c):
            j = c["Data"].idxmin()
            usados.add(j)
            out.append({**a.to_dict(), "Crédito": ex.loc[j, "Data"], "Pagador": ex.loc[j, "Pagador"],
                        "Dias": (ex.loc[j, "Data"] - a["Data"]).days, "Status": "CONCILIADO"})
        elif a["Data"] > fim:
            out.append({**a.to_dict(), "Status": "AGUARDA EXTRATO"})
        else:
            out.append({**a.to_dict(), "Status": "NÃO CAIU NO BANCO"})
    return pd.DataFrame(out), ex[~ex.index.isin(usados)].copy()


def main() -> None:
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    ex_arq, pdv_arq, portais = sys.argv[1], sys.argv[2], sys.argv[3:]
    ex = ler_extrato(ex_arq)
    lojas = carrega_lojas()
    v = vendas_pix(pdv_arq, portais, lojas)
    v = v[v["Filial"].isin(ex["Loja"].unique())]  # só lojas presentes no extrato
    res, sobra = casa_banco(v, ex)

    resumo = []
    for (f, d), g in res.groupby(["Filial", "Data"]):
        resumo.append({
            "Filial": f, "Venda": d, "Fonte": "/".join(sorted(g["Fonte"].unique())),
            "PIX vendido (qtd)": len(g), "PIX vendido (R$)": g["cent"].sum() / 100,
            "Caiu no banco (R$)": g[g["Status"] == "CONCILIADO"]["cent"].sum() / 100,
            "Crédito em": ", ".join(sorted({str(x) for x in g["Crédito"].dropna()})) if "Crédito" in g else "",
            "Aguarda extrato (R$)": g[g["Status"] == "AGUARDA EXTRATO"]["cent"].sum() / 100,
            "Não caiu (R$)": g[g["Status"] == "NÃO CAIU NO BANCO"]["cent"].sum() / 100,
        })
    resumo = pd.DataFrame(resumo)
    d0, d1 = ex["Data"].min(), ex["Data"].max()
    out = RAIZ / "saida" / f"concilia_pix_banco_{d0:%Y-%m-%d}_a_{d1:%Y-%m-%d}.xlsx"
    res["Valor"] = res["cent"] / 100
    cols = ["Filial", "Data", "dt", "Fonte", "Ref", "Valor", "Status", "Crédito", "Dias", "Pagador"]
    with pd.ExcelWriter(out) as xw:
        resumo.to_excel(xw, sheet_name="Resumo loja x dia", index=False)
        res[[c for c in cols if c in res.columns]].to_excel(xw, sheet_name="Venda x banco", index=False)
        res[res["Status"] != "CONCILIADO"][[c for c in cols if c in res.columns]].to_excel(xw, sheet_name="Pendências", index=False)
        sobra[["Loja", "Data", "Descrição", "Valor"]].to_excel(xw, sheet_name="Banco sem venda", index=False)
    n = res["Status"].value_counts()
    print(f"Extrato {d0:%d/%m}–{d1:%d/%m}: {len(ex)} PIX R$ {ex['cent'].sum()/100:,.2f} | vendas PIX {len(res)} | "
          f"conciliado {n.get('CONCILIADO',0)} | aguarda extrato {n.get('AGUARDA EXTRATO',0)} | "
          f"não caiu {n.get('NÃO CAIU NO BANCO',0)} | banco sem venda {len(sobra)}")
    print(f"gravado em {out}\n")
    print(resumo.to_string(index=False))


if __name__ == "__main__":
    main()
