#!/usr/bin/env python3
"""Etapa 1 · PIX: cruza o relatório de cupons do PDV com o Histórico de vendas do portal Sicredi.

Uso:
    python3 scripts/concilia_pix.py <pdv_cupons.xlsx> <portal1.xlsx> [portal2.xlsx ...]

A loja de cada arquivo do portal é identificada pelo EC (Estabelecimento comercial) via
`lojas.csv` (colunas: EC, Filial). Casamento: mesma loja, mesmo dia, mesmo valor em
centavos, hora mais próxima. A hora do PDV é a do fechamento do cupom, que pode ser
bem depois do pagamento na máquina, por isso a hora só desempata — não bloqueia.

Só contam PIX autorizados/aprovados na maquininha. Recusado e cancelado ficam fora e
são listados à parte. PIX sem taxa: valor previsto = bruto.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from pdv_pix import ler_pdv, pix_do_pdv  # noqa: E402
from portal_sicredi import ler_portal, pix_do_portal  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
OK_STATUS = {"AUTORIZADA", "APROVADA"}


def carrega_lojas() -> dict:
    f = RAIZ / "lojas.csv"
    if not f.exists():
        return {}
    t = pd.read_csv(f, dtype=str)
    return dict(zip(t["EC"].str.strip(), t["Filial"].str.strip()))


def casa(pdv: pd.DataFrame, maq: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Casa linha a linha. Retorna (resultado por linha do PDV, linhas da máquina sem par)."""
    usados: set = set()
    linhas = []
    pdv = pdv.sort_values("dt_pdv")
    for _, a in pdv.iterrows():
        cand = maq[(maq["cent"] == a["cent"]) & (maq["Data"] == a["Data"]) & (~maq.index.isin(usados))]
        if len(cand) == 0:
            linhas.append({**a.to_dict(), "Status": "SEM PAR NA MAQUININHA"})
            continue
        j = (cand["dt"] - a["dt_pdv"]).abs().idxmin()
        usados.add(j)
        m = maq.loc[j]
        linhas.append({
            **a.to_dict(),
            "Hora maq": m["dt"].strftime("%H:%M:%S"),
            "Autorização": m["Cód. de autorização"],
            "Terminal": m["Número do terminal"],
            "Comprovante": m["Comprovante de venda"],
            "Dif. hora (min)": round((a["dt_pdv"] - m["dt"]).total_seconds() / 60),
            "Status": "CONCILIADO",
        })
    sobra = maq[~maq.index.isin(usados)].copy()
    return pd.DataFrame(linhas), sobra


def main() -> None:
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    pdv_arq, portais = sys.argv[1], sys.argv[2:]
    lojas = carrega_lojas()

    pdv_all = pix_do_pdv(ler_pdv(pdv_arq))
    pdv_all["dt_pdv"] = pd.to_datetime(
        pdv_all["Data"].astype(str) + " " + pdv_all["Hora"].astype(str).str[:5]
    )

    todas, sobras, recusados, resumo = [], [], [], []
    for arq in portais:
        df, meta = ler_portal(arq)
        ec = meta["EC"]
        filial = lojas.get(ec)
        if not filial:
            sys.exit(f"EC {ec} ({arq}) não está em lojas.csv — adicione a linha EC,Filial")
        maq = pix_do_portal(df)
        ok = maq[maq["Status"].str.upper().isin(OK_STATUS)]
        rec = maq[~maq["Status"].str.upper().isin(OK_STATUS)]
        d0, d1 = df["Data"].min(), df["Data"].max()
        p = pdv_all[(pdv_all["Filial"] == filial) & (pdv_all["Data"] >= d0) & (pdv_all["Data"] <= d1)]

        res, sobra = casa(p, ok)
        res["EC"] = ec
        sobra["Filial"] = filial
        rec = rec.assign(Filial=filial)
        todas.append(res); sobras.append(sobra); recusados.append(rec)

        for dia, g in res.groupby("Data"):
            conc = g[g["Status"] == "CONCILIADO"]["cent"].sum()
            sem_par = g[g["Status"] != "CONCILIADO"]["cent"].sum()
            maq_dia = ok[ok["Data"] == dia]["cent"].sum()
            sobra_dia = sobra[sobra["Data"] == dia]["cent"].sum()
            resumo.append({
                "Filial": filial, "EC": ec, "Data": dia,
                "PIX PDV (qtd)": len(g), "PIX PDV (R$)": g["cent"].sum() / 100,
                "PIX maquininha (qtd)": int((ok["Data"] == dia).sum()), "PIX maquininha (R$)": maq_dia / 100,
                "Conciliado (R$)": conc / 100,
                "PDV sem par (R$)": sem_par / 100,
                "Maquininha sem par (R$)": sobra_dia / 100,
                "Δ PDV − maquininha (R$)": (g["cent"].sum() - maq_dia) / 100,
            })
        # dias que só existem na máquina
        for dia in sorted(set(ok["Data"]) - set(res["Data"])):
            maq_dia = ok[ok["Data"] == dia]["cent"].sum()
            resumo.append({"Filial": filial, "EC": ec, "Data": dia, "PIX PDV (qtd)": 0, "PIX PDV (R$)": 0,
                           "PIX maquininha (qtd)": int((ok["Data"] == dia).sum()), "PIX maquininha (R$)": maq_dia / 100,
                           "Conciliado (R$)": 0, "PDV sem par (R$)": 0, "Maquininha sem par (R$)": maq_dia / 100,
                           "Δ PDV − maquininha (R$)": -maq_dia / 100})
        print(f"{filial} (EC {ec}) {d0:%d/%m}–{d1:%d/%m}: PDV {len(p)} PIX R$ {p['cent'].sum()/100:,.2f} | "
              f"maquininha {len(ok)} PIX R$ {ok['cent'].sum()/100:,.2f} | "
              f"sem par PDV {int((res['Status']!='CONCILIADO').sum())} | sem par maq {len(sobra)} | recusados {len(rec)}")

    res = pd.concat(todas, ignore_index=True)
    resumo = pd.DataFrame(resumo).sort_values(["Filial", "Data"])
    sobra = pd.concat(sobras, ignore_index=True)
    rec = pd.concat(recusados, ignore_index=True)

    col_pdv = ["Filial", "EC", "Data", "Hora", "Cupom", "Caixa", "Desc. pagam.", "Vl. pagamento",
               "Status", "Hora maq", "Dif. hora (min)", "Autorização", "Terminal", "Comprovante"]
    col_maq = ["Filial", "EC", "Data", "dt", "Cód. de autorização", "Valor Bruto", "Status",
               "Número do terminal", "Comprovante de venda", "Canal"]
    d0, d1 = resumo["Data"].min(), resumo["Data"].max()
    out = RAIZ / "saida" / f"concilia_pix_{d0:%Y-%m-%d}_a_{d1:%Y-%m-%d}.xlsx"
    with pd.ExcelWriter(out) as xw:
        resumo.to_excel(xw, sheet_name="Resumo loja x dia", index=False)
        res[[c for c in col_pdv if c in res.columns]].to_excel(xw, sheet_name="PDV x maquininha", index=False)
        res[res["Status"] != "CONCILIADO"][[c for c in col_pdv if c in res.columns]].to_excel(
            xw, sheet_name="PDV sem par", index=False)
        sobra[[c for c in col_maq if c in sobra.columns]].to_excel(xw, sheet_name="Maquininha sem par", index=False)
        rec[[c for c in col_maq if c in rec.columns]].to_excel(xw, sheet_name="Recusados-cancelados", index=False)
    print(f"\ngravado em {out}")
    print(resumo.to_string(index=False))


if __name__ == "__main__":
    main()
