#!/usr/bin/env python3
"""Etapa 1 · Crédito e Débito: cruza o PDV com o Histórico de vendas do portal Sicredi.

Uso:
    python3 scripts/concilia_cartao.py <pdv_cupons.xlsx> <portal1.xlsx> [portal2.xlsx ...]

Gera saida/concilia_credito_*.xlsx e saida/concilia_debito_*.xlsx. Casamento por loja, valor
bruto e hora: o lançamento no PDV pode vir de 30 min antes até 36 h depois da venda na
máquina (POS não integrado é digitado, às vezes só no dia seguinte). Resumo por dia usa a
data calendário da maquininha de um lado e a do PDV (com virada às 06:00) do outro.
A ponta do banco (D+1, MDR, antecipação) entra na etapa 2.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from concilia_pix import TAXA_SERVICO, TOLERANCIA_CENT, carrega_lojas  # noqa: E402
from pdv_pix import HORA_VIRADA  # noqa: E402
from pdv_pix import ler_pdv  # noqa: E402
from portal_sicredi import ler_portais  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
MODAL = {
    "credito": {"pdv": {"TEF - CREDITO", "CARTAO CREDITO"}, "maq": "Crédito", "rotulo": "Crédito"},
    "debito": {"pdv": {"TEF - DEBITO", "CARTAO DEBITO"}, "maq": "Débito", "rotulo": "Débito"},
}
JANELA_ANTES_MIN, JANELA_DEPOIS_H = 30, 36


def casa(maq: pd.DataFrame, pdv: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    usados: set = set()
    linhas = []
    pdv = pdv.sort_values("dt_pdv")
    for i, m in maq.sort_values("dt").iterrows():
        dif = (pdv["dt_pdv"] - m["dt"]).dt.total_seconds() / 60
        c = pdv[(pdv["cent"] == m["cent"]) & (dif >= -JANELA_ANTES_MIN) & (dif <= JANELA_DEPOIS_H * 60) & (~pdv.index.isin(usados))]
        base = {"Data": m["Data"], "Hora maq": m["dt"].strftime("%H:%M:%S"), "Bandeira": m["Bandeira"], "Canal": m["Canal"],
                "Valor Bruto": m["Valor Bruto"], "MDR": m["Valor da taxa (MDR)"], "Líquido": m["Valor líquido"],
                "Terminal": m["Número do terminal"], "Autorização": m["Cód. de autorização"], "cent": m["cent"], "dt": m["dt"]}
        if len(c) == 0:
            linhas.append({**base, "Status": "MAQUININHA SEM PAR NO PDV"})
            continue
        j = (c["dt_pdv"] - m["dt"]).abs().idxmin()
        usados.add(j)
        a = pdv.loc[j]
        linhas.append({**base, "Status": "CONCILIADO", "Dia PDV": a["Data"], "Hora PDV": a["dt_pdv"].strftime("%H:%M"),
                       "Cupom": a["Cupom"], "Caixa": a["Caixa"], "Forma PDV": a["Desc. pagam."],
                       "Dif. hora (min)": round((a["dt_pdv"] - m["dt"]).total_seconds() / 60)})
    res = pd.DataFrame(linhas)
    for c in ("Dia PDV", "Hora PDV", "Cupom", "Caixa", "Forma PDV", "Dif. hora (min)", "Serviço (R$)", "Dif. valor (R$)"):
        if c not in res.columns:
            res[c] = None
    res["Dia PDV"] = res["Dia PDV"].astype(object)
    sobra = pdv[~pdv.index.isin(usados)].copy()
    # cupom + 10% de serviço: máquina sem par = cupom sem par × 1,10, mesma janela
    if len(res) and len(sobra):
        usados_p: set = set()
        for i, m in res[res["Status"] != "CONCILIADO"].iterrows():
            dif = (sobra["dt_pdv"] - m["dt"]).dt.total_seconds() / 60
            c = sobra[((sobra["cent"] * (1 + TAXA_SERVICO)).round() == m["cent"]) & (dif >= -JANELA_ANTES_MIN) & (dif <= JANELA_DEPOIS_H * 60) & (~sobra.index.isin(usados_p))]
            if len(c) == 0:
                continue
            j = (c["dt_pdv"] - m["dt"]).abs().idxmin()
            usados_p.add(j)
            a = sobra.loc[j]
            res.loc[i, ["Status", "Dia PDV", "Hora PDV", "Cupom", "Caixa", "Forma PDV", "Dif. hora (min)", "Serviço (R$)"]] = [
                f"SERVIÇO NÃO BAIXADO · máquina cobrou R$ {m['cent']/100:.2f} (+{TAXA_SERVICO:.0%}), PDV fechou com R$ {a['cent']/100:.2f}", a["Data"], a["dt_pdv"].strftime("%H:%M"),
                a["Cupom"], a["Caixa"], a["Desc. pagam."], round((a["dt_pdv"] - m["dt"]).total_seconds() / 60),
                (m["cent"] - a["cent"]) / 100]
        sobra = sobra[~sobra.index.isin(usados_p)]
    # tolerância: até R$ 1,00 de diferença fecha
    if len(res) and len(sobra):
        usados_t: set = set()
        for i, m in res[res["Status"] == "MAQUININHA SEM PAR NO PDV"].iterrows():
            tm = m["dt"]
            dif = (sobra["dt_pdv"] - tm).dt.total_seconds() / 60
            c = sobra[((sobra["cent"] - m["cent"]).abs() <= TOLERANCIA_CENT) & (dif >= -JANELA_ANTES_MIN) & (dif <= JANELA_DEPOIS_H * 60) & (~sobra.index.isin(usados_t))]
            if len(c) == 0:
                continue
            j = ((c["cent"] - m["cent"]).abs() * 1e6 + (c["dt_pdv"] - tm).abs().dt.total_seconds()).idxmin()
            usados_t.add(j)
            a = sobra.loc[j]
            res.loc[i, ["Status", "Dia PDV", "Hora PDV", "Cupom", "Caixa", "Forma PDV", "Dif. hora (min)", "Dif. valor (R$)"]] = [
                f"CONCILIADO (dif. R$ {(m['cent'] - a['cent'])/100:+.2f})", a["Data"], a["dt_pdv"].strftime("%H:%M"), a["Cupom"], a["Caixa"], a["Desc. pagam."],
                round((a["dt_pdv"] - tm).total_seconds() / 60), (m["cent"] - a["cent"]) / 100]
        sobra = sobra[~sobra.index.isin(usados_t)]
    return res, sobra


def roda(mod: str, bruto: pd.DataFrame, portais: list[str], lojas: dict) -> None:
    cfg = MODAL[mod]
    pdv_all = bruto[bruto["Desc. pagam."].str.upper().isin(cfg["pdv"])].copy()
    todas, sobras, resumo = [], [], []
    for (filial, ec), df in ler_portais(portais, lojas).items():
        maq = df[df["Modalidade"].str.startswith(cfg["maq"]) & (df["Status"].str.upper() == "APROVADA")]
        d0, d1 = df["Data"].min(), df["Data"].max()
        p = pdv_all[(pdv_all["Filial"] == filial) & (pdv_all["Data"] >= d0) & (pdv_all["Data"] <= d1)]
        res, sobra = casa(maq, p)
        res["Filial"], res["EC"] = filial, ec
        sobra["Filial"], sobra["EC"] = filial, ec
        todas.append(res); sobras.append(sobra)
        for dia in pd.date_range(d0, d1).date:
            m = maq[maq["Data"] == dia]; q = p[p["Data"] == dia]
            r = res[res["Data"] == dia]
            resumo.append({"Filial": filial, "EC": ec, "Data": dia,
                           "PDV (qtd)": len(q), "PDV (R$)": q["cent"].sum() / 100,
                           "Maquininha (qtd)": len(m), "Maquininha bruto (R$)": m["cent"].sum() / 100,
                           "MDR (R$)": round(m["Valor da taxa (MDR)"].sum(), 2), "Líquido (R$)": round(m["Valor líquido"].sum(), 2),
                           "Casado (R$)": r[r["Status"].str.startswith("CONCILIADO")]["cent"].sum() / 100,
                           "Serviço não baixado (R$)": round(float(r["Serviço (R$)"].sum()), 2) if "Serviço (R$)" in r.columns else 0.0,
                           "Diferença tolerada (R$)": round(float(r["Dif. valor (R$)"].sum()), 2) if "Dif. valor (R$)" in r.columns else 0.0,
                           "Maquininha sem par (R$)": r[~r["Status"].str.startswith("CONCILIADO") & ~r["Status"].str.startswith("SERVIÇO")]["cent"].sum() / 100,
                           "PDV sem par (R$)": sobra[sobra["Data"] == dia]["cent"].sum() / 100,
                           "Δ PDV − maquininha (R$)": (q["cent"].sum() - m["cent"].sum()) / 100})
        n_ok = int(res["Status"].str.startswith("CONCILIADO").sum()) if len(res) else 0
        n_sv = int(res["Status"].str.startswith("SERVIÇO").sum()) if len(res) else 0
        print(f"{cfg['rotulo']:8} {filial} {d0:%d/%m}–{d1:%d/%m}: maquininha {len(maq)} R$ {maq['cent'].sum()/100:,.2f} | "
              f"PDV {len(p)} R$ {p['cent'].sum()/100:,.2f} | casados {n_ok} | serviço não baixado {n_sv} | maq sem par {len(maq)-n_ok-n_sv} | PDV sem par {len(sobra)}")
    res = pd.concat(todas, ignore_index=True); resumo = pd.DataFrame(resumo); sobra = pd.concat(sobras, ignore_index=True)
    out = RAIZ / "saida" / f"concilia_{mod}_{resumo['Data'].min():%Y-%m-%d}_a_{resumo['Data'].max():%Y-%m-%d}.xlsx"
    cols = ["Filial", "EC", "Data", "Hora maq", "Bandeira", "Canal", "Valor Bruto", "MDR", "Líquido", "Terminal", "Autorização",
            "Status", "Dia PDV", "Hora PDV", "Cupom", "Caixa", "Forma PDV", "Dif. hora (min)", "Serviço (R$)", "Dif. valor (R$)"]
    cols_p = ["Filial", "EC", "Data", "Dia movimento", "Hora", "Cupom", "Caixa", "Desc. pagam.", "Vl. pagamento"]
    with pd.ExcelWriter(out) as xw:
        resumo.to_excel(xw, sheet_name="Resumo loja x dia", index=False)
        res[[c for c in cols if c in res.columns]].to_excel(xw, sheet_name="Maquininha x PDV", index=False)
        res[~res["Status"].str.startswith("CONCILIADO") & ~res["Status"].str.startswith("SERVIÇO")][[c for c in cols if c in res.columns]].to_excel(xw, sheet_name="Maquininha sem par", index=False)
        res[res["Status"].str.startswith("SERVIÇO")][[c for c in cols if c in res.columns]].to_excel(xw, sheet_name="Serviço não baixado", index=False)
        sobra[[c for c in cols_p if c in sobra.columns]].to_excel(xw, sheet_name="PDV sem par", index=False)
    print(f"gravado em {out}")


def main() -> None:
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    pdv_arq, portais = sys.argv[1], sys.argv[2:]
    lojas = carrega_lojas()
    bruto = ler_pdv(pdv_arq)
    bruto["dt_pdv"] = pd.to_datetime(bruto["Data"].astype(str) + " " + bruto["Hora"].astype(str).str[:5])
    bruto.loc[bruto["dt_pdv"].dt.hour < HORA_VIRADA, "dt_pdv"] += pd.Timedelta(days=1)
    bruto["Dia movimento"] = bruto["Data"]
    bruto["Data calendário"] = bruto["dt_pdv"].dt.date
    for mod in MODAL:
        roda(mod, bruto, portais, lojas)


if __name__ == "__main__":
    main()
