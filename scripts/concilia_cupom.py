#!/usr/bin/env python3
"""Cruzamento por TOTAL DA COMANDA, atravessando as formas de pagamento.

Pega os lançamentos do PDV que ficaram sem par (PIX, crédito e débito) e as vendas da máquina
sem par (nas três modalidades). Para cada cupom, soma o que ficou sem par e procura na máquina
um conjunto de 1 a 4 vendas, de qualquer modalidade, que some exatamente esse valor, passadas
de até 36 h antes a 30 min depois do fechamento do cupom e com até 3 h entre si.

Caso típico: mesa dividida, cada cliente pagou de um jeito, o caixa baixou com outra divisão
de formas. A comanda bate; a forma, não.

Uso:
    python3 scripts/concilia_cupom.py            # lê saida/concilia_{pix,credito,debito}_*.xlsx
"""
import glob
from itertools import combinations
from pathlib import Path

import pandas as pd

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from pdv_pix import HORA_VIRADA  # noqa: E402


def _dt(d, h):
    """Datetime calendário a partir do dia de movimento + hora: antes da virada (10h) é o dia seguinte."""
    t = pd.Timestamp(f"{pd.to_datetime(d).date()} {str(h)}")
    return t + pd.Timedelta(days=1) if t.hour < HORA_VIRADA else t

RAIZ = Path(__file__).resolve().parent.parent
S = RAIZ / "saida"
JANELA_ANTES_H, JANELA_DEPOIS_MIN, ESPALHA_H, MAX_G, MAX_CAND = 36, 30, 3, 4, 30


def le(padrao, aba):
    fs = sorted(glob.glob(str(S / padrao)))
    return pd.concat([pd.read_excel(f, sheet_name=aba) for f in fs], ignore_index=True) if fs else pd.DataFrame()


def main() -> None:
    # ---- PDV sem par (as três formas), excluindo o que já tem explicação (forma errada, serviço)
    P = []
    px = le("concilia_pix_2*.xlsx", "PDV x maquininha")
    px = px[px["Status"].astype(str).str.startswith("SEM PAR")]
    for _, x in px.iterrows():
        P.append(dict(Forma="PIX", Filial=x["Filial"], Cupom=str(x["Cupom"]), Caixa=x["Caixa"], Data=pd.to_datetime(x["Data"]).date(),
                      dt=_dt(x["Data"], str(x["Hora"])[:5]), cent=int(round(float(x["Vl. pagamento"]) * 100)), ForPDV=x["Desc. pagam."]))
    for mod, nome in (("credito", "Crédito"), ("debito", "Débito")):
        sp = le(f"concilia_{mod}_2*.xlsx", "PDV sem par")
        for _, x in sp.iterrows():
            P.append(dict(Forma=nome, Filial=x["Filial"], Cupom=str(x["Cupom"]), Caixa=x["Caixa"], Data=pd.to_datetime(x["Data"]).date(),
                          dt=_dt(x["Data"], str(x["Hora"])[:5]), cent=int(round(float(x["Vl. pagamento"]) * 100)), ForPDV=x["Desc. pagam."]))
    P = pd.DataFrame(P)
    # ---- máquina sem par (as três modalidades)
    M = []
    mp = le("concilia_pix_2*.xlsx", "Maquininha sem par")
    for _, x in mp.iterrows():
        M.append(dict(Forma="PIX", Filial=x["Filial"], Data=pd.to_datetime(x["Data"]).date(), dt=pd.to_datetime(x["dt"]), cent=int(round(float(x["Valor Bruto"]) * 100)), Aut=str(x["Cód. de autorização"]), Terminal=str(x["Número do terminal"]), Bandeira="PIX"))
    for mod, nome in (("credito", "Crédito"), ("debito", "Débito")):
        sm = le(f"concilia_{mod}_2*.xlsx", "Maquininha sem par")
        for _, x in sm.iterrows():
            M.append(dict(Forma=nome, Filial=x["Filial"], Data=pd.to_datetime(x["Data"]).date(), dt=_dt(x["Data"], str(x["Hora maq"])[:8]), cent=int(round(float(x["Valor Bruto"]) * 100)), Aut=str(x["Autorização"]), Terminal=str(x["Terminal"]), Bandeira=str(x["Bandeira"])))
    M = pd.DataFrame(M)
    print(f"PDV sem par: {len(P)} linhas em {P.groupby(['Filial','Cupom']).ngroups} cupons | máquina sem par: {len(M)}")

    usados_m, resolvidos, linhas = set(), [], []
    for (fil, cup), g in P.groupby(["Filial", "Cupom"]):
        alvo = int(g["cent"].sum()); dtc = g["dt"].max()
        c = M[(M["Filial"] == fil) & (M["Data"] == g["Data"].iloc[0]) & (~M.index.isin(usados_m))].copy()   # mesmo dia de movimento
        dif = (dtc - c["dt"]).dt.total_seconds() / 60
        c = c[(dif >= -JANELA_DEPOIS_MIN) & (dif <= JANELA_ANTES_H * 60)]
        if len(c) == 0:
            continue
        c["dist"] = (c["dt"] - dtc).abs().dt.total_seconds(); c = c.sort_values("dist").head(MAX_CAND)
        idx, cents, dts, dist = list(c.index), list(c["cent"]), list(c["dt"]), list(c["dist"])
        achado = None
        for k in range(1, min(MAX_G, len(idx)) + 1):
            melhor = None
            for comb in combinations(range(len(idx)), k):
                if sum(cents[i] for i in comb) != alvo:
                    continue
                if (max(dts[i] for i in comb) - min(dts[i] for i in comb)) > pd.Timedelta(hours=ESPALHA_H):
                    continue
                d = sum(dist[i] for i in comb)
                if melhor is None or d < melhor[0]:
                    melhor = (d, [idx[i] for i in comb])
            if melhor:
                achado = melhor[1]; break
        if not achado:
            continue
        # só vale se as formas realmente diferem (senão o motor por forma já teria casado)
        mm = M.loc[achado]
        if sorted(g["Forma"]) == sorted(mm["Forma"]) and len(g) == len(mm):
            continue
        usados_m.update(achado)
        tipo = "FORMA ERRADA" if (len(g) == 1 and len(mm) == 1) else "COMANDA"
        resolvidos.append(dict(Tipo=tipo, Filial=fil, Cupom=cup, Data=g["Data"].min(), Hora=dtc.strftime("%H:%M"), Caixa=g["Caixa"].iloc[0], Total=alvo / 100,
                               FormasPDV=" + ".join(f"{r.Forma} {r.cent/100:.2f}" for r in g.itertuples()),
                               FormasMaq=" + ".join(f"{r.Forma} {r.cent/100:.2f} ({r.dt:%H:%M})" for r in mm.itertuples()),
                               Autorizacoes=" + ".join(mm["Aut"]), n_pdv=len(g), n_maq=len(mm)))
        for r in g.itertuples():
            linhas.append(dict(Tipo=tipo, Lado="PDV", Filial=fil, Cupom=cup, Forma=r.Forma, Data=r.Data, Hora=r.dt.strftime("%H:%M"), Valor=r.cent / 100, Ref=r.Cupom, Caixa=r.Caixa, ForPDV=r.ForPDV,
                               FormaMaq=mm["Forma"].iloc[0], DataMaq=mm["dt"].iloc[0].date(), HoraMaq=mm["dt"].iloc[0].strftime("%H:%M"), AutMaq=mm["Aut"].iloc[0], TerminalMaq=mm["Terminal"].iloc[0], BandeiraMaq=mm["Bandeira"].iloc[0]))
        for r in mm.itertuples():
            linhas.append(dict(Tipo=tipo, Lado="Máquina", Filial=fil, Cupom=cup, Forma=r.Forma, Data=r.dt.date(), Hora=r.dt.strftime("%H:%M:%S"), Valor=r.cent / 100, Ref=r.Aut))
    R = pd.DataFrame(resolvidos); L = pd.DataFrame(linhas)
    out = S / "concilia_cupom.xlsx"
    with pd.ExcelWriter(out) as xw:
        R.to_excel(xw, sheet_name="Cupons resolvidos", index=False)
        L.to_excel(xw, sheet_name="Linhas", index=False)
    print(f"cupons resolvidos por total: {len(R)} | linhas PDV: {int((L['Lado']=='PDV').sum()) if len(L) else 0} | linhas máquina: {int((L['Lado']=='Máquina').sum()) if len(L) else 0} | gravado em {out}")
    if len(R):
        print(R.Tipo.value_counts().to_dict())
        print(R.groupby(["Filial", "Tipo"]).size().unstack(fill_value=0).to_string())
        print(R[["Filial", "Cupom", "Data", "Hora", "Total", "FormasPDV", "FormasMaq"]].head(12).to_string(index=False))


if __name__ == "__main__":
    main()
