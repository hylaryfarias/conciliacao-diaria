#!/usr/bin/env python3
"""Etapa 2 · PIX: cruza os PIX da maquininha (e do PDV, onde o portal não cobre) com o extrato.

Uso:
    python3 scripts/concilia_pix_banco.py <extrato.xlsx> <pdv_cupons.xlsx> [portal1.xlsx ...] [--de AAAA-MM-DD] [--ate AAAA-MM-DD]

Sem --de/--ate, o período é o coberto pelos relatórios do portal (quando houver).

Extrato: export do painel de extratos (colunas Data, Descrição, Valor, Loja, Ag./Conta ...).
Entram só as linhas "RECEBIMENTO PIX". Regra medida no ROBS (01–13/10/2026, 54 de 54):
PIX cai D0 em dia útil; sábado, domingo e feriado caem no próximo dia útil. Sem taxa.

Casamento: mesma loja, mesmo valor, crédito na data da venda ou na primeira data
posterior em que há crédito. Venda posterior à última data do extrato = "AGUARDA EXTRATO",
que não é divergência.
"""
import glob
import sys
from pathlib import Path

import pandas as pd
from python_calamine import CalamineWorkbook

sys.path.insert(0, str(Path(__file__).parent))
from concilia_pix import HORA_VIRADA, carrega_lojas  # noqa: E402
from pdv_pix import ler_pdv, pix_do_pdv  # noqa: E402
from portal_sicredi import ler_portal, ler_tef, pix_do_portal  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent


def chave_loja(nome: str) -> str:
    """'BIGGS 07 - JAMIL E ACAI' e 'BIGGS 07 E AÇAI' -> 'BIGGS 07'; 'ROBS' -> 'ROBS'."""
    t = str(nome).upper().replace("-", " ").split()
    return " ".join(t[:2]) if len(t) > 1 and t[1].isdigit() else t[0] if t else ""


def ler_extrato(caminho: str, filiais_pdv=()) -> pd.DataFrame:
    """Só 'Vendas no Pix' (RECEBIMENTO PIX). Estornos, VR, transferências e demais categorias ficam fora.
    A coluna Loja é traduzida para o nome da filial no PDV pela chave 'MARCA NN'."""
    wb = CalamineWorkbook.from_path(caminho)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    h = next(i for i, r in enumerate(rows) if "Descrição" in r)
    ex = pd.DataFrame(rows[h + 1 :], columns=rows[h])
    ex = ex[ex["Descrição"].astype(str).str.upper().str.startswith("RECEBIMENTO PIX")].copy()
    if "Categoria" in ex.columns and ex["Categoria"].astype(str).str.contains("Vendas no Pix").any():
        ex = ex[ex["Categoria"].astype(str) == "Vendas no Pix"]
    ex["Valor"] = pd.to_numeric(ex["Valor"], errors="coerce").fillna(0)
    ex["cent"] = (ex["Valor"] * 100).round().astype(int)
    ex["Data"] = pd.to_datetime(ex["Data"]).dt.date
    ex["Pagador"] = ex["Descrição"].str.replace("RECEBIMENTO PIX", "", regex=False).str.strip()
    mapa = {chave_loja(f): f for f in filiais_pdv}
    ex["Loja"] = ex["Loja"].map(lambda x: mapa.get(chave_loja(x), str(x)))
    return ex.reset_index(drop=True)


def ler_devolucoes(caminho: str, filiais_pdv=()) -> pd.DataFrame:
    """DEVOLUCAO PIX: a loja devolveu um PIX ao cliente (pagamento em duplicidade, em geral)."""
    wb = CalamineWorkbook.from_path(caminho)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    h = next(i for i, r in enumerate(rows) if "Descrição" in r)
    ex = pd.DataFrame(rows[h + 1 :], columns=rows[h])
    ex = ex[ex["Descrição"].astype(str).str.upper().str.startswith("DEVOLUCAO PIX")].copy()
    ex["Valor"] = pd.to_numeric(ex["Valor"], errors="coerce").fillna(0)
    ex["cent"] = (ex["Valor"].abs() * 100).round().astype(int)
    ex["Data"] = pd.to_datetime(ex["Data"]).dt.date
    ex["doc"] = ex["Descrição"].str.extract(r"DEVOLUCAO PIX (\d+)")[0]
    mapa = {chave_loja(f): f for f in filiais_pdv}
    ex["Loja"] = ex["Loja"].map(lambda x: mapa.get(chave_loja(x), str(x)))
    return ex.reset_index(drop=True)


def pdv_sem_par_pix() -> pd.DataFrame:
    """Baixas PIX do PDV que ficaram sem par na máquina (etapa 1), para casar direto com o banco."""
    fs = sorted(glob.glob(str(RAIZ / "saida" / "concilia_pix_2*.xlsx")))
    if not fs:
        return pd.DataFrame()
    df = pd.concat([pd.read_excel(f, sheet_name="PDV sem par") for f in fs], ignore_index=True)
    df = df[df["Status"].astype(str).str.startswith("SEM PAR")].copy()
    df["Data"] = pd.to_datetime(df["Data"]).dt.date
    df["cent"] = (pd.to_numeric(df["Vl. pagamento"]) * 100).round().astype(int)
    return df


def classifica_sobra(sobra: pd.DataFrame, vendas: pd.DataFrame, devol: pd.DataFrame) -> pd.DataFrame:
    """Crédito sem venda própria: duplicidade (há venda de mesmo valor na loja, já paga por outro crédito),
    devolvido (há DEVOLUCAO PIX do mesmo pagador e valor) ou outra natureza."""
    sobra = sobra.copy()
    sobra["doc"] = sobra["Descrição"].str.extract(r"RECEBIMENTO PIX (\d+)")[0]
    sp = pdv_sem_par_pix()
    usados_sp: set = set()
    cls, evid, cupons = [], [], []
    for _, x in sobra.iterrows():
        d = x["Data"]
        dias = {d - pd.Timedelta(days=k) for k in range(0, 4)}
        v = vendas[(vendas["Filial"] == x["Loja"]) & (vendas["cent"] == x["cent"]) & (vendas["Data"].isin(dias))]
        dv = devol[(devol["Loja"] == x["Loja"]) & (devol["cent"] == x["cent"]) & (devol["doc"] == x["doc"]) & (devol["Data"] >= d)]
        c = sp[(sp["Filial"] == x["Loja"]) & (sp["cent"] == x["cent"]) & (sp["Data"].isin(dias)) & (~sp.index.isin(usados_sp))] if len(sp) else sp
        if len(dv):
            cls.append("PIX EM DUPLICIDADE · DEVOLVIDO"); evid.append(f"devolvido em {dv['Data'].iloc[0]:%d/%m} ao mesmo pagador")
        elif len(c):
            j = c.index[0]; usados_sp.add(j); r = sp.loc[j]
            cls.append("PIX PAGO · MÁQUINA NÃO REGISTROU"); evid.append(f"baixa PIX no PDV: cupom {r['Cupom']} de {r['Data']:%d/%m} {str(r['Hora'])[:5]} ({r['Caixa']}), sem aprovação na máquina"); cupons.append(str(r["Cupom"]))
        elif len(v):
            r = v.iloc[0]
            cls.append("PIX EM DUPLICIDADE · A DEVOLVER"); evid.append(f"venda de mesmo valor em {r['Data']:%d/%m} {r['dt']:%H:%M} ({r['Fonte']}, {r['Ref']}) já paga por outro crédito")
        else:
            cls.append("OUTRA NATUREZA"); evid.append("sem venda de mesmo valor na loja em 4 dias; pagador " + ("CNPJ" if len(str(x["doc"])) == 14 else "CPF"))
        if cls[-1] != "PIX PAGO · MÁQUINA NÃO REGISTROU":
            cupons.append("")
    sobra["Classificação"] = cls; sobra["Evidência"] = evid; sobra["Cupom PDV"] = cupons
    return sobra


def vendas_pix(pdv_arq: str, portais: list[str], lojas: dict, tef: pd.DataFrame = pd.DataFrame()) -> pd.DataFrame:
    """Uma linha por PIX vendido. Fonte = maquininha onde houver relatório; senão PDV."""
    b = ler_pdv(pdv_arq)
    b["dt_pdv"] = pd.to_datetime(b["Data"].astype(str) + " " + b["Hora"].astype(str).str[:5])
    b.loc[b["dt_pdv"].dt.hour < HORA_VIRADA, "dt_pdv"] += pd.Timedelta(days=1)
    b["Dia movimento"] = b["Data"]
    b["Data"] = b["dt_pdv"].dt.date
    p = pix_do_pdv(b).rename(columns={"dt_pdv": "dt"})[["Filial", "Data", "Dia movimento", "dt", "Cupom", "cent"]].assign(Fonte="PDV", Ref=lambda x: "cupom " + x["Cupom"].astype(str))
    partes = []
    for arq in portais:
        df, meta = ler_portal(arq)
        filial = lojas.get(meta["EC"])
        if not filial:
            sys.exit(f"EC {meta['EC']} ({arq}) não está em lojas.csv")
        m = pix_do_portal(df)
        m = m[m["Status"].str.upper().isin({"AUTORIZADA", "APROVADA"})]
        partes.append(m[["Data", "dt", "cent"]].assign(Filial=filial, Fonte="Maquininha", Ref="aut. " + m["Cód. de autorização"].astype(str)))
        if len(tef):
            t = tef[(tef["Filial"] == filial) & (tef["Data"] >= df["Data"].min()) & (tef["Data"] <= df["Data"].max()) & (tef["Status"].str.upper() == "APROVADA")]
            partes.append(t[["Data", "dt", "cent"]].assign(Filial=filial, Fonte="Maquininha", Ref="TEF " + t["dt"].dt.strftime("%H:%M:%S")))
        # dias cobertos pelo portal saem do PDV para não contar duas vezes
        p = p[~((p["Filial"] == filial) & (p["Data"] >= df["Data"].min()) & (p["Data"] <= df["Data"].max()))]
    return pd.concat(partes + [p], ignore_index=True).sort_values(["Filial", "dt"]).reset_index(drop=True)


def casa_banco(v: pd.DataFrame, ex: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    fim = ex["Data"].max()
    usados: set = set()
    out = []
    for _, a in v.iterrows():
        c = ex[(ex["Loja"] == a["Filial"]) & (ex["cent"] == a["cent"]) & (ex["Data"] >= a["Data"]) & (~ex.index.isin(usados))]
        if len(c):
            j = c["Data"].idxmin()
            usados.add(j)
            out.append({**a.to_dict(), "Crédito": ex.loc[j, "Data"], "Pagador": ex.loc[j, "Pagador"],
                        "Dias": (ex.loc[j, "Data"] - a["Data"]).days, "Status": "CONCILIADO"})
        elif a["Data"] > fim:
            out.append({**a.to_dict(), "Status": "AGUARDA EXTRATO"})
        elif a["Fonte"] == "PDV":
            out.append({**a.to_dict(), "Status": "SEM PAR NO BANCO · AGUARDA MAQUININHA"})
        else:
            out.append({**a.to_dict(), "Status": "NÃO CAIU NO BANCO"})
    return pd.DataFrame(out), ex[~ex.index.isin(usados)].copy()


def main() -> None:
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    args = sys.argv[1:]
    de = ate = None
    if "--de" in args:
        i = args.index("--de"); de = pd.Timestamp(args[i + 1]).date(); del args[i:i + 2]
    if "--ate" in args:
        i = args.index("--ate"); ate = pd.Timestamp(args[i + 1]).date(); del args[i:i + 2]
    tef = pd.DataFrame()
    if "--tef" in args:
        i = args.index("--tef"); tef = ler_tef(args[i + 1]); del args[i:i + 2]
    ex_arq, pdv_arq, portais = args[0], args[1], args[2:]
    lojas = carrega_lojas()
    v = vendas_pix(pdv_arq, portais, lojas, tef)
    ex = ler_extrato(ex_arq, v["Filial"].unique())
    v = v[v["Filial"].isin(ex["Loja"].unique())]  # só lojas presentes no extrato
    if portais and de is None and ate is None:
        v = v[v["Fonte"] == "Maquininha"]        # período = o que o portal cobre
    if de:
        v = v[(v["Data"] >= de) | (v["Fonte"].eq("PDV") & (v["Dia movimento"] >= de))]
    if ate:
        v = v[(v["Data"] <= ate) | (v["Fonte"].eq("PDV") & (v["Dia movimento"] <= ate))]
    d0, d1 = v["Data"].min(), v["Data"].max()
    ex = ex[(ex["Data"] >= d0) & (ex["Loja"].isin(v["Filial"].unique()))]   # só o período e as lojas com venda
    res, sobra = casa_banco(v, ex)
    sobra = classifica_sobra(sobra, v, ler_devolucoes(ex_arq, v["Filial"].unique()))

    resumo = []
    for (f, d), g in res.groupby(["Filial", "Data"]):
        resumo.append({
            "Filial": f, "Venda": d, "Fonte": "/".join(sorted(g["Fonte"].unique())),
            "PIX vendido (qtd)": len(g), "PIX vendido (R$)": g["cent"].sum() / 100,
            "Caiu no banco (R$)": g[g["Status"] == "CONCILIADO"]["cent"].sum() / 100,
            "Crédito em": ", ".join(sorted({str(x) for x in g["Crédito"].dropna()})) if "Crédito" in g else "",
            "Aguarda extrato (R$)": g[g["Status"] == "AGUARDA EXTRATO"]["cent"].sum() / 100,
            "Não caiu (R$)": g[g["Status"] == "NÃO CAIU NO BANCO"]["cent"].sum() / 100,
            "PDV sem par no banco (R$)": g[g["Status"].str.startswith("SEM PAR NO BANCO")]["cent"].sum() / 100,
        })
    resumo = pd.DataFrame(resumo)
    out = RAIZ / "saida" / f"concilia_pix_banco_{d0:%Y-%m-%d}_a_{d1:%Y-%m-%d}.xlsx"
    res["Valor"] = res["cent"] / 100
    cols = ["Filial", "Data", "dt", "Fonte", "Ref", "Valor", "Status", "Crédito", "Dias", "Pagador"]
    with pd.ExcelWriter(out) as xw:
        resumo.to_excel(xw, sheet_name="Resumo loja x dia", index=False)
        res[[c for c in cols if c in res.columns]].to_excel(xw, sheet_name="Venda x banco", index=False)
        res[res["Status"] != "CONCILIADO"][[c for c in cols if c in res.columns]].to_excel(xw, sheet_name="Pendências", index=False)
        sobra[["Loja", "Data", "Descrição", "Valor", "Classificação", "Evidência", "Cupom PDV"]].sort_values(["Loja", "Data"]).to_excel(xw, sheet_name="Banco sem venda", index=False)
    n = res["Status"].value_counts()
    print(f"Vendas {d0:%d/%m}–{d1:%d/%m} · extrato até {ex['Data'].max():%d/%m}: {len(ex)} PIX R$ {ex['cent'].sum()/100:,.2f} | vendas PIX {len(res)} | "
          f"conciliado {n.get('CONCILIADO',0)} | aguarda extrato {n.get('AGUARDA EXTRATO',0)} | "
          f"não caiu {n.get('NÃO CAIU NO BANCO',0)} | PDV sem par no banco {int(res['Status'].str.startswith('SEM PAR NO BANCO').sum())} | banco sem venda {len(sobra)} {sobra['Classificação'].value_counts().to_dict() if len(sobra) else ''}")
    print(f"gravado em {out}\n")
    print(resumo.to_string(index=False))


if __name__ == "__main__":
    main()
