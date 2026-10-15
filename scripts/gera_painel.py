#!/usr/bin/env python3
"""Gera painel/index.html a partir das planilhas de saida/ e do relatório do PDV.

Uso:
    python3 scripts/gera_painel.py <pdv_cupons.xlsx> --de AAAA-MM-DD --ate AAAA-MM-DD

Lê saida/concilia_pix_*.xlsx (etapa 1) e saida/concilia_pix_banco_*.xlsx (etapa 2),
monta um JSON e injeta em painel/template.html no marcador /*DADOS*/.
"""
import glob
import json
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from concilia_pix import HORA_VIRADA, carrega_lojas  # noqa: E402
from pdv_pix import ler_pdv, pix_do_pdv  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
DOW = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]
FERIADOS = {date(2026, 10, 12)}
MARCA = {"ROBS": "robs", "MAPOLI": "mapoli"}


def marca(filial: str) -> str:
    for k, v in MARCA.items():
        if filial.startswith(k):
            return v
    return filial.split()[0].lower()


def r2(x) -> float:
    return round(float(x), 2) if pd.notna(x) else 0.0


def main() -> None:
    args = sys.argv[1:]
    de = pd.Timestamp(args[args.index("--de") + 1]).date()
    ate = pd.Timestamp(args[args.index("--ate") + 1]).date()
    pdv_arq = args[0]

    # ---- PDV: todas as lojas, PIX por dia calendário no período
    b = ler_pdv(pdv_arq)
    b["dt_pdv"] = pd.to_datetime(b["Data"].astype(str) + " " + b["Hora"].astype(str).str[:5])
    b.loc[b["dt_pdv"].dt.hour < HORA_VIRADA, "dt_pdv"] += pd.Timedelta(days=1)
    b["Data"] = b["dt_pdv"].dt.date
    p = pix_do_pdv(b)
    p = p[(p["Data"] >= de) & (p["Data"] <= ate)]
    pdv_dia = p.groupby(["Filial", "Data"])["cent"].agg(qtd="count", cent="sum").reset_index()

    # ---- etapa 1 e 2
    e1 = pd.concat([pd.read_excel(f, sheet_name="Resumo loja x dia") for f in glob.glob(str(RAIZ / "saida/concilia_pix_2*.xlsx"))], ignore_index=True) if glob.glob(str(RAIZ / "saida/concilia_pix_2*.xlsx")) else pd.DataFrame()
    e1p = pd.concat([pd.read_excel(f, sheet_name="PDV sem par") for f in glob.glob(str(RAIZ / "saida/concilia_pix_2*.xlsx"))], ignore_index=True) if len(e1) else pd.DataFrame()
    e1m = pd.concat([pd.read_excel(f, sheet_name="Maquininha sem par") for f in glob.glob(str(RAIZ / "saida/concilia_pix_2*.xlsx"))], ignore_index=True) if len(e1) else pd.DataFrame()
    e2 = pd.concat([pd.read_excel(f, sheet_name="Resumo loja x dia") for f in glob.glob(str(RAIZ / "saida/concilia_pix_banco_*.xlsx"))], ignore_index=True) if glob.glob(str(RAIZ / "saida/concilia_pix_banco_*.xlsx")) else pd.DataFrame()
    e2p = pd.concat([pd.read_excel(f, sheet_name="Pendências") for f in glob.glob(str(RAIZ / "saida/concilia_pix_banco_*.xlsx"))], ignore_index=True) if len(e2) else pd.DataFrame()
    e2s = pd.concat([pd.read_excel(f, sheet_name="Banco sem venda") for f in glob.glob(str(RAIZ / "saida/concilia_pix_banco_*.xlsx"))], ignore_index=True) if len(e2) else pd.DataFrame()
    for df, col in ((e1, "Data"), (e2, "Venda"), (e1p, "Data"), (e2p, "Data"), (e2s, "Data")):
        if len(df):
            df[col] = pd.to_datetime(df[col]).dt.date
    ec_de = {v: k for k, v in carrega_lojas().items()}
    extrato_ate = None
    if len(e2):
        creds = pd.to_datetime(e2["Crédito em"].dropna().astype(str).str.split(", ").explode()).dt.date
        extrato_ate = max(creds) if len(creds) else None

    lojas = []
    for filial in sorted(b["Filial"].unique()):
        dias = []
        for d in pd.date_range(de, ate).date:
            row = {"data": d.isoformat(), "dow": DOW[d.weekday()], "feriado": d in FERIADOS}
            x = pdv_dia[(pdv_dia["Filial"] == filial) & (pdv_dia["Data"] == d)]
            row["pdv_qtd"] = int(x["qtd"].sum()); row["pdv"] = r2(x["cent"].sum() / 100)
            y = e1[(e1["Filial"] == filial) & (e1["Data"] == d)] if len(e1) else pd.DataFrame()
            if len(y):
                y = y.iloc[0]
                row.update(maq_qtd=int(y["PIX maquininha (qtd)"]), maq=r2(y["PIX maquininha (R$)"]),
                           conc=r2(y["Conciliado (R$)"]), pdv_sem_par=r2(y["PDV sem par (R$)"]),
                           maq_sem_par=r2(y["Maquininha sem par (R$)"]), delta=r2(y["Δ PDV − maquininha (R$)"]))
            z = e2[(e2["Filial"] == filial) & (e2["Venda"] == d)] if len(e2) else pd.DataFrame()
            if len(z):
                z = z.iloc[0]
                row.update(banco=r2(z["Caiu no banco (R$)"]), credito=str(z["Crédito em"]) if pd.notna(z["Crédito em"]) else "",
                           aguarda=r2(z["Aguarda extrato (R$)"]), nao_caiu=r2(z["Não caiu (R$)"]))
            dias.append(row)
        tem_maq = any("maq" in r for r in dias)
        tem_banco = any("banco" in r for r in dias)
        pend = []
        if len(e1p):
            for _, q in e1p[e1p["Filial"] == filial].iterrows():
                pend.append({"etapa": 1, "data": q["Data"].isoformat(), "valor": r2(q["Vl. pagamento"]), "ref": f"cupom {q['Cupom']} · {q['Hora']} · {q['Caixa']}",
                             "status": str(q["Status"]), "detalhe": f"terminal {q['Terminal']} · aut. {q['Autorização']} · comprovante {q['Comprovante']}" if pd.notna(q.get("Terminal")) else ""})
        if len(e1m):
            for _, q in e1m[e1m["Filial"] == filial].iterrows():
                pend.append({"etapa": 1, "data": q["Data"].isoformat(), "valor": r2(q["Valor Bruto"]), "ref": f"aut. {q['Cód. de autorização']} · terminal {q['Número do terminal']}",
                             "status": "MAQUININHA SEM PAR NO PDV", "detalhe": ""})
        if len(e2p):
            for _, q in e2p[e2p["Filial"] == filial].iterrows():
                pend.append({"etapa": 2, "data": q["Data"].isoformat(), "valor": r2(q["Valor"]), "ref": f"{q['Ref']} · {str(q['dt'])[11:16]}",
                             "status": str(q["Status"]), "detalhe": ""})
        if len(e2s):
            for _, q in e2s[e2s["Loja"] == filial].iterrows():
                pend.append({"etapa": 2, "data": q["Data"].isoformat(), "valor": r2(q["Valor"]), "ref": str(q["Descrição"]),
                             "status": "BANCO SEM VENDA", "detalhe": ""})
        lojas.append({"nome": filial, "marca": marca(filial), "ec": ec_de.get(filial), "tem_maq": tem_maq, "tem_banco": tem_banco,
                      "dias": dias, "pendencias": sorted(pend, key=lambda q: (q["etapa"], q["data"]))})

    dados = {"gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "de": de.isoformat(), "ate": ate.isoformat(),
             "extrato_ate": extrato_ate.isoformat() if extrato_ate else None, "lojas": lojas}
    tpl = (RAIZ / "painel/template.html").read_text()
    out = RAIZ / "painel/index.html"
    out.write_text(tpl.replace("/*DADOS*/", json.dumps(dados, ensure_ascii=False)))
    n_ok = sum(l["tem_maq"] for l in lojas)
    print(f"{len(lojas)} lojas · {n_ok} com maquininha · {sum(l['tem_banco'] for l in lojas)} com banco · gravado em {out}")


if __name__ == "__main__":
    main()
