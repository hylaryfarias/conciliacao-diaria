#!/usr/bin/env python3
"""Gera painel/index.html a partir das planilhas de saida/ e do relatório do PDV.

Uso:
    python3 scripts/gera_painel.py <pdv_cupons.xlsx> --de AAAA-MM-DD --ate AAAA-MM-DD

Lê saida/concilia_pix_*.xlsx e concilia_pix_banco_*.xlsx (PIX, etapas 1 e 2),
concilia_credito_*.xlsx e concilia_debito_*.xlsx (cartão, etapa 1) e injeta um JSON em
painel/template.html no marcador /*DADOS*/.
"""
import glob
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from concilia_pix import HORA_VIRADA, carrega_lojas  # noqa: E402
from pdv_pix import ler_pdv  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
DOW = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]
FERIADOS = {date(2026, 10, 12)}
MARCA = {"ROBS": "robs", "MAPOLI": "mapoli"}
FORMAS = {
    "pix": {"PIX MAQUININHA", "PIX - TEF", "TEF - PIX"},
    "credito": {"TEF - CREDITO", "CARTAO CREDITO"},
    "debito": {"TEF - DEBITO", "CARTAO DEBITO"},
}


def marca(filial: str) -> str:
    for k, v in MARCA.items():
        if filial.startswith(k):
            return v
    return filial.split()[0].lower()


def r2(x) -> float:
    return round(float(x), 2) if pd.notna(x) else 0.0


def le(padrao: str, aba: str, col_data="Data") -> pd.DataFrame:
    fs = glob.glob(str(RAIZ / "saida" / padrao))
    if not fs:
        return pd.DataFrame()
    df = pd.concat([pd.read_excel(f, sheet_name=aba) for f in fs], ignore_index=True)
    if col_data in df.columns:
        df[col_data] = pd.to_datetime(df[col_data]).dt.date
    return df


def main() -> None:
    args = sys.argv[1:]
    de = pd.Timestamp(args[args.index("--de") + 1]).date()
    ate = pd.Timestamp(args[args.index("--ate") + 1]).date()

    b = ler_pdv(args[0])
    b["dt_pdv"] = pd.to_datetime(b["Data"].astype(str) + " " + b["Hora"].astype(str).str[:5])
    b.loc[b["dt_pdv"].dt.hour < HORA_VIRADA, "dt_pdv"] += pd.Timedelta(days=1)
    b["Data"] = b["dt_pdv"].dt.date
    b = b[(b["Data"] >= de) & (b["Data"] <= ate)]
    b["forma"] = b["Desc. pagam."].str.upper()
    pdv_dia = {m: b[b["forma"].isin(f)].groupby(["Filial", "Data"])["cent"].agg(qtd="count", cent="sum") for m, f in FORMAS.items()}

    e1 = le("concilia_pix_2*.xlsx", "Resumo loja x dia"); e1p = le("concilia_pix_2*.xlsx", "PDV sem par"); e1m = le("concilia_pix_2*.xlsx", "Maquininha sem par")
    e2 = le("concilia_pix_banco_*.xlsx", "Resumo loja x dia", "Venda"); e2p = le("concilia_pix_banco_*.xlsx", "Pendências"); e2s = le("concilia_pix_banco_*.xlsx", "Banco sem venda")
    cart = {m: (le(f"concilia_{m}_2*.xlsx", "Resumo loja x dia"), le(f"concilia_{m}_2*.xlsx", "Maquininha sem par"), le(f"concilia_{m}_2*.xlsx", "PDV sem par")) for m in ("credito", "debito")}
    csv_ = {m: le(f"concilia_{m}_2*.xlsx", "Serviço não baixado") for m in ("credito", "debito")}
    l1 = le("concilia_pix_2*.xlsx", "PDV x maquininha")              # linhas PIX, lado PDV
    lb = le("concilia_pix_banco_*.xlsx", "Venda x banco")             # linhas PIX, lado banco
    lc = {m: le(f"concilia_{m}_2*.xlsx", "Maquininha x PDV") for m in ("credito", "debito")}
    # pendências só do período do painel
    no_per = lambda df: df[(df["Data"] >= de) & (df["Data"] <= ate)] if len(df) else df
    e1p, e1m, e2p, e2s = no_per(e1p), no_per(e1m), no_per(e2p), no_per(e2s)
    cart = {m: (r_, no_per(sm_), no_per(sp_)) for m, (r_, sm_, sp_) in cart.items()}
    csv_ = {m: no_per(x) for m, x in csv_.items()}
    banco_por_aut = {}
    if len(lb):
        for _, q in lb[lb["Status"] == "CONCILIADO"].iterrows():
            banco_por_aut[str(q["Ref"]).replace("aut. ", "")] = (str(q["Crédito"]), str(q["Pagador"]))
    ec_de = {v: k for k, v in carrega_lojas().items()}
    MODK = {"PIX": "pix", "Crédito": "credito", "Débito": "debito"}
    cupL = le("concilia_cupom.xlsx", "Linhas")
    cupR = le("concilia_cupom.xlsx", "Cupons resolvidos")
    resolv_pdv, resolv_maq, expl = {}, {}, {}
    if len(cupL):
        cupL = cupL[(cupL["Data"] >= de) & (cupL["Data"] <= ate)]
        for _, q in cupR.iterrows():
            expl[(q["Filial"], str(q["Cupom"]))] = f"comanda {q['Cupom']} fechou {q['Total']:.2f}: PDV {q['FormasPDV']} · máquina {q['FormasMaq']}"
        ferr = {}   # (filial, mod, data, cupom) -> dados da venda na máquina (forma errada 1:1)
        for _, q in cupL.iterrows():
            k = (q["Filial"], MODK[q["Forma"]], q["Data"])
            if q["Tipo"] == "FORMA ERRADA":
                if q["Lado"] == "PDV":
                    ferr[(q["Filial"], MODK[q["Forma"]], q["Data"], str(q["Ref"]).replace(".0", ""))] = q
                continue
            if q["Lado"] == "PDV":
                resolv_pdv.setdefault(k, {})[str(q["Ref"]).replace(".0", "")] = (r2(q["Valor"]), str(q["Cupom"]))
            else:
                resolv_maq.setdefault(k, {})[str(q["Ref"])] = (r2(q["Valor"]), str(q["Cupom"]))

    def aplica_comanda(filial, m, d, row, pend):
        """Marca linhas e pendências resolvidas pelo total da comanda; ajusta as somas do dia."""
        # 1:1 de formas diferentes = baixa errada: vira "provável forma errada" com botão de transferir
        for p in pend:
            q = ferr.get((filial, m, d, str(p.get("cupom", ""))))
            if q is not None and p["status"].startswith("SEM PAR") and p["data"] == d.isoformat():
                dm = pd.to_datetime(q["DataMaq"]).date()
                p["status"] = f"PROVÁVEL FORMA ERRADA: {q['FormaMaq']}{'' if q['BandeiraMaq'] == q['FormaMaq'] else ' ' + str(q['BandeiraMaq'])} na máquina em {dm:%d/%m} {q['HoraMaq']}"
                p["detalhe"] = f"terminal {q['TerminalMaq']} · aut. {q['AutMaq']}"
                p["aut"] = str(q["AutMaq"]); p["para"] = MODK[q["FormaMaq"]]; p["data_maq"] = dm.isoformat()
                for r in row.get("lanc", []):
                    if r.get("c") == p["cupom"] and str(r.get("s", "")).startswith("SEM PAR"):
                        r["s"] = p["status"]; r["hm"] = q["HoraMaq"]; r["t"] = str(q["TerminalMaq"]); r["a"] = str(q["AutMaq"])
        rp = resolv_pdv.get((filial, m, d), {}); rm = resolv_maq.get((filial, m, d), {})
        if not rp and not rm:
            return pend
        for r in row.get("lanc", []):
            if r.get("c") in rp and not str(r.get("s", "")).startswith("CONCILIADO"):
                r["s"] = "CONCILIADO POR TOTAL DA COMANDA"; r["cm"] = expl.get((filial, rp[r["c"]][1]), "")
            if r.get("a") in rm and str(r.get("s", "")).startswith("MAQUININHA SEM PAR"):
                r["s"] = "CONCILIADO POR TOTAL DA COMANDA"; r["c"] = rm[r["a"]][1]; r["cm"] = expl.get((filial, rm[r["a"]][1]), "")
        if "pdv_sem_par" in row:
            row["pdv_sem_par"] = r2(max(0, row["pdv_sem_par"] - sum(v for v, _ in rp.values())))
        if "maq_sem_par" in row:
            row["maq_sem_par"] = r2(max(0, row["maq_sem_par"] - sum(v for v, _ in rm.values())))
        row["comanda"] = r2(sum(v for v, _ in rp.values()))
        novo = []
        for p in pend:
            if p["data"] != d.isoformat():
                novo.append(p); continue
            if p.get("cupom") in rp and p["status"].startswith("SEM PAR"):
                p = {**p, "status": "CONCILIADO POR TOTAL DA COMANDA", "detalhe": expl.get((filial, rp[p["cupom"]][1]), ""), "comanda": True}
            elif p["status"].startswith("MAQUININHA SEM PAR") and any(f"aut. {a}" in str(p["ref"]) for a in rm):
                continue
            novo.append(p)
        return novo

    extrato_ate = None
    if len(e2):
        creds = pd.to_datetime(e2["Crédito em"].dropna().astype(str).str.split(", ").explode()).dt.date
        extrato_ate = max(creds) if len(creds) else None

    def sel(df, filial, d, col="Data"):
        if not len(df):
            return None
        x = df[(df["Filial"] == filial) & (df[col] == d)]
        return x.iloc[0] if len(x) else None

    def hh(x):
        x = str(x)
        return x[11:16] if len(x) > 10 and x[10] == " " else x[:5]

    def lanc_pdv(filial, m, d):
        q = b[(b["Filial"] == filial) & (b["Data"] == d) & (b["forma"].isin(FORMAS[m]))].sort_values("dt_pdv")
        return [{"h": hh(x["dt_pdv"]), "c": str(x["Cupom"]), "cx": str(x["Caixa"]), "f": str(x["Desc. pagam."]), "v": r2(x["cent"] / 100), "s": "PDV"} for _, x in q.iterrows()]

    def lanc_pix(filial, d):
        if not len(l1) or not len(l1[(l1["Filial"] == filial) & (l1["Data"] == d)]):
            return lanc_pdv(filial, "pix", d)
        rows = []
        for _, x in l1[(l1["Filial"] == filial) & (l1["Data"] == d)].sort_values("Hora").iterrows():
            r = {"h": hh(x["Hora"]), "c": str(x["Cupom"]), "cx": str(x["Caixa"]), "f": str(x["Desc. pagam."]), "v": r2(x["Vl. pagamento"]), "s": str(x["Status"])}
            if pd.notna(x.get("Hora maq")):
                r.update(hm=hh(x["Hora maq"]), t=str(x["Terminal"]), a=str(x["Autorização"]))
                if pd.notna(x.get("Dia maq")) and pd.to_datetime(x["Dia maq"]).date() != d:
                    r["dm"] = pd.to_datetime(x["Dia maq"]).date().isoformat()
                if str(x["Status"]) == "CONCILIADO" and str(x["Autorização"]) in banco_por_aut:
                    r["cr"], r["pg"] = banco_por_aut[str(x["Autorização"])]
            rows.append(r)
        if "Dia maq" in l1.columns:
            vm = l1[(l1["Filial"] == filial) & l1["Dia maq"].notna()]
            vm = vm[(pd.to_datetime(vm["Dia maq"]).dt.date == d) & (vm["Data"] != d)]
            for _, x in vm.iterrows():
                rows.append({"h": hh(x["Hora"]), "dp": x["Data"].isoformat(), "c": str(x["Cupom"]), "cx": str(x["Caixa"]), "f": str(x["Desc. pagam."]), "v": r2(x["Vl. pagamento"]),
                             "s": str(x["Status"]), "hm": hh(x["Hora maq"]), "t": str(x["Terminal"]), "a": str(x["Autorização"])})
        if len(e1m):
            for _, x in e1m[(e1m["Filial"] == filial) & (e1m["Data"] == d)].iterrows():
                rows.append({"s": "MAQUININHA SEM PAR NO PDV", "v": r2(x["Valor Bruto"]), "hm": hh(x["dt"]), "t": str(x["Número do terminal"]), "a": str(x["Cód. de autorização"])})
        rows.sort(key=lambda r: r.get("hm") or r.get("h") or "")
        return rows

    def lanc_cartao(filial, m, d):
        L = lc[m]
        if not len(L) or not len(L[(L["Filial"] == filial) & (pd.to_datetime(L["Data"]).dt.date == d)]):
            return lanc_pdv(filial, m, d)
        rows = []
        for _, x in L[(L["Filial"] == filial) & (pd.to_datetime(L["Data"]).dt.date == d)].sort_values("Hora maq").iterrows():
            r = {"s": str(x["Status"]), "v": r2(x["Valor Bruto"]), "hm": hh(x["Hora maq"]), "b": str(x["Bandeira"]), "t": str(x["Terminal"]), "a": str(x["Autorização"]), "mdr": r2(x["MDR"])}
            if str(x["Status"]).startswith(("CONCILIADO", "SERVIÇO")):
                r.update(h=hh(x["Hora PDV"]), c=str(x["Cupom"]), cx=str(x["Caixa"]), f=str(x["Forma PDV"]))
                if pd.notna(x["Dia PDV"]) and pd.to_datetime(x["Dia PDV"]).date() != d:
                    r["dp"] = pd.to_datetime(x["Dia PDV"]).date().isoformat()
            rows.append(r)
        vp = L[(L["Filial"] == filial) & L["Dia PDV"].notna()]
        vp = vp[(pd.to_datetime(vp["Dia PDV"]).dt.date == d) & (pd.to_datetime(vp["Data"]).dt.date != d)]
        for _, x in vp.sort_values("Hora PDV").iterrows():
            rows.append({"s": str(x["Status"]), "v": r2(x["Valor Bruto"]), "hm": hh(x["Hora maq"]), "dm": pd.to_datetime(x["Data"]).date().isoformat(), "b": str(x["Bandeira"]), "t": str(x["Terminal"]), "a": str(x["Autorização"]), "mdr": r2(x["MDR"]),
                         "h": hh(x["Hora PDV"]), "c": str(x["Cupom"]), "cx": str(x["Caixa"]), "f": str(x["Forma PDV"])})
        sp = cart[m][2]
        if len(sp):
            for _, x in sp[(sp["Filial"] == filial) & (sp["Data"] == d)].iterrows():
                rows.append({"s": "SEM PAR NA MAQUININHA", "h": hh(x["Hora"]), "c": str(x["Cupom"]), "cx": str(x["Caixa"]), "f": str(x["Desc. pagam."]), "v": r2(x["Vl. pagamento"])})
        return rows

    def viradas(filial, m):
        """Pares casados em que o dia da máquina difere do dia do PDV (cupom fechado depois da virada)."""
        out = []
        if m == "pix":
            if not len(l1) or "Dia maq" not in l1.columns:
                return out
            q = l1[(l1["Filial"] == filial) & l1["Dia maq"].notna() & l1["Status"].astype(str).str.startswith(("CONCILIADO", "SERVIÇO"))]
            for _, x in q.iterrows():
                dm = pd.to_datetime(x["Dia maq"]).date()
                if dm != x["Data"]:
                    out.append({"dm": dm, "dp": x["Data"], "v": r2(x["Vl. pagamento"]), "cupom": str(x["Cupom"]), "caixa": str(x["Caixa"]), "hp": hh(x["Hora"]), "hm": hh(x["Hora maq"]), "t": str(x["Terminal"]), "forma": str(x["Desc. pagam."])})
        else:
            L = lc[m]
            if not len(L):
                return out
            q = L[(L["Filial"] == filial) & L["Dia PDV"].notna() & L["Status"].astype(str).str.startswith(("CONCILIADO", "SERVIÇO"))]
            for _, x in q.iterrows():
                dm, dp = pd.to_datetime(x["Data"]).date(), pd.to_datetime(x["Dia PDV"]).date()
                if dm != dp:
                    out.append({"dm": dm, "dp": dp, "v": r2(x["Valor Bruto"]), "cupom": str(x["Cupom"]).replace(".0", ""), "caixa": str(x["Caixa"]), "hp": hh(x["Hora PDV"]), "hm": hh(x["Hora maq"]), "t": str(x["Terminal"]), "forma": str(x["Forma PDV"])})
        return out

    def aplica_viradas(filial, m, dias, pend):
        por_dia = {row["data"]: row for row in dias}
        for v in viradas(filial, m):
            rm, rp = por_dia.get(v["dm"].isoformat()), por_dia.get(v["dp"].isoformat())
            if rm is None and rp is None:
                continue
            if rm is not None and "maq" in rm:
                rm["virada_saiu"] = r2(rm.get("virada_saiu", 0) + v["v"])
            if rp is not None:
                rp["virada_entrou"] = r2(rp.get("virada_entrou", 0) + v["v"])
            pend.append({"etapa": 1, "data": v["dm"].isoformat(), "valor": v["v"], "ref": f"cupom {v['cupom']} · {v['caixa']}", "status": "CONCILIADO · VIRADA DE DIA",
                         "detalhe": f"pago na máquina {v['dm']:%d/%m} {v['hm']} (terminal {v['t']}) · fechado no PDV {v['dp']:%d/%m} {v['hp']} como {v['forma']}",
                         "cupom": v["cupom"], "comanda": True, "virada": True})
        return pend

    lojas = []
    for filial in sorted(b["Filial"].unique()):
        mods = {}
        # ---- PIX
        dias, pend = [], []
        for d in pd.date_range(de, ate).date:
            row = {"data": d.isoformat(), "dow": DOW[d.weekday()], "feriado": d in FERIADOS}
            x = pdv_dia["pix"].loc[(filial, d)] if (filial, d) in pdv_dia["pix"].index else None
            row["pdv_qtd"] = int(x["qtd"]) if x is not None else 0; row["pdv"] = r2(x["cent"] / 100) if x is not None else 0.0
            y = sel(e1, filial, d)
            if y is not None:
                row.update(maq_qtd=int(y["PIX maquininha (qtd)"]), maq=r2(y["PIX maquininha (R$)"]), conc=r2(y["Conciliado (R$)"]), servico=r2(y.get("Serviço não baixado (R$)", 0)),
                           pdv_sem_par=r2(y["PDV sem par (R$)"]), maq_sem_par=r2(y["Maquininha sem par (R$)"]), delta=r2(y["Δ PDV − maquininha (R$)"]))
            z = sel(e2, filial, d, "Venda")
            if z is not None:
                row.update(banco=r2(z["Caiu no banco (R$)"]), credito=str(z["Crédito em"]) if pd.notna(z["Crédito em"]) else "",
                           aguarda=r2(z["Aguarda extrato (R$)"]), nao_caiu=r2(z["Não caiu (R$)"]),
                           sem_banco=r2(z.get("PDV sem par no banco (R$)", 0)))
            row["lanc"] = lanc_pix(filial, d)
            dias.append(row)
        if len(e1p):
            for _, q in e1p[e1p["Filial"] == filial].iterrows():
                item = {"etapa": 1, "data": q["Data"].isoformat(), "valor": r2(q["Vl. pagamento"]), "ref": f"cupom {q['Cupom']} · {q['Hora']} · {q['Caixa']}",
                        "status": str(q["Status"]), "detalhe": f"terminal {q['Terminal']} · aut. {q['Autorização']} · comprovante {q['Comprovante']}" if pd.notna(q.get("Terminal")) else "",
                        "cupom": str(q["Cupom"]), "caixa": str(q["Caixa"]), "hora": str(q["Hora"])[:5], "forma": str(q["Desc. pagam."])}
                if str(q["Status"]).startswith("SERVIÇO"):
                    item["servico"] = r2(q.get("Serviço (R$)", 0))
                mm = re.search(r"FORMA ERRADA: (Crédito|Débito) .* em (\d\d)/(\d\d)", str(q["Status"]))
                if mm:
                    item.update(aut=str(q["Autorização"]), para="credito" if mm.group(1) == "Crédito" else "debito",
                                data_maq=f"{q['Data'].year}-{mm.group(3)}-{mm.group(2)}")
                pend.append(item)
        if len(e1m):
            for _, q in e1m[e1m["Filial"] == filial].iterrows():
                pend.append({"etapa": 1, "data": q["Data"].isoformat(), "valor": r2(q["Valor Bruto"]), "ref": f"aut. {q['Cód. de autorização']} · terminal {q['Número do terminal']}", "status": "MAQUININHA SEM PAR NO PDV", "detalhe": ""})
        if len(e2p):
            for _, q in e2p[e2p["Filial"] == filial].iterrows():
                pend.append({"etapa": 2, "data": q["Data"].isoformat(), "valor": r2(q["Valor"]), "ref": f"{q['Ref']} · {str(q['dt'])[11:16]}", "status": str(q["Status"]), "detalhe": ""})
        if len(e2s):
            for _, q in e2s[e2s["Loja"] == filial].iterrows():
                item = {"etapa": 2, "data": q["Data"].isoformat(), "valor": r2(q["Valor"]), "ref": str(q["Descrição"]), "status": str(q.get("Classificação", "BANCO SEM VENDA")), "detalhe": str(q.get("Evidência", "")), "banco": True}
                cp = str(q.get("Cupom PDV", "")).replace(".0", "")
                if cp and cp != "nan":
                    item["cupom_pago"] = cp
                    for pp in pend:
                        if pp.get("cupom") == cp and pp["status"].startswith("SEM PAR"):
                            pp["status"] = "CONCILIADO COM O BANCO · máquina não registrou"; pp["detalhe"] = f"PIX de R$ {r2(q['Valor']):.2f} entrou na conta em {q['Data']:%d/%m} ({str(q['Descrição'])[16:]})"; pp["comanda"] = True
                    for row in dias:
                        for r in row.get("lanc", []):
                            if r.get("c") == cp and str(r.get("s", "")).startswith("SEM PAR"):
                                r["s"] = "CONCILIADO COM O BANCO"; r["cr"] = q["Data"].isoformat(); r["pg"] = str(q["Descrição"])[16:]
                                row["pdv_sem_par"] = r2(max(0, row.get("pdv_sem_par", 0) - r["v"]))
                pend.append(item)
        for row in dias:
            pend = aplica_comanda(filial, "pix", pd.Timestamp(row["data"]).date(), row, pend)
        pend = aplica_viradas(filial, "pix", dias, pend)
        mods["pix"] = {"tem_maq": any("maq" in r for r in dias), "tem_banco": any("banco" in r for r in dias), "dias": dias, "pendencias": sorted(pend, key=lambda q: (q["etapa"], q["data"]))}

        # ---- cartão
        for m, (res, sem_m, sem_p) in cart.items():
            dias, pend = [], []
            for d in pd.date_range(de, ate).date:
                row = {"data": d.isoformat(), "dow": DOW[d.weekday()], "feriado": d in FERIADOS}
                x = pdv_dia[m].loc[(filial, d)] if (filial, d) in pdv_dia[m].index else None
                row["pdv_qtd"] = int(x["qtd"]) if x is not None else 0; row["pdv"] = r2(x["cent"] / 100) if x is not None else 0.0
                y = sel(res, filial, d)
                if y is not None:
                    row.update(maq_qtd=int(y["Maquininha (qtd)"]), maq=r2(y["Maquininha bruto (R$)"]), mdr=r2(y["MDR (R$)"]), liquido=r2(y["Líquido (R$)"]),
                               conc=r2(y["Casado (R$)"]), servico=r2(y.get("Serviço não baixado (R$)", 0)), pdv_sem_par=r2(y["PDV sem par (R$)"]), maq_sem_par=r2(y["Maquininha sem par (R$)"]), delta=r2(y["Δ PDV − maquininha (R$)"]))
                row["lanc"] = lanc_cartao(filial, m, d)
                dias.append(row)
            if len(sem_m):
                for _, q in sem_m[sem_m["Filial"] == filial].iterrows():
                    pend.append({"etapa": 1, "data": q["Data"].isoformat(), "valor": r2(q["Valor Bruto"]), "ref": f"{q['Hora maq']} · {q['Bandeira']} · terminal {q['Terminal']} · aut. {q['Autorização']}", "status": "MAQUININHA SEM PAR NO PDV", "detalhe": ""})
            if len(sem_p):
                for _, q in sem_p[sem_p["Filial"] == filial].iterrows():
                    pend.append({"etapa": 1, "data": q["Data"].isoformat(), "valor": r2(q["Vl. pagamento"]), "ref": f"cupom {q['Cupom']} · {q['Hora']} · {q['Caixa']}", "status": "SEM PAR NA MAQUININHA", "detalhe": "",
                                 "cupom": str(q["Cupom"]), "caixa": str(q["Caixa"]), "hora": str(q["Hora"])[:5], "forma": str(q["Desc. pagam."])})
            sv = csv_[m]
            if len(sv):
                for _, q in sv[sv["Filial"] == filial].iterrows():
                    pend.append({"etapa": 1, "data": q["Data"].isoformat(), "valor": r2(q["Valor Bruto"]), "ref": f"cupom {q['Cupom']} · {q['Hora PDV']} · {q['Caixa']}", "status": str(q["Status"]),
                                 "detalhe": f"{q['Bandeira']} · terminal {q['Terminal']} · aut. {q['Autorização']} · máquina {q['Hora maq']}", "servico": r2(q["Serviço (R$)"])})
            for row in dias:
                pend = aplica_comanda(filial, m, pd.Timestamp(row["data"]).date(), row, pend)
            pend = aplica_viradas(filial, m, dias, pend)
            mods[m] = {"tem_maq": any("maq" in r for r in dias), "tem_banco": False, "dias": dias, "pendencias": sorted(pend, key=lambda q: q["data"])}

        lojas.append({"nome": filial, "marca": marca(filial), "ec": ec_de.get(filial), "mod": mods})

    dados = {"gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "de": de.isoformat(), "ate": ate.isoformat(),
             "extrato_ate": extrato_ate.isoformat() if extrato_ate else None, "lojas": lojas}
    out = RAIZ / "painel/index.html"
    out.write_text((RAIZ / "painel/template.html").read_text().replace("/*DADOS*/", json.dumps(dados, ensure_ascii=False)))
    print(f"{len(lojas)} lojas · PIX com maquininha: {sum(l['mod']['pix']['tem_maq'] for l in lojas)} · com banco: {sum(l['mod']['pix']['tem_banco'] for l in lojas)} · gravado em {out}")


if __name__ == "__main__":
    main()
