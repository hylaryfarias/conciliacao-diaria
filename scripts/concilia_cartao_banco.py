#!/usr/bin/env python3
"""Etapa 2 · cartão: o que a Sicredi devia pagar × o que caiu na conta.

Uso:
    python3 scripts/concilia_cartao_banco.py <extrato.xlsx[,...]> <portal1.xlsx> [portal2.xlsx ...] --de AAAA-MM-DD --ate AAAA-MM-DD [--aluguel "entrada/aluguel/*.xlsx"]

A chave é (loja, data de pagamento, produto, bandeira). O portal Sicredi informa, venda a venda, o
líquido (bruto − MDR) e a data efetiva do pagamento; o extrato traz o crédito já separado nesse eixo:
`SICREDI DEBITO VISA |0001-05`, `SICREDI ANTEC MASTER |...`, `SICREDI CREDITO ELO |...`.

Regras medidas no dado (skill ragga-conciliacao e BIGGS 04, 01–18/10):
  débito : D+1 útil, recebido = líquido, centavo a centavo.
  crédito: D+1 útil pela antecipação automática. recebido = líquido × (1 − taxa do dia). A taxa muda
           todo dia e sai do próprio extrato: (líquido − creditado) ÷ líquido, por bandeira. Se as
           bandeiras concordam entre si, a taxa está certa. Parte da ELO cai fora da antecipação, como
           `CREDITO ELO`, pelo líquido exato.

Status por lote (loja × data de pagamento × produto):
  CONCILIADO        as duas pontas existem e fecham (débito: |Δ| ≤ R$ 0,05; crédito: bandeiras concordam
                    na taxa e |Δ| ≤ R$ 0,05 depois de aplicar a taxa do dia).
  DIVERGENTE        as duas pontas existem e não fecham; o detalhe diz por bandeira.
  AGUARDA EXTRATO   data de pagamento depois da última data lida do extrato.
  SEM CRÉDITO       o portal diz que pagou e o extrato não tem nada nesse dia.
  VENDA ANTES DO PERÍODO  crédito no extrato de vendas anteriores ao Histórico lido (ex.: 01/10 paga 30/09).
"""
import glob
import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from concilia_pix import carrega_lojas  # noqa: E402
from concilia_pix_banco import _extrato_bruto, chave_loja  # noqa: E402
from portal_sicredi import ler_portais  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
OK_STATUS = {"APROVADA", "AUTORIZADA"}
BAND = {"MDS": "MASTER", "Mastercard": "MASTER", "Maestro": "MASTER", "Visa": "VISA", "Visa Electron": "VISA",
        "ELO Full": "ELO", "Elo": "ELO", "ELO": "ELO", "Amex Full": "AMEX", "Amex": "AMEX", "Hipercard": "HIPER", "Hiper": "HIPER"}
PROD = {"Débito": "debito", "Crédito a Vista": "credito", "Crédito Parcelado": "credito", "Crédito": "credito"}
TOL_CENT = 5          # até R$ 0,05 é arredondamento de lote (sobe com o tamanho: ver tol())
TOL_TAXA = 0.0015     # bandeiras concordam na taxa de antecipação até 0,15 ponto percentual


def tol(cent: int) -> int:
    """Arredondamento venda a venda: R$ 0,05 ou 0,02% do lote, o que for maior (lote de R$ 20 mil → R$ 4,00)."""
    return max(TOL_CENT, int(abs(cent) * 0.0002))


FERIADOS = {pd.Timestamp("2026-10-12").date(), pd.Timestamp("2026-11-02").date(), pd.Timestamp("2026-11-15").date(), pd.Timestamp("2026-11-20").date()}


def proximo_util(d):
    d = d + pd.Timedelta(days=1)
    while d.weekday() >= 5 or d in FERIADOS:
        d = d + pd.Timedelta(days=1)
    return d


def vendas_cartao(df: pd.DataFrame) -> pd.DataFrame:
    v = df[df["Status"].astype(str).str.upper().isin(OK_STATUS) & df["Produto"].isin(PROD)].copy()
    v["prod"] = v["Produto"].map(PROD)
    v["band"] = v["Bandeira"].map(lambda b: BAND.get(str(b), str(b).upper().split()[0]))
    v["pago_em"] = pd.to_datetime(v["Data efetiva do pagamento da venda"], dayfirst=True, errors="coerce").dt.date
    v["previsto_em"] = pd.to_datetime(v["Data prevista pagamento da venda"], dayfirst=True, errors="coerce").dt.date
    # sem "Data efetiva" (export feito antes do pagamento): D+1 útil da data calendário da venda, que é a regra
    # medida no dado para débito e para crédito com antecipação automática. A "prevista" do portal é o D+30 sem antecipação.
    regra = pd.to_datetime(v["dt"]).dt.normalize().map(lambda t: proximo_util(t.date()))
    v["pagto"] = v["pago_em"].where(v["pago_em"].notna(), regra)
    for c in ("Valor Bruto", "Valor da taxa (MDR)", "Valor líquido"):
        v[c] = pd.to_numeric(v[c], errors="coerce").fillna(0)
    v["liq_cent"] = (v["Valor líquido"] * 100).round().astype(int)
    v["mdr_cent"] = (v["Valor da taxa (MDR)"] * 100).round().astype(int)
    v["bruto_cent"] = (v["Valor Bruto"] * 100).round().astype(int)
    return v


def extrato_cartao(caminho: str, filiais) -> pd.DataFrame:
    ex = _extrato_bruto(caminho)
    ex = ex[ex["Descrição"].astype(str).str.upper().str.startswith("SICREDI ")].copy()
    m = ex["Descrição"].str.extract(r"SICREDI (DEBITO|ANTECIPACAO|ANTEC|CREDITO) (\w+)")
    ex["tipo"] = m[0].replace({"ANTECIPACAO": "ANTEC"})
    ex["band"] = m[1]
    ex = ex[ex["tipo"].notna() & (ex["band"] != "VOUCHER")]
    ex["prod"] = ex["tipo"].map(lambda t: "debito" if t == "DEBITO" else "credito")
    ex["Valor"] = pd.to_numeric(ex["Valor"], errors="coerce").fillna(0)
    ex["cent"] = (ex["Valor"] * 100).round().astype(int)
    ex["Data"] = pd.to_datetime(ex["Data"]).dt.date
    mapa = {chave_loja(f): f for f in filiais}
    ex["Loja"] = ex["Loja"].map(lambda x: mapa.get(chave_loja(x), str(x)))
    return ex[ex["Loja"].isin(filiais)].reset_index(drop=True)


def ler_aluguel(caminhos, lojas: dict) -> dict:
    """Histórico de aluguel do portal Sicredi (um arquivo por EC): a Sicredi desconta o aluguel das máquinas da
    liquidação do dia 'Data efetiva da cobrança' (medido: BIGGS 01, 02/10, R$ 483,00 a menos no débito Master).
    Devolve {(filial, data): centavos}."""
    from python_calamine import CalamineWorkbook
    out = {}
    for arq in caminhos:
        wb = CalamineWorkbook.from_path(arq)
        rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
        try:
            h = next(i for i, r in enumerate(rows) if "Valor do aluguel" in r)
        except StopIteration:
            continue
        t = pd.DataFrame(rows[h + 1:], columns=rows[h])
        t = t[t["Status"].astype(str).str.upper().str.startswith("DEBIT")]
        for _, r in t.iterrows():
            filial = lojas.get(str(r["Estabelecimento"]).strip())
            if not filial:
                continue
            d = pd.to_datetime(r["Data efetiva da cobrança"], dayfirst=True).date()
            out[(filial, d)] = out.get((filial, d), 0) + int(round(abs(float(r["Valor do aluguel"])) * 100))
    return out


def lote(filial: str, d, prod: str, v: pd.DataFrame, x: pd.DataFrame, extrato_ate, aluguel_cent: int = 0) -> dict:
    """Um lote = loja × data de pagamento × produto. v: vendas pagas nesse dia; x: linhas do extrato desse dia."""
    liq = int(v["liq_cent"].sum()); bruto = int(v["bruto_cent"].sum()); mdr = int(v["mdr_cent"].sum())
    antec = int(x[x["tipo"] == "ANTEC"]["cent"].sum()); normal = int(x[x["tipo"] != "ANTEC"]["cent"].sum())
    recebido = antec + normal
    sem_efet = int(v[v["pago_em"].isna()]["liq_cent"].sum()) if len(v) else 0
    por_dia = json.dumps({str(k): int(x) for k, x in v.groupby("Data")["liq_cent"].sum().items()}) if len(v) else "{}"   # líquido por dia de venda (dia de movimento), para o recorte do painel
    out = {"Filial": filial, "Pagamento": d, "Produto": "Crédito" if prod == "credito" else "Débito", "Vendas (qtd)": len(v),
           "Vendas de": ", ".join(sorted({str(s) for s in v["Data"]})) if len(v) else "", "Por dia": por_dia, "Sem data efetiva (R$)": sem_efet / 100, "Vendas sem data efetiva (qtd)": int(v["pago_em"].isna().sum()) if len(v) else 0,
           "Bruto (R$)": bruto / 100, "MDR (R$)": mdr / 100, "Líquido previsto (R$)": liq / 100,
           "Antecipação (R$)": 0.0, "Taxa antecipação (%)": None, "Esperado na conta (R$)": liq / 100,
           "Aluguel (R$)": aluguel_cent / 100, "Recebido (R$)": recebido / 100, "Δ recebido − esperado (R$)": 0.0, "Status": "", "Detalhe": ""}
    if not len(v):
        out.update({"Status": "VENDA ANTES DO PERÍODO" if d <= extrato_ate else "AGUARDA EXTRATO", "Esperado na conta (R$)": None,
                    "Detalhe": "crédito no extrato sem venda no Histórico lido (vendas anteriores ao período)"})
        return out
    if extrato_ate is None or d > extrato_ate:
        out.update({"Status": "AGUARDA EXTRATO", "Recebido (R$)": None, "Detalhe": f"pagamento previsto para {d:%d/%m}, extrato lido até {extrato_ate:%d/%m}" if extrato_ate else "sem extrato"})
        return out
    if not len(x):
        out.update({"Status": "SEM CRÉDITO", "Δ recebido − esperado (R$)": -liq / 100, "Detalhe": "o portal diz que pagou e não há crédito Sicredi nesse dia no extrato"})
        return out
    por_band = []
    if prod == "debito":
        esperado = liq
        for b, g in v.groupby("band"):
            e_b = int(x[x["band"] == b]["cent"].sum()); l_b = int(g["liq_cent"].sum())
            if abs(e_b - l_b) > tol(l_b):
                por_band.append(f"{b}: previsto R$ {l_b/100:.2f}, recebido R$ {e_b/100:.2f} (Δ {(e_b-l_b)/100:+.2f})")
        for b in set(x["band"]) - set(v["band"]):
            por_band.append(f"{b}: recebido R$ {int(x[x['band']==b]['cent'].sum())/100:.2f} sem venda prevista")
    else:
        # crédito: parte não antecipada (CREDITO <band>) entra pelo líquido exato; o resto é antecipado com a taxa do dia
        taxas = {}
        for b, g in v.groupby("band"):
            l_b = int(g["liq_cent"].sum()); cr_b = int(x[(x["band"] == b) & (x["tipo"] != "ANTEC")]["cent"].sum()); an_b = int(x[(x["band"] == b) & (x["tipo"] == "ANTEC")]["cent"].sum())
            base = l_b - cr_b
            if base > 0 and an_b > 0:
                taxas[b] = (base - an_b) / base
        taxa = sorted(taxas.values())[len(taxas) // 2] if taxas else 0.0
        if taxas and max(taxas.values()) - min(taxas.values()) > TOL_TAXA:
            por_band.append("taxa de antecipação não concorda entre bandeiras: " + ", ".join(f"{b} {t*100:.4f}%" for b, t in sorted(taxas.items())))
        esperado = 0
        for b, g in v.groupby("band"):
            l_b = int(g["liq_cent"].sum()); cr_b = int(x[(x["band"] == b) & (x["tipo"] != "ANTEC")]["cent"].sum()); an_b = int(x[(x["band"] == b) & (x["tipo"] == "ANTEC")]["cent"].sum())
            base = max(0, l_b - cr_b)
            esp_b = cr_b + int(round(base * (1 - taxa)))
            esperado += esp_b
            if abs((cr_b + an_b) - esp_b) > tol(l_b):
                por_band.append(f"{b}: líquido R$ {l_b/100:.2f}, esperado R$ {esp_b/100:.2f}, recebido R$ {(cr_b+an_b)/100:.2f} (Δ {((cr_b+an_b)-esp_b)/100:+.2f})")
        for b in set(x["band"]) - set(v["band"]):
            por_band.append(f"{b}: recebido R$ {int(x[x['band']==b]['cent'].sum())/100:.2f} sem venda prevista")
        out["Antecipação (R$)"] = (liq - esperado) / 100
        out["Taxa antecipação (%)"] = round(taxa * 100, 4)
    out["Esperado na conta (R$)"] = esperado / 100
    out["Δ recebido − esperado (R$)"] = (recebido - esperado) / 100
    fecha = abs(recebido - esperado) <= tol(esperado)
    if fecha:
        out["Status"] = "CONCILIADO"
    elif sem_efet > liq / 2 or (sem_efet > 0 and abs((recebido - esperado) + sem_efet) <= tol(esperado) + int(sem_efet * 0.03)):
        # o Histórico foi exportado antes de a Sicredi registrar o pagamento: a data é a da regra (D+1 útil), e a parte
        # da ELO que cai fora da antecipação (D+2) muda de lote. Não é divergência: é reexportar o Histórico.
        out["Status"] = "AGUARDA PORTAL"
        por_band.insert(0, "vendas ainda sem data efetiva de pagamento no Histórico exportado; reexportar o Histórico para confirmar")
    else:
        out["Status"] = "DIVERGENTE"
    out["Detalhe"] = "; ".join(por_band)
    return out


def aplica_aluguel(lotes_dia: list, aluguel_cent: int) -> None:
    """A Sicredi desconta o aluguel do dia da liquidação, repartido entre débito e crédito como couber (medido:
    BIGGS 01 inteiro no débito; BIGGS 10 e 13 inteiros no crédito; BIGGS 11 R$ 912,64 no débito e R$ 164,36 no
    crédito). Se as faltas dos lotes do dia somam o aluguel, cada lote recebe a sua parte e fecha."""
    if not aluguel_cent or not lotes_dia:
        return
    pagos = [l for l in lotes_dia if l["Recebido (R$)"] is not None and l["Esperado na conta (R$)"] is not None and l["Status"] not in ("AGUARDA EXTRATO", "VENDA ANTES DO PERÍODO")]
    if not pagos:
        return
    falta = {id(l): int(round(-l["Δ recebido − esperado (R$)"] * 100)) for l in pagos}
    total = sum(max(0, f) for f in falta.values())
    folga = max(TOL_CENT * len(pagos), int(aluguel_cent * 0.001))
    if abs(total - aluguel_cent) <= folga and total > 0:
        resto = aluguel_cent
        for l in sorted(pagos, key=lambda l: -falta[id(l)]):
            parte = min(max(0, falta[id(l)]), resto); resto -= parte
            if parte <= 0:
                continue
            esperado = int(round(l["Esperado na conta (R$)"] * 100)) - parte
            recebido = int(round(l["Recebido (R$)"] * 100))
            l["Aluguel (R$)"] = parte / 100; l["Esperado na conta (R$)"] = esperado / 100; l["Δ recebido − esperado (R$)"] = (recebido - esperado) / 100
            l["Status"] = "CONCILIADO" if abs(recebido - esperado) <= tol(esperado) else l["Status"]
            l["Detalhe"] = f"aluguel das máquinas de R$ {parte/100:.2f} descontado desta liquidação" + (f" (de R$ {aluguel_cent/100:.2f} no dia)" if parte != aluguel_cent else "")
    else:
        # as faltas do dia não somam o aluguel: cobre só o que falta (até o valor do aluguel) e registra o resto,
        # que a Sicredi deve ter descontado de outra liquidação (ex.: a de 01/10, que paga 30/09, fora do período)
        resto = aluguel_cent
        for l in sorted(pagos, key=lambda l: -falta[id(l)]):
            parte = min(max(0, falta[id(l)]), resto)
            if parte <= 0:
                continue
            esperado = int(round(l["Esperado na conta (R$)"] * 100)) - parte; recebido = int(round(l["Recebido (R$)"] * 100))
            if abs(recebido - esperado) > tol(esperado):
                continue   # a falta não é o aluguel: deixa o lote como está
            resto -= parte
            l["Aluguel (R$)"] = parte / 100; l["Esperado na conta (R$)"] = esperado / 100; l["Δ recebido − esperado (R$)"] = (recebido - esperado) / 100
            l["Status"] = "CONCILIADO"; l["Detalhe"] = f"aluguel das máquinas de R$ {parte/100:.2f} descontado desta liquidação (de R$ {aluguel_cent/100:.2f} cobrados no dia)"
        if resto > 0:
            nota = f"R$ {resto/100:.2f} do aluguel de R$ {aluguel_cent/100:.2f} cobrado neste dia não aparecem descontados dos lotes do dia; provável desconto em outra liquidação (ex.: 01/10, que paga vendas de 30/09)"
            for l in pagos:
                l["Detalhe"] = (str(l["Detalhe"]) + "; " + nota).strip("; ") if l["Status"] != "CONCILIADO" or not l["Detalhe"] else l["Detalhe"]


def main() -> None:
    args = sys.argv[1:]
    if len(args) < 2 or "--de" not in args or "--ate" not in args:
        sys.exit(__doc__)
    de = pd.Timestamp(args[args.index("--de") + 1]).date(); ate = pd.Timestamp(args[args.index("--ate") + 1]).date()
    for k in ("--de", "--ate"):
        i = args.index(k); del args[i:i + 2]
    aluguel_arqs = []
    if "--aluguel" in args:
        i = args.index("--aluguel"); aluguel_arqs = sorted(glob.glob(args[i + 1])); del args[i:i + 2]
    extrato_arq, portais = args[0], args[1:]
    lojas = carrega_lojas()
    aluguel = ler_aluguel(aluguel_arqs, lojas)
    filiais = list(lojas.values())
    ex = extrato_cartao(extrato_arq, filiais)
    extrato_ate = ex["Data"].max() if len(ex) else None

    lotes, por_dia_venda = [], []
    for (filial, ec), df in ler_portais(portais, lojas).items():
        v = vendas_cartao(df)
        # o ciclo de pagamento segue a data CALENDÁRIO da venda (a madrugada de 01/10 paga em 02/10 junto com o dia 01),
        # por isso o recorte aqui é por calendário; o dia de movimento (Data) fica só para ligar ao painel
        v["cal"] = pd.to_datetime(v["dt"]).dt.date
        v = v[(v["cal"] >= de) & (v["cal"] <= ate)]
        xl = ex[ex["Loja"] == filial]
        datas = sorted(set(v["pagto"].dropna()) | set(xl["Data"]))
        for d in datas:
            if d < de:
                continue
            do_dia = []
            for prod in ("debito", "credito"):
                vv = v[(v["pagto"] == d) & (v["prod"] == prod)]
                xx = xl[(xl["Data"] == d) & (xl["prod"] == prod)]
                if not len(vv) and not len(xx):
                    continue
                do_dia.append(lote(filial, d, prod, vv, xx, extrato_ate))
            aplica_aluguel(do_dia, aluguel.get((filial, d), 0))
            lotes.extend(do_dia)
        # por dia de venda: para onde foi o líquido de cada dia e se chegou
        st = {(l["Pagamento"], l["Produto"]): l["Status"] for l in lotes if l["Filial"] == filial}
        for (d, prod), g in v.groupby(["Data", "prod"]):
            pagos = sorted(set(g["pagto"].dropna()))
            status = [st.get((p, "Crédito" if prod == "credito" else "Débito"), "AGUARDA EXTRATO") for p in pagos]
            por_dia_venda.append({"Filial": filial, "Venda": d, "Produto": "Crédito" if prod == "credito" else "Débito", "Líquido (R$)": int(g["liq_cent"].sum()) / 100,
                                  "Pago em": ", ".join(p.strftime("%d/%m") for p in pagos),
                                  "Status": "CONCILIADO" if status and all(s == "CONCILIADO" for s in status) else "DIVERGENTE" if any(s in ("DIVERGENTE", "SEM CRÉDITO") for s in status) else "AGUARDA PORTAL" if any(s == "AGUARDA PORTAL" for s in status) else "AGUARDA EXTRATO"})

    L = pd.DataFrame(lotes).sort_values(["Filial", "Pagamento", "Produto"])
    V = pd.DataFrame(por_dia_venda).sort_values(["Filial", "Venda", "Produto"])
    res = (L[L["Status"] != "VENDA ANTES DO PERÍODO"].groupby(["Filial", "Produto"])
           .agg(**{"Líquido previsto (R$)": ("Líquido previsto (R$)", "sum"), "Antecipação (R$)": ("Antecipação (R$)", "sum"), "Esperado na conta (R$)": ("Esperado na conta (R$)", "sum"),
                   "Recebido (R$)": ("Recebido (R$)", "sum"), "Δ (R$)": ("Δ recebido − esperado (R$)", "sum"),
                   "Lotes": ("Status", "size"), "Conciliados": ("Status", lambda s: int((s == "CONCILIADO").sum())),
                   "Divergentes": ("Status", lambda s: int(s.isin(["DIVERGENTE", "SEM CRÉDITO"]).sum())), "Aguardam extrato": ("Status", lambda s: int((s == "AGUARDA EXTRATO").sum())), "Aguardam portal": ("Status", lambda s: int((s == "AGUARDA PORTAL").sum()))})
           .round(2).reset_index())
    out = RAIZ / "saida" / f"concilia_cartao_banco_{de}_a_{ate}.xlsx"
    with pd.ExcelWriter(out) as w:
        res.to_excel(w, sheet_name="Resumo loja", index=False)
        L.to_excel(w, sheet_name="Recebimentos", index=False)
        V.to_excel(w, sheet_name="Por dia de venda", index=False)
        ex.sort_values(["Loja", "Data"]).to_excel(w, sheet_name="Extrato cartão", index=False)
    n = L["Status"].value_counts().to_dict()
    print(f"{len(L)} lotes · {n} · extrato lido até {extrato_ate} · gravado em {out}")


if __name__ == "__main__":
    main()
