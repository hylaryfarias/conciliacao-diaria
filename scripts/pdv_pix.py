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

HORA_VIRADA = 10  # a venda vai da abertura do caixa até a reabertura da loja (10h): tudo entre meia-noite e 09:59 ainda é do dia anterior


def dia_movimento(dt: pd.Series) -> pd.Series:
    """Dia de movimento de um datetime calendário (Sicredi, extrato): antes das 06:00 pertence ao dia anterior."""
    return (pd.to_datetime(dt) - pd.Timedelta(hours=HORA_VIRADA)).dt.date

FORMAS_PIX = {"PIX MAQUININHA", "PIX - TEF", "TEF - PIX"}


def sem_cancelados(df: pd.DataFrame) -> pd.DataFrame:
    """Cupons cancelados no PDV continuam no relatório de cupons, mas não são venda. O relatório analítico só de
    cancelados (entrada/pdv_cancelados_*.xlsx, exportado no Cloudfy com o filtro de cancelados) diz quais são:
    saem da conciliação por (Filial, Data, Cupom)."""
    import glob
    arqs = sorted(glob.glob(str(Path(__file__).resolve().parent.parent / "entrada" / "pdv_cancelados_*.xlsx")))
    if not arqs or not len(df):
        return df
    canc = set()
    for a in arqs:
        wb = CalamineWorkbook.from_path(a)
        rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
        idx = next((i for i, r in enumerate(rows) if "Cupom" in r and "Filial" in r), None)
        if idx is None:
            continue
        h = rows[idx]; fi, di, ci = h.index("Filial"), h.index("Data"), h.index("Cupom")
        for r in rows[idx + 1:]:
            if r and str(r[fi]).strip():
                canc.add((str(r[fi]).strip(), pd.to_datetime(str(r[di])[:10], dayfirst=True).date(), str(r[ci]).replace(".0", "").strip()))
    if not canc:
        return df
    k = list(zip(df["Filial"], pd.to_datetime(df["Data"], dayfirst=True).dt.date, df["Cupom"].astype(str).str.replace(".0", "", regex=False).str.strip()))
    mask = pd.Series([x in canc for x in k], index=df.index)
    return df[~mask]


def ler_pdv(caminho: str) -> pd.DataFrame:
    """Relatório de cupons de vendas. Aceita vários arquivos (separados por vírgula ou curinga), ex.: o
    relatório 01–14/10 mais o diário de 15/10; cupons repetidos entre arquivos entram uma vez só."""
    import glob
    arqs = [a for parte in str(caminho).split(",") for a in (sorted(glob.glob(parte.strip())) or [parte.strip()])]
    partes = []
    for a in arqs:
        wb = CalamineWorkbook.from_path(a)
        rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
        # cabeçalho achado pelo conteúdo, nunca por índice fixo
        idx = next(i for i, r in enumerate(rows) if "Desc. pagam." in r)
        partes.append(pd.DataFrame(rows[idx + 1 :], columns=rows[idx]))
    # repetição ENTRE arquivos sai; repetição DENTRO do mesmo arquivo fica (conta dividida em partes iguais tem
    # duas linhas idênticas no mesmo cupom, ex.: ROBS 60212, dois PIX de 32,45)
    for i, p in enumerate(partes):
        p["_arq"] = i
    df = pd.concat(partes, ignore_index=True)
    df["Filial"] = df["Filial"].astype(str).str.strip()   # o relatório às vezes vem com espaço no fim ("BIGGS 04 - WINSTON ")
    df = df[df["Filial"] != ""]
    chave = [c for c in ("Filial", "Caixa", "Data", "Hora", "Cupom", "Desc. pagam.", "Vl. pagamento") if c in df.columns]
    df["_seq"] = df.groupby(["_arq"] + chave).cumcount()
    df = df.drop_duplicates(subset=chave + ["_seq"]).drop(columns=["_arq", "_seq"])
    df = sem_cancelados(df)
    df["Vl. pagamento"] = pd.to_numeric(df["Vl. pagamento"], errors="coerce").fillna(0)
    df["cent"] = (df["Vl. pagamento"] * 100).round().astype(int)
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True).dt.date
    df["Hora"] = df["Hora"].astype(str).str[:5]
    return df


VOUCHER_MOD = {"Alimentação", "Refeição", "Voucher", "Multiplo", "Flex", "Multibeneficios", "Gift"}


def explica_sem_par(sobra: pd.DataFrame, maq_tudo: pd.DataFrame, pdv_loja: pd.DataFrame, cupons_casados) -> pd.DataFrame:
    """Para baixas do PDV sem venda na máquina, tenta duas explicações medidas no dado:
    1. voucher (Ticket, Alelo, Pluxee, Fepas, Elo/Visa voucher) passado na máquina no mesmo minuto e baixado como cartão/PIX;
    2. cupom vizinho idêntico (mesmo caixa, mesmo valor, até 5 min) que já casou com a venda da máquina: fica como
       baixa sem par (o que o operador baixou vale, e é falta de caixa), só com a observação em "Cupom igual".
    sobra precisa de dt_pdv, Data, cent, Caixa, Cupom; maq_tudo de dt, Data, cent, Modalidade, Bandeira, Status."""
    sobra = sobra.copy()
    if "Status" not in sobra.columns:
        sobra["Status"] = "SEM PAR NA MAQUININHA"
    for c in ("Modalidade maq", "Bandeira maq", "Hora maq", "Autorização", "Terminal", "Cupom igual"):
        if c not in sobra.columns:
            sobra[c] = None
    sobra["Status"] = sobra["Status"].astype(object); sobra["Cupom igual"] = sobra["Cupom igual"].astype(object)
    vou = maq_tudo[maq_tudo["Modalidade"].isin(VOUCHER_MOD) & maq_tudo["Status"].str.upper().isin({"APROVADA", "AUTORIZADA"})]
    usados: set = set()
    casados = {str(c) for c in cupons_casados}
    for i, a in sobra.iterrows():
        if str(a["Status"]) != "SEM PAR NA MAQUININHA":
            continue
        if len(vou):
            dif = (vou["dt"] - a["dt_pdv"]).abs().dt.total_seconds() / 60
            c = vou[(vou["cent"] == a["cent"]) & (vou["Data"] == a["Data"]) & (dif <= 30) & (~vou.index.isin(usados))]
            if len(c):
                j = dif[c.index].idxmin(); usados.add(j); m = vou.loc[j]
                sobra.loc[i, ["Status", "Modalidade maq", "Bandeira maq", "Hora maq", "Autorização", "Terminal"]] = [
                    f"PROVÁVEL FORMA ERRADA: VOUCHER {m['Bandeira']} na máquina às {m['dt']:%H:%M}", m["Modalidade"], m["Bandeira"], m["dt"].strftime("%H:%M:%S"), m["Cód. de autorização"], m["Número do terminal"]]
                continue
        d = pdv_loja[(pdv_loja["Caixa"] == a["Caixa"]) & (pdv_loja["cent"] == a["cent"]) & (pdv_loja["Cupom"].astype(str) != str(a["Cupom"]))
                     & ((pdv_loja["dt_pdv"] - a["dt_pdv"]).abs() <= pd.Timedelta(minutes=5))]
        d = d[d["Cupom"].astype(str).isin(casados)]
        if len(d):
            sobra.loc[i, "Cupom igual"] = str(d.iloc[0]["Cupom"])
    return sobra


def ler_analitico_mesa(caminhos=None) -> pd.DataFrame:
    """Relatório analítico do PDV (itens por cupom): quanto de cada cupom veio de mesa/comanda.
    A taxa de serviço (10%) incide só sobre os itens da mesa; o que é lançado direto no caixa
    (bombom, bebida avulsa) entra sem taxa. Sem caminhos, usa entrada/pdv_analitico_*.xlsx.
    Retorna uma linha por (Filial, Caixa, Cupom) com cent_mesa (centavos dos itens com mesa)."""
    import glob
    raiz = Path(__file__).resolve().parent.parent
    arqs = caminhos or sorted(glob.glob(str(raiz / "entrada" / "pdv_analitico_*.xlsx")))
    partes = []
    for a in arqs:
        wb = CalamineWorkbook.from_path(a)
        rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
        try:
            idx = next(i for i, r in enumerate(rows) if "Nr. mesa/comanda" in r and "Cupom" in r)
        except StopIteration:
            continue
        cab, vistos = [], {}
        for c in rows[idx]:   # "Total" aparece duas vezes (cupom e item): a segunda vira "Total.1", como o pandas faz
            n = vistos.get(c, 0); vistos[c] = n + 1
            cab.append(c if n == 0 else f"{c}.{n}")
        t = pd.DataFrame(rows[idx + 1 :], columns=cab)
        partes.append(t[["Filial", "Caixa", "Cupom", "Item", "Total.1", "Nr. mesa/comanda"]])
    if not partes:
        return pd.DataFrame(columns=["Filial", "Caixa", "Cupom", "cent_mesa"])
    t = pd.concat(partes, ignore_index=True)
    t["Filial"] = t["Filial"].astype(str).str.strip()
    t = t[t["Filial"] != ""]
    t["Cupom"] = pd.to_numeric(t["Cupom"], errors="coerce")
    t = t.dropna(subset=["Cupom"]).drop_duplicates(subset=["Filial", "Caixa", "Cupom", "Item"])
    t["Cupom"] = t["Cupom"].astype(int)
    mesa = t["Nr. mesa/comanda"].astype(str).str.strip().replace({"": None, "None": None, "nan": None}).notna()
    t["cent_item"] = (pd.to_numeric(t["Total.1"], errors="coerce").fillna(0) * 100).round().astype(int)
    t["cent_mesa"] = t["cent_item"].where(mesa, 0)
    return t.groupby(["Filial", "Caixa", "Cupom"], as_index=False)["cent_mesa"].sum()


def com_base_mesa(bruto: pd.DataFrame, analitico: pd.DataFrame | None = None) -> pd.DataFrame:
    """Acrescenta cent_mesa às linhas do relatório de cupons (0 quando o cupom não tem itens de mesa ou o
    analítico não cobre o cupom)."""
    a = ler_analitico_mesa() if analitico is None else analitico
    if not len(a):
        bruto["cent_mesa"] = 0
        return bruto
    b = bruto.copy()
    b["Cupom"] = pd.to_numeric(b["Cupom"], errors="coerce")
    b = b.merge(a, on=["Filial", "Caixa", "Cupom"], how="left")
    b["cent_mesa"] = b["cent_mesa"].fillna(0).astype(int)
    b["Cupom"] = bruto["Cupom"].values
    b.index = bruto.index
    return b


def pix_do_pdv(df: pd.DataFrame) -> pd.DataFrame:
    pix = df[df["Desc. pagam."].str.upper().isin(FORMAS_PIX) & (df["cent"] > 0)].copy()   # cupom de R$ 0,00 não é pendência
    cols = ["Filial", "Caixa", "Data", "Hora", "Cupom", "Desc. pagam.", "Vl. pagamento", "cent", "Chave"]
    cols += [c for c in ("dt_pdv", "Dia movimento", "cent_mesa") if c in pix.columns]
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
