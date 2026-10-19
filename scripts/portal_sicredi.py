#!/usr/bin/env python3
"""Leitor do 'Relatório Histórico de vendas' do portal da maquininha Sicredi.

Layout de 22 colunas (versão com 'Modalidade' e 'Hora da venda'), exportado em .xlsx.
Preâmbulo de tamanho variável: o cabeçalho é achado pelo conteúdo ("Data da venda").
"""
import re

import pandas as pd
from python_calamine import CalamineWorkbook

from pdv_pix import dia_movimento

NUM = ["Valor Bruto", "Valor da taxa (MDR)", "Valor líquido", "Valor cancelado"]
# O "Relatório de vendas detalhado" tem as mesmas vendas do Histórico (simplificado) com outros nomes de coluna
# e mais seis colunas de liquidação (data prevista, data efetiva e código do pagamento). Lê-se igual.
DETALHADO = {"Código de autorização": "Cód. de autorização", "Código do estabelecimento": "Número do estabelecimento",
             "Valor bruto transação": "Valor Bruto", "Valor líquido da parcela/transação": "Valor líquido", "Cód. do pedido": "Código do pedido"}


def ler_portal(caminho: str) -> tuple[pd.DataFrame, dict]:
    wb = CalamineWorkbook.from_path(caminho)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    h = next(i for i, r in enumerate(rows) if "Data da venda" in r)
    meta = {}
    for r in rows[:h]:
        txt = str(r[0]).strip()
        if ":" in txt:
            k, v = txt.split(":", 1)
            meta[k.strip()] = v.strip()
    m = re.search(r"(\d{8})", meta.get("Estabelecimento comercial", ""))
    meta["EC"] = m.group(1) if m else None

    df = pd.DataFrame(rows[h + 1 :], columns=rows[h]).rename(columns=DETALHADO)
    df = df[df["Data da venda"].astype(str).str.match(r"\d{2}/\d{2}/\d{4}")]
    for c in NUM:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0) if c in df.columns else 0.0
    df["dt"] = pd.to_datetime(
        df["Data da venda"] + " " + df["Hora da venda"].astype(str).str.strip(), dayfirst=True
    )
    df["Data calendário"] = df["dt"].dt.date
    df["Data"] = dia_movimento(df["dt"])          # mesmo dia do PDV: a madrugada fica no dia anterior
    df["cent"] = (df["Valor Bruto"] * 100).round().astype(int)
    df["EC"] = df["Número do estabelecimento"].astype(str)
    return df.reset_index(drop=True), meta


def ler_portais(arquivos, lojas: dict) -> dict:
    """Lê vários Históricos de vendas e junta os de uma mesma loja (ex.: duas semanas) num só DataFrame.
    Devolve {(filial, ec): df}, sem vendas repetidas quando os exports se sobrepõem."""
    por_loja = {}
    for arq in arquivos:
        df, meta = ler_portal(arq)
        ec = meta["EC"]
        filial = lojas.get(ec)
        if not filial:
            raise SystemExit(f"EC {ec} ({arq}) não está em lojas.csv — adicione a linha EC,Filial")
        por_loja.setdefault((filial, ec), []).append(df)
    return {k: pd.concat(v, ignore_index=True).drop_duplicates(subset=["Cód. de autorização", "dt", "cent"]).sort_values("dt").reset_index(drop=True)
            for k, v in sorted(por_loja.items())}


def pix_do_portal(df: pd.DataFrame) -> pd.DataFrame:
    """PIX vem como Modalidade 'Via QRCODE' (Produto/Bandeira = PIX), status 'Autorizada'."""
    return df[df["Modalidade"].str.upper().str.contains("QRCODE|PIX")].copy()


def ler_tef(caminho: str) -> pd.DataFrame:
    """Relatório de PIX pela TEF (colunas Data, Hora, Valor, Status, Loja - Descrição).
    Sai com as mesmas colunas que pix_do_portal, para entrar junto no casamento."""
    import glob
    arqs = [a for parte in str(caminho).split(",") for a in (sorted(glob.glob(parte.strip())) or [parte.strip()])]
    partes = []
    for a in arqs:
        wb = CalamineWorkbook.from_path(a)
        rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
        h = next(i for i, r in enumerate(rows) if "Método Pagamento" in r or "Loja - Descrição" in r)
        partes.append(pd.DataFrame(rows[h + 1 :], columns=rows[h]))
    df = pd.concat(partes, ignore_index=True)
    df = df[df["Data"].astype(str).str.match(r"\d{2}/\d{2}/\d{4}")].copy()
    df = df.drop_duplicates(subset=[c for c in ("Data", "Hora", "Valor", "Status", "Loja - Descrição") if c in df.columns])
    df = df[df["Método Pagamento"].astype(str).str.lower().str.contains("pix")]
    df["Valor Bruto"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
    df["dt"] = pd.to_datetime(df["Data"] + " " + df["Hora"].astype(str).str.strip(), dayfirst=True)
    df["Data calendário"] = df["dt"].dt.date
    df["Data"] = dia_movimento(df["dt"])
    df["cent"] = (df["Valor Bruto"] * 100).round().astype(int)
    df["Filial"] = df["Loja - Descrição"].astype(str).str.strip()
    df["Status"] = df["Status"].astype(str).str.strip()
    df["Modalidade"] = "PIX TEF"
    df["Canal"] = "TEF PIX"
    df["Cód. de autorização"] = df["Código Pedido"].astype(str).replace("", "-") if "Código Pedido" in df.columns else "-"
    df["Número do terminal"] = "TEF"
    df["Comprovante de venda"] = "-"
    df["Valor da taxa (MDR)"] = 0.0
    df["Valor líquido"] = df["Valor Bruto"]
    return df.reset_index(drop=True)
