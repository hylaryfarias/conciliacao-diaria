#!/usr/bin/env python3
"""Leitor do 'Relatório Histórico de vendas' do portal da maquininha Sicredi.

Layout de 22 colunas (versão com 'Modalidade' e 'Hora da venda'), exportado em .xlsx.
Preâmbulo de tamanho variável: o cabeçalho é achado pelo conteúdo ("Data da venda").
"""
import re

import pandas as pd
from python_calamine import CalamineWorkbook

NUM = ["Valor Bruto", "Valor da taxa (MDR)", "Valor líquido", "Valor cancelado"]


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

    df = pd.DataFrame(rows[h + 1 :], columns=rows[h])
    df = df[df["Data da venda"].astype(str).str.match(r"\d{2}/\d{2}/\d{4}")]
    for c in NUM:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["dt"] = pd.to_datetime(
        df["Data da venda"] + " " + df["Hora da venda"].astype(str).str.strip(), dayfirst=True
    )
    df["Data"] = df["dt"].dt.date
    df["cent"] = (df["Valor Bruto"] * 100).round().astype(int)
    df["EC"] = df["Número do estabelecimento"].astype(str)
    return df.reset_index(drop=True), meta


def pix_do_portal(df: pd.DataFrame) -> pd.DataFrame:
    """PIX vem como Modalidade 'Via QRCODE' (Produto/Bandeira = PIX), status 'Autorizada'."""
    return df[df["Modalidade"].str.upper().str.contains("QRCODE|PIX")].copy()
