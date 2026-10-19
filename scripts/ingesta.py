"""Identifica arquivos exportados (PDV, Sicredi vendas, antecipação, extrato, PIX TEF) pelo conteúdo e guarda em entrada/.

    python3 scripts/ingesta.py <arquivo.xlsx> [...]

Nomeia por loja e período; avisa quando o período não é o do fechamento (ex.: mês errado).
"""
import re
import shutil
import sys
from pathlib import Path

import pandas as pd
from python_calamine import CalamineWorkbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from concilia_pix import carrega_lojas  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
HOJE = pd.Timestamp.today().strftime("%Y-%m-%d")
COD = {"BIGGS 01": "01", "BIGGS 02": "02", "BIGGS 03": "03", "BIGGS 04": "04", "BIGGS 05": "05", "BIGGS 06": "06", "BIGGS 07": "07",
       "BIGGS 08": "08", "BIGGS 10": "10", "BIGGS 11": "11", "BIGGS 12": "12", "BIGGS 13": "13", "ISAIAS 01": "isa01", "ISAIAS 02": "isa02",
       "ISAIAS 03": "isa03", "MAPOLI": "mapoli01", "ROBS": "robs01"}


def cod_loja(filial: str) -> str:
    for k, v in COD.items():
        if filial.upper().startswith(k):
            return v
    return re.sub(r"\W+", "", filial.lower())[:10]


def ddmm(s):
    return pd.to_datetime(s, dayfirst=True).strftime("%Y-%m-%d")


def identifica(arq: str, lojas: dict):
    wb = CalamineWorkbook.from_path(arq)
    rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    topo = " | ".join(str(r[0]) for r in rows[:14])
    cab = next((r for r in rows[:20] if sum(1 for x in r if str(x).strip()) >= 5), [])
    cab = [str(c) for c in cab]
    if "Desc. pagam." in cab:
        h = rows.index(next(r for r in rows if "Desc. pagam." in r))
        df = pd.DataFrame(rows[h + 1:], columns=rows[h]); df = df[df["Filial"].astype(str).str.strip() != ""]
        d = pd.to_datetime(df["Data"], dayfirst=True).dt.date
        return "pdv", f"entrada/pdv_cupons_{d.min():%Y-%m-%d}_a_{d.max():%Y-%m-%d}.xlsx", f"PDV cupons {d.min():%d/%m}–{d.max():%d/%m} · {len(df)} linhas · {df['Filial'].nunique()} lojas", (d.min(), d.max())
    if "Data da venda" in cab:
        h = rows.index(next(r for r in rows if "Data da venda" in r))
        df = pd.DataFrame(rows[h + 1:], columns=rows[h]); df = df[df["Data da venda"].astype(str).str.match(r"\d{2}/\d{2}/\d{4}")]
        ec = re.search(r"(\d{8})", " ".join(str(r[0]) for r in rows[:h]))
        ec = ec.group(1) if ec else str(df.get("Número do estabelecimento", df.get("Código do estabelecimento")).iloc[0])
        filial = lojas.get(ec, f"EC {ec}")
        d = pd.to_datetime(df["Data da venda"], dayfirst=True).dt.date
        tipo = "detalhado" if "Código do pagamento" in cab else "simplificado"
        return "vendas", f"entrada/portal/vendas_{cod_loja(filial)}_{d.min():%d%m%Y}-{d.max():%d%m%Y}_{tipo}_{HOJE}.xlsx", f"Vendas Sicredi {tipo} · {filial} · {d.min():%d/%m}–{d.max():%d/%m} · {len(df)} linhas", (d.min(), d.max())
    if "Número da simulação" in cab:
        h = rows.index(next(r for r in rows if "Número da simulação" in r))
        df = pd.DataFrame(rows[h + 1:], columns=[str(c) for c in rows[h]]); df = df[df["Data do pagamento"].astype(str).str.match(r"\d{2}/")]
        ec = re.search(r"EC: ?(\d{8})", topo); ec = ec.group(1) if ec else "?"
        filial = lojas.get(ec, f"EC {ec}")
        p = pd.to_datetime(df["Data do pagamento"], dayfirst=True).dt.date
        tipo = "detalhado" if "Data das vendas" in cab else "simplificado"
        return "antecipacao", f"entrada/antecipacao/antecipacao_{tipo}_{cod_loja(filial)}_pgto_{p.min():%Y-%m-%d}_a_{p.max():%Y-%m-%d}_{HOJE}.xlsx", f"Antecipação {tipo} · {filial} · pagamentos {p.min():%d/%m}–{p.max():%d/%m} · {len(df)} linhas", (p.min(), p.max())
    if "Descrição" in cab and "Saldo após" in cab:
        h = rows.index(next(r for r in rows if "Descrição" in r))
        df = pd.DataFrame(rows[h + 1:], columns=rows[h])
        d = pd.to_datetime(df["Data"]).dt.date
        return "extrato", f"entrada/extrato_todas_{d.min():%Y-%m-%d}_a_{d.max():%Y-%m-%d}.xlsx", f"Extrato · {d.min():%d/%m}–{d.max():%d/%m} · {len(df)} linhas · {df['Loja'].nunique()} contas", (d.min(), d.max())
    if "Loja - Descrição" in cab or "Método Pagamento" in cab:
        h = rows.index(next(r for r in rows if "Loja - Descrição" in r or "Método Pagamento" in r))
        df = pd.DataFrame(rows[h + 1:], columns=rows[h]); df = df[df["Data"].astype(str).str.match(r"\d{2}/\d{2}/\d{4}")]
        d = pd.to_datetime(df["Data"], dayfirst=True).dt.date
        return "tef", f"entrada/pix_tef_{d.min():%Y-%m-%d}_a_{d.max():%Y-%m-%d}.xlsx", f"PIX TEF · {d.min():%d/%m}–{d.max():%d/%m} · {len(df)} linhas", (d.min(), d.max())
    return None, None, f"não reconhecido ({arq})", None


def main():
    lojas = carrega_lojas()
    mes = pd.Timestamp.today().replace(day=1).date()
    for arq in sys.argv[1:]:
        tipo, destino, desc, per = identifica(arq, lojas)
        if not tipo:
            print(f"✗ {desc}"); continue
        if per and per[1] < mes:
            print(f"✗ FORA DO MÊS · {desc} · não guardado"); continue
        dst = RAIZ / destino
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(arq, dst)
        print(f"✓ {desc} → {destino}")


if __name__ == "__main__":
    main()
