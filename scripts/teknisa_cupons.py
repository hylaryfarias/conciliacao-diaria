"""Converte o export 'Vendas Realizadas no Período' da Teknisa Retail (XLS todas as colunas - com grupo, agrupar por
nível de árvore = Não) para o layout do Relatório de cupons de vendas do Cloudfy, que o motor já lê.
Uso: python3 scripts/teknisa_cupons.py <export.xlsx> "<Filial canônica>" <sigla caixa> <saida.xlsx>
Ex.: python3 scripts/teknisa_cupons.py export.xlsx "BIGGS 07 - JAMIL E ACAI" BG07 entrada/pdv_cupons_teknisa_biggs07_2026-11-18_a_2026-11-25.xlsx
Uma linha por forma de pagamento de cada venda; Data = dia de movimento (antes das 10h é o dia anterior), Hora = relógio."""
import re, sys
import pandas as pd

VIRADA_H = 10
def forma(t: str) -> str:
    u = t.upper().strip()
    tef = "(TEF)" in u or u.endswith(" TEF")
    if "PIX" in u: return "TEF - PIX" if tef else "PIX MAQUININHA"
    if any(v in u for v in ("ALELO", "VR ", "TICKET", "PLUXEE", "SODEXO", "VOUCHER", "BEN ", "FLASH")): return "TEF - VOUCHER" if tef else "VOUCHER"
    if "DEBITO" in u: return "TEF - DEBITO" if tef else "CARTAO DEBITO"
    if "CREDITO" in u or "AMEX" in u or "HIPERCARD" in u: return "TEF - CREDITO" if tef else "CARTAO CREDITO"
    if u == "DINHEIRO": return "DINHEIRO"
    if "IFOOD" in u or "DELIVERY" in u or "ONLINE" in u: return "Pagamento online"
    if "PRAZO" in u or "FUNCIONARIO" in u: return "Venda a prazo"
    return t.strip()

def converte(arq: str, filial: str, sigla: str) -> pd.DataFrame:
    d = pd.read_excel(arq, header=None)
    out, itens = [], 0
    for _, r in d.iterrows():
        c0 = str(r[0])
        if c0.startswith("Total (Venda:"):
            m = re.search(r"Caixa : (\d+) \| Nr\. Ven\.: (\d+) \| Dt\. Ven\.: (\d\d/\d\d/\d{4} \d\d:\d\d:\d\d) \| Cup\./N\. Fiscal: ([^|]*)\| Cli\./Cons\.: ([^|]*)\|.*Modalidade: ([^|)]*)(?:\|(.*))?\)$", c0)
            if not m:
                itens = 0; continue
            caixa, nven, dt, cup, cli, mod, pags = m.groups()
            dt = pd.to_datetime(dt, dayfirst=True); mov = (dt - pd.Timedelta(hours=VIRADA_H)).date()
            for p in [x.strip() for x in (pags or "").split("|") if x.strip()]:
                f, v = p.rsplit(": R$", 1)
                out.append({"Filial": filial, "Caixa": f"{caixa} - {sigla} - TEKNISA {caixa}", "Data": mov.strftime("%d/%m/%Y"), "Cupom": int(caixa) * 1_000_000 + int(nven), "Hora": dt.strftime("%H:%M:%S"),   # Nr. Ven. repete entre caixas: o caixa entra no número
                            "Itens": itens, "Chave": cup.strip(), "CPF/CNPJ": "", "Nome do cliente": cli.strip(), "Desc. pagam.": forma(f), "Nr. parc": 1,
                            "Vl. pagamento": float(v.replace(".", "").replace(",", ".")), "Modalidade": mod.strip(), "Forma Teknisa": f.strip()})
            itens = 0
        elif c0 not in ("nan", "Código") and not c0.startswith(("Unidade:", "Total (")):
            itens += 1
    return pd.DataFrame(out)

if __name__ == "__main__":
    arq, filial, sigla, saida = sys.argv[1:5]
    df = converte(arq, filial, sigla)
    df.to_excel(saida, index=False)
    print(f"{len(df)} pagamentos · {df['Cupom'].nunique()} vendas · {df['Data'].iloc[0]} a {df['Data'].iloc[-1]} · gravado em {saida}")
    print(df.groupby("Desc. pagam.")["Vl. pagamento"].agg(["count", "sum"]).round(2).to_string())
