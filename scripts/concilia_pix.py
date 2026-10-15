#!/usr/bin/env python3
"""Etapa 1 · PIX: cruza o relatório de cupons do PDV com o Histórico de vendas do portal Sicredi.

Uso:
    python3 scripts/concilia_pix.py <pdv_cupons.xlsx> <portal1.xlsx> [portal2.xlsx ...] [--tef pix_tef.xlsx]

O PIX tem duas fontes na ponta da maquininha: o portal Sicredi ("Via QRCODE", POS) e o
relatório de PIX pela TEF (--tef), que o portal não mostra. As duas entram juntas no casamento;
a coluna Canal diz de onde veio cada uma.

A loja de cada arquivo do portal é identificada pelo EC (Estabelecimento comercial) via
`lojas.csv` (colunas: EC, Filial). Casamento: mesma loja, mesmo dia, mesmo valor em
centavos, hora mais próxima. A hora do PDV é a do fechamento do cupom, que pode ser
bem depois do pagamento na máquina, por isso a hora só desempata — não bloqueia.

Só contam PIX autorizados/aprovados na maquininha. Recusado e cancelado ficam fora e
são listados à parte. PIX sem taxa: valor previsto = bruto.

Data do PDV é o DIA DE MOVIMENTO: cupom de madrugada (antes de 06:00) pertence ao dia
seguinte do calendário, que é o que a maquininha usa. Verificado em 299 de 301 caixas-dia.

PIX do PDV sem par na maquininha: procura um débito/crédito aprovado na máquina, de mesmo
valor, no mesmo dia ou no dia anterior, que também não tenha par no PDV. Se achar, marca
"PROVÁVEL FORMA ERRADA" — o caixa lançou cartão como PIX.
"""
HORA_VIRADA = 6  # cupom antes dessa hora conta no dia calendário seguinte
JANELA_ANTES_MIN, JANELA_DEPOIS_H = 30, 36  # PDV pode fechar o cupom antes/depois do PIX na máquina
GRUPO_MAX, GRUPO_CAND, GRUPO_ESPALHA_H = 4, 24, 3  # conta dividida: até 4 por grupo, todos em até 3 h
TAXA_SERVICO = 0.10  # restaurante: a máquina cobra o cupom + 10% de serviço, o PDV registra só o cupom
CARTAO_PDV = ("CARTAO CREDITO", "CARTAO DEBITO", "TEF - CREDITO", "TEF - DEBITO")
import sys
from itertools import combinations
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from pdv_pix import ler_pdv, pix_do_pdv  # noqa: E402
from portal_sicredi import ler_portal, ler_tef, pix_do_portal  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
OK_STATUS = {"AUTORIZADA", "APROVADA"}


def carrega_lojas() -> dict:
    f = RAIZ / "lojas.csv"
    if not f.exists():
        return {}
    t = pd.read_csv(f, dtype=str)
    return dict(zip(t["EC"].str.strip(), t["Filial"].str.strip()))


def casa(pdv: pd.DataFrame, maq: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Casa linha a linha. Retorna (resultado por linha do PDV, linhas da máquina sem par)."""
    usados: set = set()
    linhas = []
    pdv = pdv.sort_values("dt_pdv")
    for _, a in pdv.iterrows():
        # mesmo valor; o cupom do PDV fecha de 30 min antes até 36 h depois do PIX na máquina
        dif = (a["dt_pdv"] - maq["dt"]).dt.total_seconds() / 60
        cand = maq[(maq["cent"] == a["cent"]) & (dif >= -JANELA_ANTES_MIN) & (dif <= JANELA_DEPOIS_H * 60) & (~maq.index.isin(usados))]
        if len(cand) == 0:
            linhas.append({**a.to_dict(), "Status": "SEM PAR NA MAQUININHA"})
            continue
        j = (cand["dt"] - a["dt_pdv"]).abs().idxmin()
        usados.add(j)
        m = maq.loc[j]
        linhas.append({
            **a.to_dict(),
            "Hora maq": m["dt"].strftime("%H:%M:%S"),
            "Dia maq": m["Data"],
            "Autorização": m["Cód. de autorização"],
            "Terminal": m["Número do terminal"],
            "Comprovante": m["Comprovante de venda"],
            "Canal": m["Canal"],
            "Dif. hora (min)": round((a["dt_pdv"] - m["dt"]).total_seconds() / 60),
            "Status": "CONCILIADO",
        })
    sobra = maq[~maq.index.isin(usados)].copy()
    return pd.DataFrame(linhas), sobra


def _subconjunto(alvo: int, cand: pd.DataFrame, ref_dt) -> list | None:
    """Menor grupo (2..GRUPO_MAX) de linhas de `cand` cuja soma em centavos é exatamente `alvo`.
    Empate: o grupo mais próximo no tempo de ref_dt."""
    cand = cand.copy()
    cand["dist"] = (cand["dt"] - ref_dt).abs().dt.total_seconds()
    cand = cand.sort_values("dist").head(GRUPO_CAND)
    idx, cents, dist, dts = list(cand.index), list(cand["cent"]), list(cand["dist"]), list(cand["dt"])
    for k in range(2, min(GRUPO_MAX, len(idx)) + 1):
        melhor = None
        for comb in combinations(range(len(idx)), k):
            if sum(cents[i] for i in comb) == alvo and (max(dts[i] for i in comb) - min(dts[i] for i in comb)) <= pd.Timedelta(hours=GRUPO_ESPALHA_H):
                d = sum(dist[i] for i in comb)
                if melhor is None or d < melhor[0]:
                    melhor = (d, [idx[i] for i in comb])
        if melhor:
            return melhor[1]
    return None


def casa_servico(res: pd.DataFrame, sobra: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Cupom sem par cujo valor × (1 + TAXA_SERVICO) bate exato com um PIX sem par na janela."""
    if not len(res) or not len(sobra):
        return res, sobra
    usados: set = set()
    for i, a in res[res["Status"] == "SEM PAR NA MAQUININHA"].iterrows():
        alvo = int(round(a["cent"] * (1 + TAXA_SERVICO)))
        dif = (a["dt_pdv"] - sobra["dt"]).dt.total_seconds() / 60
        cand = sobra[(sobra["cent"] == alvo) & (dif >= -JANELA_ANTES_MIN) & (dif <= JANELA_DEPOIS_H * 60) & (~sobra.index.isin(usados))]
        if len(cand) == 0:
            continue
        j = (cand["dt"] - a["dt_pdv"]).abs().idxmin()
        usados.add(j)
        m = sobra.loc[j]
        res.loc[i, ["Status", "Hora maq", "Dia maq", "Autorização", "Terminal", "Comprovante", "Canal", "Dif. hora (min)", "Serviço (R$)"]] = [
            f"SERVIÇO NÃO BAIXADO · máquina cobrou R$ {m['cent']/100:.2f} (+{TAXA_SERVICO:.0%}), PDV fechou sem a taxa", m["dt"].strftime("%H:%M:%S"), m["Data"],
            str(m["Cód. de autorização"]), str(m["Número do terminal"]), str(m["Comprovante de venda"]), str(m["Canal"]),
            round((a["dt_pdv"] - m["dt"]).total_seconds() / 60), (m["cent"] - a["cent"]) / 100]
    return res, sobra[~sobra.index.isin(usados)]


def casa_grupos(res: pd.DataFrame, sobra: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Conta dividida. 1:N — um cupom pago por vários PIX na máquina; N:1 — um PIX que paga vários cupons."""
    if not len(res) or not len(sobra):
        return res, sobra
    usados: set = set()
    pend = res[res["Status"] == "SEM PAR NA MAQUININHA"].sort_values("cent", ascending=False)
    # 1 cupom : N PIX
    for i, a in pend.iterrows():
        dif = (a["dt_pdv"] - sobra["dt"]).dt.total_seconds() / 60
        cand = sobra[(dif >= -JANELA_ANTES_MIN) & (dif <= JANELA_DEPOIS_H * 60) & (~sobra.index.isin(usados)) & (sobra["cent"] < a["cent"])]
        g = _subconjunto(int(a["cent"]), cand, a["dt_pdv"]) if len(cand) >= 2 else None
        if not g:
            continue
        usados.update(g)
        m = sobra.loc[g].sort_values("dt")
        res.loc[i, ["Status", "Hora maq", "Dia maq", "Autorização", "Terminal", "Comprovante", "Canal", "Dif. hora (min)"]] = [
            f"CONCILIADO (grupo 1:{len(g)})", m["dt"].iloc[0].strftime("%H:%M:%S"), m["Data"].iloc[0],
            " + ".join(str(x) for x in m["Cód. de autorização"]), " + ".join(str(x) for x in m["Número do terminal"].unique()),
            " + ".join(str(x) for x in m["Comprovante de venda"]), " + ".join(str(x) for x in m["Canal"].unique()),
            round((a["dt_pdv"] - m["dt"].iloc[0]).total_seconds() / 60)]
    sobra = sobra[~sobra.index.isin(usados)]
    # N cupons : 1 PIX
    usados_p: set = set()
    usados_m: set = set()
    pend = res[res["Status"] == "SEM PAR NA MAQUININHA"]
    for j, m in sobra.sort_values("cent", ascending=False).iterrows():
        dif = (pend["dt_pdv"] - m["dt"]).dt.total_seconds() / 60
        cand = pend[(dif >= -JANELA_ANTES_MIN) & (dif <= GRUPO_ESPALHA_H * 60) & (~pend.index.isin(usados_p)) & (pend["cent"] < m["cent"])].rename(columns={"dt_pdv": "dt"})
        g = _subconjunto(int(m["cent"]), cand, m["dt"]) if len(cand) >= 2 else None
        if not g:
            continue
        usados_p.update(g); usados_m.add(j)
        for i in g:
            res.loc[i, ["Status", "Hora maq", "Dia maq", "Autorização", "Terminal", "Comprovante", "Canal", "Dif. hora (min)"]] = [
                f"CONCILIADO (grupo {len(g)}:1 · PIX de R$ {m['cent']/100:.2f})", m["dt"].strftime("%H:%M:%S"), m["Data"],
                str(m["Cód. de autorização"]), str(m["Número do terminal"]), str(m["Comprovante de venda"]), str(m["Canal"]),
                round((res.loc[i, "dt_pdv"] - m["dt"]).total_seconds() / 60)]
    return res, sobra[~sobra.index.isin(usados_m)]


def caca_forma_errada(res: pd.DataFrame, maq: pd.DataFrame, cartao_pdv: pd.DataFrame) -> pd.DataFrame:
    """Para cada PIX do PDV sem par, procura cartão aprovado na máquina (D ou D-1), mesmo valor,
    que também não tenha par no PDV. Órfão com órfão de mesmo valor = provável forma errada."""
    if "Status" not in res.columns or res["Status"].str.startswith("CONCILIADO").all():
        return res
    cart_m = maq[maq["Modalidade"].isin(["Crédito", "Débito"]) & (maq["Status"].str.upper() == "APROVADA")]
    usados: set = set()
    for _, x in cart_m.sort_values("dt").iterrows():
        dif = (cartao_pdv["dt_pdv"] - x["dt"]).dt.total_seconds() / 60
        c = cartao_pdv[(cartao_pdv["cent"] == x["cent"]) & (dif >= -30) & (dif <= 36 * 60)
                       & (~cartao_pdv.index.isin(usados))]
        if len(c):
            usados.add((c["dt_pdv"] - x["dt"]).abs().idxmin())
    orfaos = cart_m[~cart_m.index.isin(usados)].copy()  # cartão na máquina sem lançamento no PDV
    usados_m: set = set()
    for i, a in res[res["Status"] == "SEM PAR NA MAQUININHA"].iterrows():
        dias = {a["Data"], a["Data"] - pd.Timedelta(days=1)}
        c = orfaos[(orfaos["cent"] == a["cent"]) & (orfaos["Data"].isin(dias)) & (~orfaos.index.isin(usados_m))]
        if len(c) == 0:
            continue
        j = (c["dt"] - a["dt_pdv"]).abs().idxmin()
        usados_m.add(j)
        m = orfaos.loc[j]
        res.loc[i, ["Status", "Hora maq", "Autorização", "Terminal", "Comprovante", "Dif. hora (min)"]] = [
            f"PROVÁVEL FORMA ERRADA: {m['Modalidade']} {m['Bandeira']} na máquina em {m['dt']:%d/%m %H:%M}",
            m["dt"].strftime("%H:%M:%S"), m["Cód. de autorização"], m["Número do terminal"],
            m["Comprovante de venda"], round((a["dt_pdv"] - m["dt"]).total_seconds() / 60),
        ]
    return res


def main() -> None:
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    args = sys.argv[1:]
    tef = pd.DataFrame()
    if "--tef" in args:
        i = args.index("--tef"); tef = ler_tef(args[i + 1]); del args[i:i + 2]
    pdv_arq, portais = args[0], args[1:]
    lojas = carrega_lojas()

    bruto = ler_pdv(pdv_arq)
    bruto["dt_pdv"] = pd.to_datetime(bruto["Data"].astype(str) + " " + bruto["Hora"].astype(str).str[:5])
    madrugada = bruto["dt_pdv"].dt.hour < HORA_VIRADA
    bruto.loc[madrugada, "dt_pdv"] += pd.Timedelta(days=1)
    bruto["Dia movimento"] = bruto["Data"]
    bruto["Data"] = bruto["dt_pdv"].dt.date          # data calendário, igual à da maquininha
    pdv_all = pix_do_pdv(bruto)
    cartao_pdv = bruto[bruto["Desc. pagam."].str.upper().isin(CARTAO_PDV)].copy()

    todas, sobras, recusados, resumo = [], [], [], []
    for arq in portais:
        df, meta = ler_portal(arq)
        ec = meta["EC"]
        filial = lojas.get(ec)
        if not filial:
            sys.exit(f"EC {ec} ({arq}) não está em lojas.csv — adicione a linha EC,Filial")
        maq = pix_do_portal(df)
        d0, d1 = df["Data"].min(), df["Data"].max()
        if len(tef):
            t = tef[(tef["Filial"] == filial) & (tef["Data"] >= d0) & (tef["Data"] <= d1)]
            maq = pd.concat([maq, t[[c for c in t.columns if c in maq.columns or c in ("dt", "Data", "cent", "Canal")]]], ignore_index=True)
        ok = maq[maq["Status"].str.upper().isin(OK_STATUS)]
        rec = maq[~maq["Status"].str.upper().isin(OK_STATUS)]
        p = pdv_all[(pdv_all["Filial"] == filial) & (pdv_all["Data"] >= d0) & (pdv_all["Data"] <= d1)]

        res, sobra = casa(p, ok)
        res, sobra = casa_servico(res, sobra)
        res, sobra = casa_grupos(res, sobra)
        res["EC"] = ec
        res = caca_forma_errada(res, df, cartao_pdv[cartao_pdv["Filial"] == filial])
        sobra["Filial"] = filial
        rec = rec.assign(Filial=filial)
        todas.append(res); sobras.append(sobra); recusados.append(rec)

        for dia, g in res.groupby("Data"):
            conc = g[g["Status"].str.startswith("CONCILIADO")]["cent"].sum()
            servico = g["Serviço (R$)"].sum() if "Serviço (R$)" in g.columns else 0.0
            sem_par = g[~g["Status"].str.startswith("CONCILIADO") & ~g["Status"].str.startswith("SERVIÇO")]["cent"].sum()
            maq_dia = ok[ok["Data"] == dia]["cent"].sum()
            sobra_dia = sobra[sobra["Data"] == dia]["cent"].sum()
            resumo.append({
                "Filial": filial, "EC": ec, "Data": dia,
                "PIX PDV (qtd)": len(g), "PIX PDV (R$)": g["cent"].sum() / 100,
                "PIX maquininha (qtd)": int((ok["Data"] == dia).sum()), "PIX maquininha (R$)": maq_dia / 100,
                "Conciliado (R$)": conc / 100,
                "Serviço não baixado (R$)": round(float(servico), 2),
                "PDV sem par (R$)": sem_par / 100,
                "Maquininha sem par (R$)": sobra_dia / 100,
                "Δ PDV − maquininha (R$)": (g["cent"].sum() - maq_dia) / 100,
            })
        # dias que só existem na máquina
        for dia in sorted(set(ok["Data"]) - set(res["Data"])):
            maq_dia = ok[ok["Data"] == dia]["cent"].sum()
            resumo.append({"Filial": filial, "EC": ec, "Data": dia, "PIX PDV (qtd)": 0, "PIX PDV (R$)": 0,
                           "PIX maquininha (qtd)": int((ok["Data"] == dia).sum()), "PIX maquininha (R$)": maq_dia / 100,
                           "Conciliado (R$)": 0, "PDV sem par (R$)": 0, "Maquininha sem par (R$)": maq_dia / 100,
                           "Δ PDV − maquininha (R$)": -maq_dia / 100})
        print(f"{filial} (EC {ec}) {d0:%d/%m}–{d1:%d/%m}: PDV {len(p)} PIX R$ {p['cent'].sum()/100:,.2f} | "
              f"maquininha {len(ok)} PIX R$ {ok['cent'].sum()/100:,.2f} | "
              f"sem par PDV {int((~res['Status'].str.startswith('CONCILIADO') & ~res['Status'].str.startswith('SERVIÇO')).sum())} | serviço não baixado {int(res['Status'].str.startswith('SERVIÇO').sum())} | grupos {int(res['Status'].str.contains('grupo').sum())} | sem par maq {len(sobra)} | recusados {len(rec)}")

    res = pd.concat(todas, ignore_index=True)
    resumo = pd.DataFrame(resumo).sort_values(["Filial", "Data"])
    sobra = pd.concat(sobras, ignore_index=True)
    rec = pd.concat(recusados, ignore_index=True)

    col_pdv = ["Filial", "EC", "Data", "Dia movimento", "Hora", "Cupom", "Caixa", "Desc. pagam.", "Vl. pagamento",
               "Status", "Hora maq", "Dia maq", "Dif. hora (min)", "Autorização", "Terminal", "Comprovante", "Canal", "Serviço (R$)"]
    col_maq = ["Filial", "EC", "Data", "dt", "Cód. de autorização", "Valor Bruto", "Status",
               "Número do terminal", "Comprovante de venda", "Canal"]
    d0, d1 = resumo["Data"].min(), resumo["Data"].max()
    out = RAIZ / "saida" / f"concilia_pix_{d0:%Y-%m-%d}_a_{d1:%Y-%m-%d}.xlsx"
    with pd.ExcelWriter(out) as xw:
        resumo.to_excel(xw, sheet_name="Resumo loja x dia", index=False)
        res[[c for c in col_pdv if c in res.columns]].to_excel(xw, sheet_name="PDV x maquininha", index=False)
        res[~res["Status"].str.startswith("CONCILIADO")][[c for c in col_pdv if c in res.columns]].to_excel(
            xw, sheet_name="PDV sem par", index=False)
        res[res["Status"].str.startswith("SERVIÇO")][[c for c in col_pdv if c in res.columns]].to_excel(
            xw, sheet_name="Serviço não baixado", index=False)
        sobra[[c for c in col_maq if c in sobra.columns]].to_excel(xw, sheet_name="Maquininha sem par", index=False)
        rec[[c for c in col_maq if c in rec.columns]].to_excel(xw, sheet_name="Recusados-cancelados", index=False)
    print(f"\ngravado em {out}")
    print(resumo.to_string(index=False))


if __name__ == "__main__":
    main()
