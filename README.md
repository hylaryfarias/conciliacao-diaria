# Conciliação manual · Grupo Ragga

Conciliação diária de PIX, débito e crédito em três pontas: **PDV** (relatório de cupons),
**maquininha Sicredi** (relatórios do portal) e **extrato bancário**.

## Etapa 1 · PIX: PDV × maquininha

```bash
pip install python-calamine pandas openpyxl
python3 scripts/pdv_pix.py entrada/<relatorio_cupons>.xlsx saida
```

Gera `saida/pdv_pix_<inicio>_a_<fim>.xlsx` com o resumo de PIX por loja e dia e a lista
por cupom. Formas de pagamento do PDV tratadas como PIX de maquininha:
`PIX MAQUININHA`, `PIX - TEF`, `TEF - PIX`.

## Pastas

- `entrada/` arquivos recebidos (PDV, portal Sicredi, extrato)
- `scripts/` leitores e motor de cruzamento
- `saida/` resultados de cada rodada

## Etapa 2 · PIX: maquininha × extrato

```bash
python3 scripts/concilia_pix_banco.py entrada/<extrato>.xlsx entrada/<pdv_cupons>.xlsx entrada/<portal_loja>.xlsx ...
```

Período de fechamento = o coberto pelo portal (ou --de/--ate). Regra medida (ROBS, 54 de 54): PIX cai **D0 em dia útil**; sábado, domingo e
feriado caem no **próximo dia útil**. Sem taxa. Venda após a última data do extrato fica
como **"aguarda extrato"**, nunca como divergência.

## Regras aprendidas no dado

- **Data do PDV é dia de movimento.** Cupom antes das 06:00 pertence ao dia seguinte do calendário
  (299 de 301 caixas-dia). A maquininha e o banco usam data calendário.
- **Hora do PDV é a do fechamento do cupom**, até ~80 min depois do pagamento na máquina. Só desempata.
- **Máquinas POS (terminais APT…) não são integradas**: o caixa digita a venda, às vezes no dia
  seguinte e às vezes com a forma errada. A TEF (terminal TFI…) bate hora a hora.
- **PIX só existe pela máquina** (QR code/TEF). PIX do PDV sem par na máquina → procurar cartão
  órfão de mesmo valor em D ou D−1 (o motor já faz e marca "PROVÁVEL FORMA ERRADA").

## Crédito e débito · etapa 1

```bash
python3 scripts/concilia_cartao.py entrada/<pdv_cupons>.xlsx entrada/<portal_loja>.xlsx ...
```

Gera `saida/concilia_credito_*.xlsx` e `saida/concilia_debito_*.xlsx`: PDV × maquininha pelo
valor bruto, cupom a cupom (PDV pode lançar de 30 min antes a 36 h depois da máquina). A ponta
do banco (D+1 útil, MDR, antecipação) é a etapa 2, ainda não iniciada.

## Painel

```bash
python3 scripts/gera_painel.py entrada/<pdv_cupons>.xlsx --de 2026-10-08 --ate 2026-10-14
```

Lê as planilhas de `saida/` e grava `painel/index.html` (template em `painel/template.html`,
cores e tipografia do Manual de Identidade Visual · Grupo Ragga). Publicado em
https://claude.ai/artifact/8hmLnVrM6CiE3omwzP7oms — republicar a cada rodada.
