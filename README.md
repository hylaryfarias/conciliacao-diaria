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
