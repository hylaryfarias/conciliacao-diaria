#!/usr/bin/env bash
# Roda o cruzamento completo e gera o painel. Uso: scripts/roda_tudo.sh <de> <ate>   (datas AAAA-MM-DD, dia de movimento)
set -e
cd "$(dirname "$0")/.."
DE=${1:-2026-10-01}; ATE=${2:-2026-10-14}
PDV="entrada/pdv_cupons_*.xlsx"; EXT="entrada/extrato_todas_*.xlsx"; TEF=$(ls entrada/pix_tef_*.xlsx | tr '\n' ',' | sed 's/,$//')
rm -f saida/concilia_pix_2*.xlsx saida/concilia_credito_*.xlsx saida/concilia_debito_*.xlsx saida/concilia_pix_banco_*.xlsx saida/concilia_cupom.xlsx
python3 scripts/concilia_pix.py "$PDV" entrada/portal/*.xlsx --tef "$TEF" > /dev/null
python3 scripts/concilia_cartao.py "$PDV" entrada/portal/*.xlsx > /dev/null
python3 scripts/concilia_pix_banco.py "$EXT" "$PDV" entrada/portal/*.xlsx --tef "$TEF" --de "$DE" --ate "$ATE" > /dev/null
python3 scripts/concilia_cupom.py > /dev/null
python3 scripts/gera_painel.py "$PDV" --de "$DE" --ate "$ATE"
