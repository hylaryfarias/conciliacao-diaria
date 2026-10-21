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
- **Serviço não baixado é divergência.** A taxa de serviço é um item do cupom. Quando a máquina
  cobra cupom × 1,10 e o PDV fecha sem a taxa, o cliente pagou a mais do que o PDV diz e a sobra
  sai do caixa sem registro. O motor casa os dois lados (status "SERVIÇO NÃO BAIXADO"), soma na
  coluna "Serviço não baixado" e o painel conta como em aberto. 74 casos em 01–14/10, 71 deles
  no caixa PDV 02 da ISAIAS 01, à noite.
- **Comanda bate, forma não** (`scripts/concilia_cupom.py`, roda depois das três formas): soma o
  que sobrou sem par em cada cupom e procura na máquina, em qualquer modalidade, 1 a 4 vendas
  com a mesma soma (36 h antes a 30 min depois do cupom, até 3 h entre si). Mesa dividida com
  divisão de formas diferente nos dois lados. O painel marca "conciliado por total da comanda".
- **Conta dividida** (motor, automático): cupom sem par procura (b) um grupo de 2 a 4 PIX, todos em até 3 h, cuja soma bate o cupom; (c) o inverso, um PIX que
  paga 2 a 4 cupons fechados em até 3 h. Janela larga demais nos grupos gera par falso.
- **PIX a mais no banco não trava a venda.** Crédito "Vendas no Pix" sem venda própria é classificado:
  (a) PIX PAGO · MÁQUINA NÃO REGISTROU — há baixa PIX no PDV sem par na máquina com o mesmo valor:
  a máquina recusou/cancelou, o PIX entrou, o cupom fecha contra o banco; (b) PIX EM DUPLICIDADE —
  o valor já foi pago por outro crédito: cliente pagou duas vezes, devolver (ou DEVOLVIDO quando há
  DEVOLUCAO PIX do mesmo pagador no extrato); (c) OUTRA NATUREZA — sem venda de mesmo valor.
  Janela das vendas do PDV passa a ser por dia de movimento (madrugada do último dia entra).
- **Sobras dos dois lados se compensam** (Hylary): baixa no PDV sem par e venda na máquina sem par
  são pagamento parcial de comanda. O painel abre o dia só pelo líquido |PDV sem par − máquina sem par|
  e mostra as duas listas em amarelo, para consulta.
- **Até R$ 1,00 é conciliado** (Hylary): último passo do casamento aceita diferença de até R$ 1,00
  entre PDV e máquina na janela; no painel, dia e loja só ficam vermelhos se o aberto passar de R$ 1,00.
- **PIX só existe pela máquina** (QR code/TEF). PIX do PDV sem par na máquina → procurar cartão
  órfão de mesmo valor em D ou D−1 (o motor já faz e marca "PROVÁVEL FORMA ERRADA").

## Crédito e débito · etapa 1

```bash
python3 scripts/concilia_cartao.py entrada/<pdv_cupons>.xlsx entrada/<portal_loja>.xlsx ...
```

Gera `saida/concilia_credito_*.xlsx` e `saida/concilia_debito_*.xlsx`: PDV × maquininha pelo
valor bruto, cupom a cupom (PDV pode lançar de 30 min antes a 36 h depois da máquina). A ponta
do banco (D+1 útil, MDR, antecipação) é a etapa 2, ainda não iniciada.

## Ajustes de forma de pagamento (reclassificação)

Quando o PDV baixou com a forma errada e o caixa já fechou, o ajuste é feito no painel
(botão "Transferir para…" na pendência), com justificativa, autor e hora. Fica no banco de dados
do artefato, coleção `ajustes`, e é lido de volta a cada rodada com a ferramenta de dados do
artefato. O painel aplica os ajustes em cima dos números gerados.

## Funcionamento das lojas

- MAPOLI não abre sábado e domingo: esses dias ficam sem venda nos dois lados (traço no painel).
- ISAIAS 01 não abre às terças à noite.

## O que o operador baixou vale

Dois cupons baixados e um só pagamento na máquina é falta de caixa, mesmo com cupons idênticos
(mesmo caixa, valor e minuto). O motor mantém "SEM PAR NA MAQUININHA" e preenche "Cupom igual" com o
vizinho que casou, para a loja explicar. Vale (Ticket, Alelo, Pluxee, Fepas, Elo/Visa voucher) passado na
máquina no mesmo minuto de uma baixa de cartão/PIX sem par vira "PROVÁVEL FORMA ERRADA: VOUCHER",
transferível no painel para Voucher.

## Transferência de forma com evidência é automática

Regra da Hylary (20/10): toda baixa com forma errada que tenha evidência na máquina (mesmo valor,
mesma loja, poucos minutos de diferença: PIX baixado e cartão na máquina, crédito × débito, dinheiro
baixado e PIX na máquina, vale baixado como cartão) é transferida direto, sem clique um a um. O registro
continua igual: um documento em `ajustes` no painel (id `<loja>_<data>_<cupom>_<origem>_<aut>`), com a
justificativa e a autorização da máquina, e a cópia em `ajustes/ajustes.csv`. Só o que não tem par
claro continua como pendência para decisão.

## Dia único: o dia de movimento

A venda vai da abertura do caixa até a reabertura da loja às 10h. O PDV já informa o dia de
movimento; a Sicredi e o extrato viram à meia-noite. Todos os leitores convertem a maquininha
para o mesmo dia do PDV (`pdv_pix.dia_movimento`: datetime − 10 h), então o cruzamento, as
planilhas de `saida/` e o painel usam um dia só. A coluna "Data calendário" guarda a data
original. `portal_sicredi.ler_portais` junta os Históricos de uma mesma loja (semanas diferentes,
exports sobrepostos) antes de cruzar. Exemplo: BIGGS 13, 09/10, PIX = R$ 2.440,69 nos dois lados
(40 cupons: 35 do dia, 2 do delivery fechados às 00:48 e 3 da madrugada).

## Virada de dia

Com o dia único, só sobra "virada" o cupom fechado no PDV depois das 10h do dia seguinte ao
pagamento na máquina (raro). O painel marca o Δ desses dias como "virada de dia ✓", lista o cupom nos dois dias
(com a data do outro lado em laranja) e cria uma pendência informativa "CONCILIADO · VIRADA DE
DIA" por cupom. Não é divergência e não entra no em aberto. Exemplo: BIGGS 13, 09/10, PIX
−R$ 95,08 = cupons 11328 e 11330 do delivery, pagos 09/10 à noite e fechados 10/10 00:48.

## Fechamento de dia conferido no caixa

Quando a diferença entre PDV e máquina apareceu na contagem física do caixa (sobra ou falta),
o dia é fechado no painel pelo botão "Fechar dia…" na coluna "Fechamento" da tabela de dias (o botão
existe em todo dia com relatório da máquina, para a conferência diária, inclusive nos que já batem), com o
valor conferido no caixa, justificativa, autor e hora. Fica na coleção `fechamentos` do artefato
(id `<loja>_<pix|credito|debito>_<data>`). O dia e as pendências daquela data saem do em aberto, a
loja fecha, e o registro aparece ao abrir o dia (com botão "Reabrir dia"). Primeiro caso: MAPOLI,
PIX, 08/10, R$ 6,00 a mais na máquina que sobrou no caixa físico.

## Painel

```bash
python3 scripts/gera_painel.py entrada/<pdv_cupons>.xlsx --de 2026-10-08 --ate 2026-10-14
```

Lê as planilhas de `saida/` e grava `painel/index.html` (template em `painel/template.html`,
cores e tipografia do Manual de Identidade Visual · Grupo Ragga). Publicado em
https://claude.ai/artifact/8hmLnVrM6CiE3omwzP7oms — republicar a cada rodada.

## Taxa de serviço sem o relatório por forma de pagamento

`scripts/isaias_taxa.py --loja "ISAIAS 01 - MARINGA" --de 2026-10-01 --ate 2026-10-20` cruza o relatório de
cupons (totais + operador, `entrada/pdv_totais_*.xlsx`) e o analítico (`entrada/pdv_analitico_*.xlsx`) com o
Histórico de vendas da Sicredi, sem depender da forma de pagamento: casa valor exato (até 36 h, comanda virada)
e cupom × 1,10 = máquina (até 2 h), o par mais próximo no tempo primeiro. Gera `saida/isaias_taxa_servico_<de>_a_<ate>.xlsx`
com a prova par a par (terminal exclusivo do caixa, minutos até o fechamento, concorrentes livres).
