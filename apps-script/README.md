# Controle Financeiro — Apps Script

Script único da planilha **Controle Financeiro**, em `controle-financeiro.gs`.
Substitui o script antigo do "Cartão Rafa".

## Antes de usar

1. Abra a planilha Controle Financeiro → **Extensões → Apps Script**.
2. Apague o conteúdo do arquivo antigo e cole o conteúdo de `controle-financeiro.gs`.
3. No topo do script, preencha `PLANILHA_NOTA_FISCAL_ID` com o ID da planilha do
   gerente (fica na URL, entre `/d/` e `/edit`).
4. Salve, recarregue a planilha e use o menu **Financeiro**.

A aba **MODELO** precisa existir nesta planilha: cabeçalho na linha 1 e a linha 2
formatada com todos os chips, validações, cores e formatos de número. Ela é a
única fonte de formatação do script.

## Menu

| Item | O que faz |
| --- | --- |
| 1) Importar lançamentos da Nota Fiscal | Lê a aba `Nota Fiscal` da planilha do gerente e grava cada lançamento na aba do mês do vencimento. `PARCELAS > 1` repete o lançamento nos meses seguintes. |
| 2) Copiar FIXOS para o próximo mês | Na aba do mês aberta, copia as linhas `TIPO = FIXO` para o mês seguinte, com o mesmo dia de vencimento e STATUS em branco. |
| 3) Remover duplicados (aba ativa) | Apaga as linhas repetidas da aba aberta. |
| 4) Remover duplicados (todas as abas de mês) | O mesmo, em todas as abas de mês. |
| 5) Reformatar aba ativa pelo MODELO | Reaplica a formatação do MODELO, preservando os valores. |
| 6) Reformatar TODAS as abas de mês | O mesmo, em todas as abas de mês. |

## Como o script decide que duas linhas são o mesmo lançamento

Pela combinação de **TIPO + FORNECEDOR + NF + MOTIVO + SUBMOTIVO + VALOR +
VALOR TOTAL + VENCIMENTO**. Textos são comparados sem acento, sem espaço extra e
sem diferença de maiúscula; valores como `R$ 1.304,94` e `1304.94` contam como
iguais; datas comparam só o dia, sem hora.

Uma mesma NF rateada em COMIDA, BEBIDA e DESCARTÁVEIS continua sendo três
lançamentos diferentes — só cai como duplicata o que for repetição de verdade.

Quando há repetição, fica a primeira linha. Se alguma das cópias já tiver o
STATUS preenchido (ex: PAGO) e a primeira não, a que fica é a preenchida, para
não perder a baixa.

## Cuidados

- O script só escreve nas colunas **A:K**. De **L** em diante é o dashboard e
  nunca é tocado.
- A remoção de duplicados **não usa `deleteRow`**: ela reescreve o bloco A:K
  compactado, então o dashboard em L:AF não se desloca.
- A última linha de dados é calculada pela coluna **FORNECEDOR**, não por
  `getLastRow()` — era por isso que lançamentos novos podiam cair depois da
  linha 100 (altura do dashboard) e escapar da checagem de duplicidade.
- Na planilha secundária (`DESTINO/BAR = CACO`) a formatação é reaplicada
  propriedade por propriedade a partir do mesmo MODELO, porque o `copyTo` não
  funciona entre arquivos. O resultado visual é o mesmo, exceto bordas, que a
  API do Sheets não deixa ler.
- Faça uma cópia da planilha antes de rodar a remoção de duplicados pela
  primeira vez.

## O que saiu do script antigo

- Tudo relativo a "Cartão Rafa" (menu, distribuição e diagnósticos).
- `replicarDashboardDoMesAnterior`, `recriarTabelasDinamicas` e `copiarGraficos` —
  a aba de mês já nasce clonada do MODELO, com o dashboard junto.
- `copiarFixosDoMesAnterior` e `criarProximosDoisMesesComFixos`, que viraram um
  único item: *Copiar FIXOS para o próximo mês*.
