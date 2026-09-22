# Controle Financeiro — Apps Script

Script único da planilha **O Candiá | Controle Financeiro**, em
`controle-financeiro.gs`. Substitui o script antigo do "Cartão Rafa".

Planilhas envolvidas:

| Papel | Planilha | Aba |
| --- | --- | --- |
| Destino | O Candiá \| Controle Financeiro (`1Bka7IT…`) | abas de mês + `Modelo` |
| Origem | Notas Fiscais (`12bgHKN…`) | `NF. CANDIA` |

## Instalação

1. Planilha Controle Financeiro → **Extensões → Apps Script**.
2. Apague o conteúdo do arquivo antigo e cole o de `controle-financeiro.gs`.
3. Salve, recarregue a planilha e use o menu **Financeiro**.
   Na primeira execução o Google pede autorização (é o mesmo script de sempre,
   só que agora ele também abre a planilha Notas Fiscais).

Os dois IDs já estão preenchidos no topo do arquivo.

## Menu

| Item | O que faz |
| --- | --- |
| 1) Importar lançamentos da Nota Fiscal | Lê a aba `NF. CANDIA` e grava cada lançamento na aba do mês do **vencimento**. |
| 2) Copiar FIXOS para o próximo mês | Na aba do mês aberta, copia as linhas `TIPO = FIXO` para o mês seguinte. |
| 3) Listar duplicados (só relatório) | Cria a aba `Duplicados (relatório)` mostrando o que seria removido. **Não altera nada.** |
| 4) Remover duplicados (aba ativa) | Limpa as repetições da aba aberta. |
| 5) Remover duplicados (todas as abas de mês) | O mesmo, em todas as abas de mês — inclusive as antigas e as ocultas. |
| 6) Reformatar aba ativa pelo Modelo | Reaplica a formatação da aba `Modelo` em A:K, preservando os valores. |
| 7) Reformatar todas as abas de mês | O mesmo, em todas as abas de mês. |

## De onde vem cada coluna na importação

A aba `NF. CANDIA` tem um layout diferente: a coluna A lá é `DATA DA ENTREGA`,
então tudo fica deslocado uma coluna. O mapeamento está na constante `ORIGEM`,
no topo do script:

| Controle Financeiro | NF. CANDIA |
| --- | --- |
| A TIPO | B TIPO |
| B FORMA DE PAGAMENTO | C FORMA DE PAGAMENTO |
| C FORNECEDOR | D FORNECEDOR |
| D NF | E NF |
| E MOTIVO | F MOTIVO |
| F SUBMOTIVO | G SUBMOTIVO |
| G VALOR | H VALOR |
| H VALOR TOTAL | I VALOR TOTAL |
| I VENCIMENTO | J VENCIMENTO (define a aba de destino) |
| J CONTÁBIL | K CONTÁBIL (é a competência, não o mês do vencimento — vem como está) |
| K STATUS | fica em branco, para ser preenchido aqui |

`MESES_PARA_TRAS = 1` faz a importação olhar do mês passado em diante, para não
mexer em meses já fechados. Aumente se precisar puxar mais história.

## Como o script decide que duas linhas são o mesmo lançamento

**FORNECEDOR + NF + SUBMOTIVO + VALOR + VALOR TOTAL + VENCIMENTO.**

`TIPO` e `MOTIVO` ficam de fora de propósito: eles são preenchidos aqui depois
da importação, e uma chave que os incluísse deixava de reconhecer o lançamento
na rodada seguinte — era por isso que a mesma nota voltava a entrar.

Textos são comparados sem acento, sem espaço sobrando e sem diferença de
maiúscula; `R$ 1.304,94` e `1304.94` contam como iguais; datas comparam só o
dia. Uma mesma NF rateada em COMIDA, BEBIDA e DESCARTÁVEIS continua sendo três
lançamentos distintos.

Quando há repetição fica a primeira linha, e as células vazias dela são
completadas com o que as cópias tiverem — é assim que um `PAGO` marcado numa
cópia não se perde.

**Fixos usam outra chave**: `FORNECEDOR + SUBMOTIVO + mês do vencimento`, sem o
valor. Conta de luz, água e salário mudam de valor todo mês e continuam sendo o
mesmo lançamento; com a chave comum, o item 2 criaria uma segunda cópia de cada
um deles.

## Só as abas do layout de hoje

As abas de mês anteriores a **Outubro/25** têm outro layout: em várias delas a
coluna F é `VALOR` e não `SUBMOTIVO`, e em algumas a coluna A é `CATEGORIA` ou
até `FORNECEDOR`. Comparar duplicados nessas abas usando as colunas de hoje daria
resultado errado, então o script confere o cabeçalho (colunas C, D, F, G, H e I)
e **pula** quem não bate, dizendo quais abas pulou.

São 26 abas no layout atual, de Outubro/25 a Julho/27.

## Validação de dados e a cor dos chips

**A cor do chip faz parte da regra de validação**, não do formato da célula.
E só o `copyTo` carrega essa cor: uma regra lida com `getDataValidations` e
regravada com `setDataValidations` volta como chip cinza. Por isso as colunas
que o `Modelo` define (TIPO, FORMA DE PAGAMENTO, MOTIVO, SUBMOTIVO) têm a
validação copiada do `Modelo` com `copyTo`, coluna por coluna.

As colunas que o `Modelo` **não** define voltam como estavam na aba — é o caso
de STATUS, que no `Modelo` não tem lista nenhuma: sem essa reserva o dropdown
de `PROGRAMADO/PAGO` sumiria das linhas gravadas. Essas voltam sem a cor do
chip; para elas ficarem coloridas também, basta criar a regra na linha 2 do
`Modelo`.

As listas suspensas são "rejeitar entrada" e há lançamentos com valores que não
estão nelas (`Cartão Rafa`, `Casa Caco`, `PIX` em maiúscula, `1` e `0` em
STATUS). Gravar por cima com a regra ativa derruba a rotina com *"Os dados
inseridos na célula B2 violam o respectivo conjunto de regras de validação de
dados"*. Então a gravação segue esta ordem:

1. tira a validação da faixa;
2. grava os valores;
3. copia o formato do `Modelo`;
4. devolve a validação (do `Modelo` por `copyTo`, ou a da aba onde o `Modelo`
   não define).

Nenhum valor é alterado nem apagado nesse processo. Ao reformatar, o aviso do
fim lista os valores que não estão nas listas do `Modelo` — é o que falta
acrescentar lá para o chip ficar colorido. Hoje são, entre outros,
`Cartão Rafa`, `Casa Caco`, `Cartão Carol`, `Carão Rafa` (com o erro de
digitação) e `Dinheiro/PIX` na forma de pagamento; `MÚSICA` no motivo; e
`RETIRADA DE LUCRO`, `PULSEIRA COUVERT` e `RESCISÃO CONTRATUAL` no submotivo.

## Cuidados

- O script só escreve nas colunas **A:K**. `PARCELAS`, `BAR`, `TERMINA` e os
  dashboards à direita nunca são tocados.
- A remoção de duplicados **não usa `deleteRow`**: ela reescreve o bloco A:K
  compactado, então nada que esteja à direita se desloca.
- A última linha de dados é calculada pela coluna **FORNECEDOR**, não por
  `getLastRow()` — os dashboards esticavam o `getLastRow()` e faziam o
  lançamento novo cair lá embaixo, fora do alcance da checagem de duplicidade.
- Nomes de aba são reconhecidos em qualquer formato (`Setembro26`,
  `Setembro/26`, `Setembro 26`, com espaço sobrando). Ao criar um mês novo, o
  script copia o padrão que as abas existentes já usam.
- Rode o item 3 e confira o relatório antes de usar o item 5. Faça uma cópia da
  planilha antes da primeira limpeza.
- Nos itens que varrem todas as abas, cada aba roda no seu próprio `try`: uma
  aba com problema não derruba mais as outras, e o aviso do fim diz quais deram
  erro.

## O que saiu do script antigo

- Tudo relativo a "Cartão Rafa": menu, distribuição por parcelas e diagnósticos.
- `replicarDashboardDoMesAnterior`, `recriarTabelasDinamicas` e `copiarGraficos`.
  A aba de mês nova nasce clonada da aba `Modelo`.
- `copiarFixosDoMesAnterior` e `criarProximosDoisMesesComFixos`, que viraram um
  único item: *Copiar FIXOS para o próximo mês*.
- O desvio para a planilha do Caco (`DESTINO/BAR = CACO`). A aba `NF. CANDIA`
  não tem coluna de bar, então o desvio nunca teria o que fazer. Se voltar a
  ser necessário, é a coluna que precisa existir na origem primeiro.

A aba `PAES CANDIA` da planilha Notas Fiscais não é importada: ela não tem
coluna de vencimento preenchida, e é o vencimento que define a aba de destino.

## Se a colagem der "SyntaxError: Unexpected end of input"

Esse erro quer dizer que o código chegou cortado no Apps Script, não que ele
tem defeito. O arquivo inteiro tem **865 linhas** e termina com `return null;`
seguido de `}`.

Para colar com segurança, use a pasta `partes/`, que tem o mesmo script dividido
em três arquivos menores. No Apps Script, crie três arquivos (o **+** ao lado de
"Arquivos" → Script) com os nomes abaixo e cole um conteúdo em cada:

| Arquivo no Apps Script | Colar | Linhas |
| --- | --- | --- |
| `1-configuracao.gs` | `partes/1-configuracao.txt` | 251 |
| `2-rotinas.gs` | `partes/2-rotinas.txt` | 327 |
| `3-auxiliares.gs` | `partes/3-auxiliares.txt` | 309 |

Os três ficam no mesmo projeto e funcionam como um script só. Cada parte termina
com um comentário `FIM DA PARTE N DE 3`: se ele não aparecer no fim do que você
colou, a colagem veio cortada de novo — apague o conteúdo do arquivo e repita.

Apague o arquivo `Código.gs` antigo (ou deixe-o vazio) para o script velho do
Cartão Rafa não rodar junto.
