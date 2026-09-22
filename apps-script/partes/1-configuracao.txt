/* PARTE 1 de 3 — configuração, menu e importação
   Cole este conteúdo num arquivo chamado 1-configuracao.gs
   (as três partes rodam juntas, no mesmo projeto de script) */

/* =========================================================
   O CANDIÁ | CONTROLE FINANCEIRO — script único

   Planilha: "O Candiá | Controle Financeiro"
   Origem:   "Notas Fiscais", aba "NF. CANDIA"

   O que este script faz:
     1) Importa os lançamentos da planilha Notas Fiscais para
        as abas de mês desta planilha.
     2) Copia os lançamentos FIXOS da aba do mês para o mês
        seguinte.
     3) Lista e remove lançamentos duplicados.
     4) Reaplica a formatação padrão (aba Modelo desta planilha).

   Regras gerais:
     - A formatação SEMPRE vem da aba Modelo desta planilha.
     - O script só escreve nas colunas A:K. De L em diante são
       PARCELAS, BAR, TERMINA e os dashboards: nunca são tocados.
     - Nenhuma linha é apagada com deleteRow. Ao remover
       duplicados o bloco A:K é reescrito compactado, então
       nada que esteja à direita se desloca.

   Colunas desta planilha:
     A TIPO              G VALOR
     B FORMA DE PAGAMENTO H VALOR TOTAL
     C FORNECEDOR        I VENCIMENTO
     D NF                J CONTÁBIL
     E MOTIVO            K STATUS
     F SUBMOTIVO
========================================================= */


/* ---------------------------------------------------------
   CONFIGURAÇÃO
--------------------------------------------------------- */

// Planilha do gerente (Notas Fiscais) e a aba de onde vêm os boletos.
const PLANILHA_NOTA_FISCAL_ID = "12bgHKNosmQTXL8npsaNTlTHJzkZHEfkXjpeSxVuzKRg";
const ABA_NOTA_FISCAL = "NF. CANDIA";

// Quantos meses para trás a importação enxerga.
// 1 = importa do mês passado em diante e não mexe em meses fechados.
const MESES_PARA_TRAS = 1;

// Posição das colunas NA ABA DE ORIGEM ("NF. CANDIA"), que tem um
// layout diferente do daqui: a coluna A lá é DATA DA ENTREGA.
// Use 0 para coluna que não existe na origem.
const ORIGEM = {
  dataEntrega: 1,      // A
  tipo: 2,             // B
  formaPagamento: 3,   // C
  fornecedor: 4,       // D
  nf: 5,               // E
  motivo: 6,           // F
  submotivo: 7,        // G
  valor: 8,            // H
  valorTotal: 9,       // I
  vencimento: 10,      // J
  contabil: 11,        // K
  parcelas: 0          // não existe na origem hoje
};

// Molde de formatação: cabeçalho na linha 1 e linha 2 formatada
// com todos os chips, validações, cores e formatos de número.
const ABA_MODELO = "Modelo";
const LINHA_MODELO = 2;

const ABA_RELATORIO = "Duplicados (relatório)";

const MESES = [
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
];

const COL_TIPO = 1;
const COL_FORMA_PAGAMENTO = 2;
const COL_FORNECEDOR = 3;
const COL_NF = 4;
const COL_MOTIVO = 5;
const COL_SUBMOTIVO = 6;
const COL_VALOR = 7;
const COL_VALOR_TOTAL = 8;
const COL_VENCIMENTO = 9;
const COL_CONTABIL = 10;
const COL_STATUS = 11;

const PRIMEIRA_LINHA_DADOS = 2;
const ULTIMA_COLUNA_GRAVACAO = COL_STATUS; // K — daqui pra direita não se mexe

// As abas de mês anteriores a Outubro/25 têm outro layout (F é VALOR em vez
// de SUBMOTIVO, A é CATEGORIA em vez de TIPO, e por aí vai). O script só
// mexe em aba cujo cabeçalho bate com o layout de hoje nestas colunas.
const CABECALHO_ESPERADO = {
  3: "FORNECEDOR",
  4: "NF",
  6: "SUBMOTIVO",
  7: "VALOR",
  8: "VALOR TOTAL",
  9: "VENCIMENTO"
};


/* ---------------------------------------------------------
   MENU
--------------------------------------------------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Financeiro")
    .addItem("1) Importar lançamentos da Nota Fiscal", "importarNotaFiscal")
    .addItem("2) Copiar FIXOS para o próximo mês", "copiarFixosParaProximoMes")
    .addSeparator()
    .addItem("3) Listar duplicados (só relatório)", "listarDuplicados")
    .addItem("4) Remover duplicados (aba ativa)", "removerDuplicadosAbaAtiva")
    .addItem("5) Remover duplicados (todas as abas de mês)", "removerDuplicadosTodasAsAbas")
    .addSeparator()
    .addItem("6) Reformatar aba ativa pelo Modelo", "reformatarAbaAtiva")
    .addItem("7) Reformatar todas as abas de mês", "reformatarTodasAsAbasDeMes")
    .addToUi();
}


/* ---------------------------------------------------------
   1) IMPORTAR LANÇAMENTOS DA NOTA FISCAL

   Cada lançamento cai na aba do mês do VENCIMENTO.
   O CONTÁBIL vem da origem (é a competência, e nem sempre é
   o mês do vencimento); se vier vazio, usa o mês da data de
   entrega e, na falta dela, o mês do vencimento.
   O STATUS fica em branco, para ser preenchido aqui.
--------------------------------------------------------- */

function importarNotaFiscal() {
  const ssOrigem = abrirPlanilhaOrigem();
  if (!ssOrigem) return;

  const origem = buscarAba(ssOrigem, ABA_NOTA_FISCAL);
  if (!origem) {
    avisar(`Aba "${ABA_NOTA_FISCAL}" não encontrada na planilha Notas Fiscais.`);
    return;
  }

  const colunas = Math.max(
    ORIGEM.dataEntrega, ORIGEM.tipo, ORIGEM.formaPagamento, ORIGEM.fornecedor,
    ORIGEM.nf, ORIGEM.motivo, ORIGEM.submotivo, ORIGEM.valor, ORIGEM.valorTotal,
    ORIGEM.vencimento, ORIGEM.contabil, ORIGEM.parcelas
  );

  const dados = lerLinhas(origem, ORIGEM.fornecedor, colunas);
  if (!dados.length) {
    avisar(`A aba "${ABA_NOTA_FISCAL}" não tem lançamentos.`);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date();
  const limite = new Date(hoje.getFullYear(), hoje.getMonth() - MESES_PARA_TRAS, 1);

  const lotes = {};
  const puladas = [];
  let semVencimento = 0;
  let antigos = 0;

  dados.forEach(linha => {
    const vencimento = normalizarData(valorDaOrigem(linha, ORIGEM.vencimento));

    if (!vencimento) {
      semVencimento++;
      return;
    }

    if (vencimento < limite) {
      antigos++;
      return;
    }

    const parcelas = quantidadeDeParcelas(valorDaOrigem(linha, ORIGEM.parcelas));

    for (let p = 0; p < parcelas; p++) {
      const data = somarMeses(vencimento, p);
      const aba = obterOuCriarAbaMes(ss, data);

      if (!layoutCompativel(aba)) {
        if (puladas.indexOf(aba.getName()) === -1) puladas.push(aba.getName());
        continue;
      }

      if (!lotes[aba.getName()]) lotes[aba.getName()] = { aba: aba, linhas: [] };
      lotes[aba.getName()].linhas.push(montarLinhaDaOrigem(linha, data));
    }
  });

  let inseridos = 0;
  let duplicados = 0;
  const abasTocadas = [];

  Object.keys(lotes).forEach(nome => {
    const resultado = anexarSemDuplicar(lotes[nome].aba, lotes[nome].linhas);
    inseridos += resultado.inseridos;
    duplicados += resultado.duplicados;
    if (resultado.inseridos) abasTocadas.push(`${nome} (${resultado.inseridos})`);
  });

  avisar(
    `${inseridos} lançamento(s) importado(s) da aba "${ABA_NOTA_FISCAL}".\n\n` +
    (abasTocadas.length ? `Abas: ${abasTocadas.join(", ")}\n\n` : "") +
    `Já existiam aqui (ignorados): ${duplicados}\n` +
    `Ignorados sem vencimento: ${semVencimento}\n` +
    `Ignorados por serem de meses fechados: ${antigos}` +
    (puladas.length ? `\n\nAbas de layout antigo puladas: ${puladas.join(", ")}` : "")
  );
}

function abrirPlanilhaOrigem() {
  if (!PLANILHA_NOTA_FISCAL_ID) {
    avisar("Preencha PLANILHA_NOTA_FISCAL_ID no topo do script.");
    return null;
  }

  try {
    return SpreadsheetApp.openById(PLANILHA_NOTA_FISCAL_ID);
  } catch (e) {
    avisar("Não consegui abrir a planilha Notas Fiscais.\n\n" + e);
    return null;
  }
}

function montarLinhaDaOrigem(linhaOrigem, vencimento) {
  const linha = linhaEmBranco();

  linha[COL_TIPO - 1] = texto(valorDaOrigem(linhaOrigem, ORIGEM.tipo));
  linha[COL_FORMA_PAGAMENTO - 1] = texto(valorDaOrigem(linhaOrigem, ORIGEM.formaPagamento));
  linha[COL_FORNECEDOR - 1] = texto(valorDaOrigem(linhaOrigem, ORIGEM.fornecedor));
  linha[COL_NF - 1] = texto(valorDaOrigem(linhaOrigem, ORIGEM.nf));
  linha[COL_MOTIVO - 1] = texto(valorDaOrigem(linhaOrigem, ORIGEM.motivo));
  linha[COL_SUBMOTIVO - 1] = texto(valorDaOrigem(linhaOrigem, ORIGEM.submotivo));
  linha[COL_VALOR - 1] = valorDaOrigem(linhaOrigem, ORIGEM.valor);
  linha[COL_VALOR_TOTAL - 1] = valorDaOrigem(linhaOrigem, ORIGEM.valorTotal);
  linha[COL_VENCIMENTO - 1] = vencimento;
  linha[COL_CONTABIL - 1] = contabilDaOrigem(linhaOrigem, vencimento);
  linha[COL_STATUS - 1] = "";

  return linha;
}

function contabilDaOrigem(linhaOrigem, vencimento) {
  const daOrigem = normalizarTexto(valorDaOrigem(linhaOrigem, ORIGEM.contabil));
  if (daOrigem) return daOrigem;

  const entrega = normalizarData(valorDaOrigem(linhaOrigem, ORIGEM.dataEntrega));
  const base = entrega || vencimento;

  return MESES[base.getMonth()].toUpperCase();
}

function valorDaOrigem(linha, coluna) {
  if (!coluna) return "";
  const valor = linha[coluna - 1];
  return valor === undefined || valor === null ? "" : valor;
}



/* ===== FIM DA PARTE 1 DE 3 =====
   Se você não está vendo esta linha no fim do que colou,
   a colagem veio cortada: apague e cole de novo. */
