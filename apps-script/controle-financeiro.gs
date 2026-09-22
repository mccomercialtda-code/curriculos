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
const COLUNA_LETRA = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];

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


/* ---------------------------------------------------------
   2) COPIAR FIXOS PARA O PRÓXIMO MÊS

   Rode estando na aba do mês atual. Todas as linhas com
   TIPO = FIXO são copiadas para o mês seguinte, com o mesmo
   dia de vencimento, o CONTÁBIL do novo mês e o STATUS em
   branco.
--------------------------------------------------------- */

function copiarFixosParaProximoMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = ss.getActiveSheet();
  const info = parseAbaMes(abaOrigem.getName());

  if (!info) {
    avisar("Entre na aba do mês (ex: Setembro26) antes de executar.");
    return;
  }

  if (!layoutCompativel(abaOrigem)) {
    avisar(avisoLayout(abaOrigem));
    return;
  }

  const proximoMes = somarMeses(new Date(2000 + info.ano, info.mes, 1), 1);
  const abaDestino = obterOuCriarAbaMes(ss, proximoMes);

  if (!layoutCompativel(abaDestino)) {
    avisar(avisoLayout(abaDestino));
    return;
  }

  const dados = lerLinhas(abaOrigem, COL_FORNECEDOR, ULTIMA_COLUNA_GRAVACAO);
  const novas = [];
  let naoFixos = 0;

  dados.forEach(linha => {
    if (normalizarTexto(linha[COL_TIPO - 1]) !== "FIXO") {
      naoFixos++;
      return;
    }

    const vencimento = normalizarData(linha[COL_VENCIMENTO - 1]);
    const dia = vencimento ? vencimento.getDate() : 1;
    const ultimoDia = new Date(proximoMes.getFullYear(), proximoMes.getMonth() + 1, 0).getDate();
    const novaData = new Date(
      proximoMes.getFullYear(), proximoMes.getMonth(), Math.min(dia, ultimoDia)
    );

    const nova = limparTextos(linha.slice(0, ULTIMA_COLUNA_GRAVACAO));
    nova[COL_VENCIMENTO - 1] = novaData;
    nova[COL_CONTABIL - 1] = MESES[novaData.getMonth()].toUpperCase();
    nova[COL_STATUS - 1] = "";

    novas.push(nova);
  });

  // fixo usa chave própria: o valor da luz, da água ou do salário muda
  // todo mês, então o que identifica o lançamento é fornecedor +
  // submotivo + mês do vencimento
  const resultado = anexarSemDuplicar(abaDestino, novas, chaveFixo);

  avisar(
    `${resultado.inseridos} fixo(s) copiado(s) de "${abaOrigem.getName()}" ` +
    `para "${abaDestino.getName()}".\n\n` +
    `Já existiam lá (ignorados): ${resultado.duplicados}\n` +
    `Linhas ignoradas por não serem FIXO: ${naoFixos}`
  );
}


/* ---------------------------------------------------------
   3, 4 e 5) DUPLICADOS

   Duas linhas são o mesmo lançamento quando coincidem
   FORNECEDOR, NF, SUBMOTIVO, VALOR, VALOR TOTAL e VENCIMENTO.

   TIPO e MOTIVO ficam de fora de propósito: eles costumam ser
   preenchidos aqui depois da importação, e era justamente
   isso que fazia a mesma nota entrar de novo a cada rodada.

   Fica a primeira linha do grupo. As células que estiverem
   vazias nela são completadas com o que as cópias tiverem
   (é assim que um STATUS "PAGO" marcado numa cópia não se
   perde).
--------------------------------------------------------- */

function listarDuplicados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const relatorio = [[
    "ABA", "LINHA MANTIDA", "LINHA REMOVIDA", "FORNECEDOR", "NF",
    "SUBMOTIVO", "VALOR", "VENCIMENTO", "STATUS"
  ]];

  const puladas = [];

  ss.getSheets().forEach(aba => {
    if (!ehAbaMes(aba.getName())) return;

    if (!layoutCompativel(aba)) {
      puladas.push(aba.getName());
      return;
    }

    analisarDuplicados(aba).grupos.forEach(grupo => {
      grupo.copias.forEach(copia => {
        relatorio.push([
          aba.getName(),
          grupo.linhaPlanilha,
          copia.linhaPlanilha,
          copia.valores[COL_FORNECEDOR - 1],
          copia.valores[COL_NF - 1],
          copia.valores[COL_SUBMOTIVO - 1],
          copia.valores[COL_VALOR - 1],
          copia.valores[COL_VENCIMENTO - 1],
          copia.valores[COL_STATUS - 1]
        ]);
      });
    });
  });

  let aba = buscarAba(ss, ABA_RELATORIO);
  if (!aba) aba = ss.insertSheet(ABA_RELATORIO);

  aba.clear();
  aba.getRange(1, 1, relatorio.length, relatorio[0].length).setValues(relatorio);
  aba.getRange(1, 1, 1, relatorio[0].length).setFontWeight("bold");
  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, relatorio[0].length);
  ss.setActiveSheet(aba);

  avisar(
    `${relatorio.length - 1} linha(s) duplicada(s) encontrada(s).\n\n` +
    `Nada foi alterado. Confira a aba "${ABA_RELATORIO}" e depois use ` +
    `"Remover duplicados" se estiver tudo certo.` +
    (puladas.length
      ? `\n\n${puladas.length} aba(s) de layout antigo foram puladas:\n` +
        puladas.join(", ")
      : "")
  );
}

function removerDuplicadosAbaAtiva() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!ehAbaMes(aba.getName())) {
    avisar("Entre na aba do mês (ex: Setembro26) antes de executar.");
    return;
  }

  if (!layoutCompativel(aba)) {
    avisar(avisoLayout(aba));
    return;
  }

  const removidas = removerDuplicados(aba);
  avisar(`${removidas} linha(s) duplicada(s) removida(s) em "${aba.getName()}".`);
}

function removerDuplicadosTodasAsAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const resposta = ui.alert(
    "Remover duplicados de todas as abas de mês?",
    "Inclusive as abas antigas e ocultas. Rode antes o item 3 (relatório) " +
    "e faça uma cópia da planilha.",
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  let removidas = 0;
  const detalhe = [];
  const puladas = [];
  const comErro = [];

  // cada aba vai no seu try: uma aba problemática não derruba o resto
  ss.getSheets().forEach(aba => {
    if (!ehAbaMes(aba.getName())) return;

    if (!layoutCompativel(aba)) {
      puladas.push(aba.getName());
      return;
    }

    try {
      const n = removerDuplicados(aba);
      if (n) detalhe.push(`${aba.getName()}: ${n}`);
      removidas += n;
    } catch (e) {
      comErro.push(`${aba.getName()}: ${e.message || e}`);
    }
  });

  avisar(
    `${removidas} linha(s) duplicada(s) removida(s).\n\n` +
    (detalhe.length ? detalhe.join("\n") : "Nenhuma aba tinha duplicados.") +
    (puladas.length
      ? `\n\n${puladas.length} aba(s) de layout antigo foram puladas:\n` + puladas.join(", ")
      : "") +
    (comErro.length ? `\n\nAbas com erro:\n` + comErro.join("\n") : "")
  );
}

function removerDuplicados(aba) {
  const analise = analisarDuplicados(aba);
  if (!analise.removidas) return 0;

  const range = aba.getRange(
    PRIMEIRA_LINHA_DADOS, 1, analise.totalLinhas, ULTIMA_COLUNA_GRAVACAO
  );

  const bloco = analise.mantidas.slice();
  while (bloco.length < analise.totalLinhas) bloco.push(linhaEmBranco());

  escreverComFormato(aba, range, bloco);

  return analise.removidas;
}

function analisarDuplicados(aba) {
  const ultima = ultimaLinhaDados(aba, COL_FORNECEDOR);
  const vazio = { mantidas: [], grupos: [], removidas: 0, totalLinhas: 0 };

  if (ultima < PRIMEIRA_LINHA_DADOS) return vazio;

  const totalLinhas = ultima - PRIMEIRA_LINHA_DADOS + 1;
  const valores = aba
    .getRange(PRIMEIRA_LINHA_DADOS, 1, totalLinhas, ULTIMA_COLUNA_GRAVACAO)
    .getValues();

  const indicePorChave = {};
  const mantidas = [];
  const grupos = [];
  let removidas = 0;

  valores.forEach((linha, i) => {
    if (linhaVazia(linha)) return;

    const linhaPlanilha = PRIMEIRA_LINHA_DADOS + i;
    const chave = chaveLancamento(linha);

    if (!(chave in indicePorChave)) {
      indicePorChave[chave] = mantidas.length;
      mantidas.push(limparTextos(linha));
      grupos.push({ linhaPlanilha: linhaPlanilha, copias: [] });
      return;
    }

    removidas++;

    const indice = indicePorChave[chave];
    grupos[indice].copias.push({ linhaPlanilha: linhaPlanilha, valores: linha });
    mantidas[indice] = completarVazios(mantidas[indice], linha);
  });

  return {
    mantidas: mantidas,
    grupos: grupos.filter(g => g.copias.length),
    removidas: removidas,
    totalLinhas: totalLinhas
  };
}

function completarVazios(base, extra) {
  return base.map((valor, i) => {
    const vazio = String(valor === null || valor === undefined ? "" : valor).trim() === "";
    if (!vazio) return valor;

    const novo = extra[i];
    return typeof novo === "string" ? novo.trim() : (novo === undefined ? "" : novo);
  });
}

// Chave dos FIXOS: sem valor, porque conta de luz, água e salário
// mudam de valor todo mês e continuam sendo o mesmo lançamento.
function chaveFixo(linha) {
  const data = normalizarData(linha[COL_VENCIMENTO - 1]);
  const mes = data
    ? Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM")
    : "";

  return [
    normalizarTexto(linha[COL_FORNECEDOR - 1]),
    normalizarTexto(linha[COL_SUBMOTIVO - 1]),
    mes
  ].join("|");
}

function chaveLancamento(linha) {
  return [
    normalizarTexto(linha[COL_FORNECEDOR - 1]),
    normalizarTexto(linha[COL_NF - 1]),
    normalizarTexto(linha[COL_SUBMOTIVO - 1]),
    chaveNumero(linha[COL_VALOR - 1]),
    chaveNumero(linha[COL_VALOR_TOTAL - 1]),
    chaveData(linha[COL_VENCIMENTO - 1])
  ].join("|");
}


/* ---------------------------------------------------------
   6 e 7) REFORMATAR PELO MODELO
--------------------------------------------------------- */

function reformatarAbaAtiva() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!ehAbaMes(aba.getName())) {
    avisar("Entre na aba do mês (ex: Setembro26) antes de executar.");
    return;
  }

  if (!layoutCompativel(aba)) {
    avisar(avisoLayout(aba));
    return;
  }

  const resultado = reformatarAba(aba);

  avisar(
    `${resultado.linhas} linha(s) reformatada(s) em "${aba.getName()}".` +
    avisoValoresForaDaLista(resultado.fora)
  );
}

function reformatarTodasAsAbasDeMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const resposta = ui.alert(
    "Reformatar todas as abas de mês?",
    "A formatação da aba Modelo será reaplicada em A:K de todas as abas de mês. " +
    "Os valores são preservados.",
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  let linhas = 0;
  let abas = 0;
  const puladas = [];
  const comErro = [];
  const fora = {};

  ss.getSheets().forEach(aba => {
    if (!ehAbaMes(aba.getName())) return;

    if (!layoutCompativel(aba)) {
      puladas.push(aba.getName());
      return;
    }

    try {
      const resultado = reformatarAba(aba);
      linhas += resultado.linhas;
      Object.keys(resultado.fora).forEach(c => {
        fora[c] = fora[c] || {};
        resultado.fora[c].forEach(v => { fora[c][v] = true; });
      });
      abas++;
    } catch (e) {
      comErro.push(`${aba.getName()}: ${e.message || e}`);
    }
  });

  const foraLista = {};
  Object.keys(fora).forEach(c => { foraLista[c] = Object.keys(fora[c]); });

  avisar(
    `${linhas} linha(s) reformatada(s) em ${abas} aba(s) de mês.` +
    avisoValoresForaDaLista(foraLista) +
    (puladas.length
      ? `\n\n${puladas.length} aba(s) de layout antigo foram puladas:\n` + puladas.join(", ")
      : "") +
    (comErro.length ? `\n\nAbas com erro:\n` + comErro.join("\n") : "")
  );
}

function reformatarAba(aba) {
  const ultima = ultimaLinhaDados(aba, COL_FORNECEDOR);
  if (ultima < PRIMEIRA_LINHA_DADOS) return { linhas: 0, fora: {} };

  const qtd = ultima - PRIMEIRA_LINHA_DADOS + 1;
  const range = aba.getRange(PRIMEIRA_LINHA_DADOS, 1, qtd, ULTIMA_COLUNA_GRAVACAO);

  // o trim() faz o texto voltar a bater com a lista do dropdown,
  // que é o que traz o chip colorido de volta
  const valores = range.getValues().map(limparTextos);

  escreverComFormato(aba, range, valores);

  return { linhas: qtd, fora: valoresForaDaLista(valores) };
}


/* ---------------------------------------------------------
   GRAVAÇÃO SEM DUPLICAR
--------------------------------------------------------- */

function anexarSemDuplicar(aba, linhas, gerarChave) {
  if (!linhas.length) return { inseridos: 0, duplicados: 0 };

  const chave = gerarChave || chaveLancamento;
  const chaves = {};
  const ultima = ultimaLinhaDados(aba, COL_FORNECEDOR);

  if (ultima >= PRIMEIRA_LINHA_DADOS) {
    aba.getRange(PRIMEIRA_LINHA_DADOS, 1, ultima - PRIMEIRA_LINHA_DADOS + 1, ULTIMA_COLUNA_GRAVACAO)
      .getValues()
      .forEach(linha => {
        if (!linhaVazia(linha)) chaves[chave(linha)] = true;
      });
  }

  const novas = [];
  let duplicados = 0;

  linhas.forEach(linha => {
    const atual = chave(linha);

    if (chaves[atual]) {
      duplicados++;
      return;
    }

    chaves[atual] = true; // evita duplicar dentro da própria rodada
    novas.push(linha);
  });

  if (novas.length) escreverBloco(aba, ultima + 1, novas);

  return { inseridos: novas.length, duplicados: duplicados };
}

function escreverBloco(aba, linhaInicial, linhas) {
  const ultimaNecessaria = linhaInicial + linhas.length - 1;

  if (ultimaNecessaria > aba.getMaxRows()) {
    aba.insertRowsAfter(aba.getMaxRows(), ultimaNecessaria - aba.getMaxRows());
  }

  const range = aba.getRange(linhaInicial, 1, linhas.length, ULTIMA_COLUNA_GRAVACAO);

  escreverComFormato(aba, range, linhas);
}


/* ---------------------------------------------------------
   FORMATAÇÃO — sempre a partir da aba Modelo desta planilha
--------------------------------------------------------- */

// Grava um bloco de valores com a formatação certa e sem esbarrar na
// validação de dados.
//
// A ordem importa por dois motivos.
//
// 1) As listas suspensas estão como "rejeitar entrada" e vários lançamentos
//    têm valores que não estão na lista ("Cartão Rafa", "Casa Caco", "PIX"
//    em maiúscula). Gravar com a regra ativa derrubava a rotina inteira com
//    "os dados inseridos na célula B2 violam a validação de dados".
//    Por isso a regra sai antes da gravação e volta depois.
//
// 2) A cor do chip faz parte da regra de validação, e só o copyTo a carrega:
//    reconstruir a regra com setDataValidations devolve o chip cinza. Então
//    as colunas que o Modelo define voltam por copyTo, a partir do Modelo.
//    As que o Modelo não define (STATUS, por exemplo) voltam como estavam.
function escreverComFormato(aba, range, valores) {
  const reserva = validacoesDeReserva(aba);

  range.clearDataValidations();
  range.setValues(valores);

  aplicarFormatoModelo(range);
  devolverValidacoes(range, reserva);
}

// Formato e listas suspensas a partir da linha 2 da aba Modelo.
function aplicarFormatoModelo(range) {
  const modelo = linhaModelo();

  // uma linha de molde copiada sobre um bloco alto se repete em todas as linhas
  modelo.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  modelo.getDataValidations()[0].forEach((regra, i) => {
    if (!regra) return;

    modelo.getSheet()
      .getRange(LINHA_MODELO, i + 1)
      .copyTo(
        colunaDoBloco(range, i + 1),
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
        false
      );
  });
}

// Regras das colunas que o Modelo não define, para não perdê-las na gravação.
// O Modelo, por exemplo, não tem lista em STATUS, e sem isso o dropdown de
// PROGRAMADO/PAGO sumiria das linhas gravadas.
function validacoesDeReserva(aba) {
  const doModelo = linhaModelo().getDataValidations()[0];
  const daAba = aba
    .getRange(PRIMEIRA_LINHA_DADOS, 1, 1, ULTIMA_COLUNA_GRAVACAO)
    .getDataValidations()[0];

  return daAba.map((regra, i) => (doModelo[i] ? null : regra));
}

function devolverValidacoes(range, reserva) {
  reserva.forEach((regra, i) => {
    if (!regra) return;

    const coluna = colunaDoBloco(range, i + 1);
    coluna.setDataValidations(repetirLinha([regra], coluna.getNumRows()));
  });
}

function colunaDoBloco(range, coluna) {
  return range.getSheet().getRange(range.getRow(), coluna, range.getNumRows(), 1);
}

// Valores gravados que a lista do Modelo não aceita. Não impedem nada
// (a regra entra depois da gravação), mas viram aviso: é o que falta
// acrescentar na linha 2 do Modelo.
function valoresForaDaLista(valores) {
  const regras = linhaModelo().getDataValidations()[0];
  const fora = {};

  regras.forEach((regra, i) => {
    if (!regra) return;
    if (regra.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return;

    const lista = (regra.getCriteriaValues()[0] || []).map(normalizarTexto);
    const achados = {};

    valores.forEach(linha => {
      const valor = String(linha[i] === null || linha[i] === undefined ? "" : linha[i]).trim();
      if (!valor) return;
      if (lista.indexOf(normalizarTexto(valor)) === -1) achados[valor] = true;
    });

    const chaves = Object.keys(achados);
    if (chaves.length) fora[COLUNA_LETRA[i]] = chaves;
  });

  return fora;
}

function avisoValoresForaDaLista(fora) {
  const colunas = Object.keys(fora);
  if (!colunas.length) return "";

  return (
    "\n\nValores que não estão nas listas do Modelo (o chip fica sem cor até " +
    "você acrescentá-los na linha 2 do Modelo):\n" +
    colunas.map(c => `${c}: ${fora[c].join(", ")}`).join("\n")
  );
}

function linhaModelo() {
  const modelo = buscarAba(SpreadsheetApp.getActiveSpreadsheet(), ABA_MODELO);

  if (!modelo) {
    throw new Error(
      `Aba "${ABA_MODELO}" não encontrada. Ela é o molde de formatação: ` +
      `cabeçalho na linha 1 e linha 2 formatada.`
    );
  }

  return modelo.getRange(LINHA_MODELO, 1, 1, ULTIMA_COLUNA_GRAVACAO);
}

function repetirLinha(linha, quantidade) {
  const bloco = [];
  for (let i = 0; i < quantidade; i++) bloco.push(linha.slice());
  return bloco;
}


/* ---------------------------------------------------------
   ABAS DE MÊS

   Os nomes são reconhecidos em qualquer formato: "Setembro26",
   "Setembro/26", "Setembro 26", com ou sem espaço sobrando.
   Ao criar uma aba nova, o script copia o padrão de nome que
   as abas já existentes usam.
--------------------------------------------------------- */

function obterOuCriarAbaMes(ss, data) {
  const existente = buscarAbaMes(ss, data);
  if (existente) return existente;

  const modelo = buscarAba(ss, ABA_MODELO);

  if (!modelo) {
    throw new Error(`Aba "${ABA_MODELO}" não encontrada: não dá para criar o mês novo.`);
  }

  const nova = modelo.copyTo(ss).setName(nomeAbaPorData(ss, data));
  const maxLinhas = nova.getMaxRows();

  if (maxLinhas >= PRIMEIRA_LINHA_DADOS) {
    nova.getRange(
      PRIMEIRA_LINHA_DADOS, 1,
      maxLinhas - PRIMEIRA_LINHA_DADOS + 1,
      ULTIMA_COLUNA_GRAVACAO
    ).clearContent();
  }

  nova.setFrozenRows(1);
  return nova;
}

function buscarAbaMes(ss, data) {
  const mes = data.getMonth();
  const ano = data.getFullYear() % 100;

  for (const aba of ss.getSheets()) {
    const info = parseAbaMes(aba.getName());
    if (info && info.mes === mes && info.ano === ano) return aba;
  }

  return null;
}

function nomeAbaPorData(ss, data) {
  let separador = "";

  for (const aba of ss.getSheets()) {
    const info = parseAbaMes(aba.getName());
    if (info) { separador = info.separador; break; }
  }

  return MESES[data.getMonth()] + separador + String(data.getFullYear()).slice(-2);
}

function ehAbaMes(nome) {
  return parseAbaMes(nome) !== null;
}


function layoutCompativel(aba) {
  if (aba.getLastColumn() < ULTIMA_COLUNA_GRAVACAO) return false;

  const cabecalho = aba.getRange(1, 1, 1, ULTIMA_COLUNA_GRAVACAO).getValues()[0];

  for (const coluna in CABECALHO_ESPERADO) {
    if (normalizarTexto(cabecalho[coluna - 1]) !== CABECALHO_ESPERADO[coluna]) {
      return false;
    }
  }

  return true;
}

function parseAbaMes(nome) {
  const limpo = String(nome || "").replace(/ /g, " ").trim();
  const partes = limpo.match(/^([^\d]+?)\s*([\/\-\s]?)\s*(\d{2}|\d{4})$/);

  if (!partes) return null;

  const mes = MESES.findIndex(m => normalizarTexto(m) === normalizarTexto(partes[1]));
  if (mes === -1) return null;

  const ano = Number(partes[3]);
  if (isNaN(ano)) return null;

  return { mes: mes, ano: ano % 100, separador: partes[2] || "" };
}

// Última linha preenchida numa coluna de referência.
// Não usa getLastRow() porque os dashboards à direita esticam
// o resultado e fariam o lançamento novo cair lá embaixo,
// fora do alcance da checagem de duplicidade.
function ultimaLinhaDados(aba, colunaReferencia) {
  const max = aba.getMaxRows();
  if (max < PRIMEIRA_LINHA_DADOS) return PRIMEIRA_LINHA_DADOS - 1;

  const coluna = aba
    .getRange(PRIMEIRA_LINHA_DADOS, colunaReferencia, max - PRIMEIRA_LINHA_DADOS + 1, 1)
    .getValues();

  for (let i = coluna.length - 1; i >= 0; i--) {
    if (String(coluna[i][0]).trim() !== "") return PRIMEIRA_LINHA_DADOS + i;
  }

  return PRIMEIRA_LINHA_DADOS - 1;
}

function lerLinhas(aba, colunaReferencia, colunas) {
  const ultima = ultimaLinhaDados(aba, colunaReferencia);
  if (ultima < PRIMEIRA_LINHA_DADOS) return [];

  const disponiveis = Math.max(Math.min(aba.getLastColumn(), colunas), 1);

  return aba
    .getRange(PRIMEIRA_LINHA_DADOS, 1, ultima - PRIMEIRA_LINHA_DADOS + 1, disponiveis)
    .getValues()
    .map(linha => {
      const completa = linha.slice(0, colunas);
      while (completa.length < colunas) completa.push("");
      return completa;
    });
}


/* ---------------------------------------------------------
   AUXILIARES
--------------------------------------------------------- */

function avisoLayout(aba) {
  return (
    `A aba "${aba.getName()}" está no layout antigo (a coluna F não é ` +
    `SUBMOTIVO). O script não mexe nessas abas para não misturar as colunas. ` +
    `Use só nas abas de Outubro/25 em diante.`
  );
}

function avisar(mensagem) {
  try {
    SpreadsheetApp.getUi().alert(mensagem);
  } catch (e) {
    Logger.log(mensagem);
  }
}

function buscarAba(ss, nome) {
  const alvo = normalizarNomeAba(nome);

  for (const aba of ss.getSheets()) {
    if (normalizarNomeAba(aba.getName()) === alvo) return aba;
  }

  return null;
}

function linhaEmBranco() {
  const linha = [];
  for (let c = 0; c < ULTIMA_COLUNA_GRAVACAO; c++) linha.push("");
  return linha;
}

function linhaVazia(linha) {
  return linha.every(v => String(v === null || v === undefined ? "" : v).trim() === "");
}

function limparTextos(linha) {
  return linha.map(v => (typeof v === "string" ? v.trim() : (v === undefined ? "" : v)));
}

function texto(valor) {
  return typeof valor === "string" ? valor.trim() : (valor === undefined || valor === null ? "" : valor);
}

function quantidadeDeParcelas(valor) {
  const numero = Math.floor(Number(valor));
  if (!numero || isNaN(numero) || numero < 1) return 1;
  return Math.min(numero, 60);
}

function somarMeses(data, quantidade) {
  const ano = data.getFullYear();
  const mes = data.getMonth() + quantidade;
  const ultimoDia = new Date(ano, mes + 1, 0).getDate();

  return new Date(ano, mes, Math.min(data.getDate(), ultimoDia));
}

function chaveNumero(valor) {
  if (valor === "" || valor === null || valor === undefined) return "";
  if (typeof valor === "number") return valor.toFixed(2);

  let t = String(valor).replace(/[R$\s ]/g, "");
  if (t.indexOf(",") >= 0) t = t.replace(/\./g, "").replace(",", ".");

  const numero = Number(t);
  return isNaN(numero) ? normalizarTexto(valor) : numero.toFixed(2);
}

function chaveData(valor) {
  const data = normalizarData(valor);
  if (!data) return normalizarTexto(valor);

  return Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function normalizarNomeAba(nome) {
  return String(nome || "").replace(/ /g, " ").replace(/\s+/g, " ").trim().toLowerCase();
}

function normalizarTexto(texto) {
  return String(texto === null || texto === undefined ? "" : texto)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/ /g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function normalizarData(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }

  if (typeof valor === "number" && valor > 0) {
    // número de série do Sheets (30/12/1899 = 0)
    const base = new Date(1899, 11, 30);
    return new Date(base.getFullYear(), base.getMonth(), base.getDate() + Math.floor(valor));
  }

  if (typeof valor === "string") {
    const partes = valor.trim().split("/");

    if (partes.length === 3) {
      const dia = Number(partes[0]);
      const mes = Number(partes[1]) - 1;
      let ano = Number(partes[2]);

      if (isNaN(dia) || isNaN(mes) || isNaN(ano)) return null;
      if (ano < 100) ano += 2000;

      return new Date(ano, mes, dia);
    }
  }

  return null;
}
