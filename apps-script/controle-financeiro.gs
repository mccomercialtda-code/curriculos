/* =========================================================
   CONTROLE FINANCEIRO — script único e simplificado

   O que este script faz:
     1) Importa os lançamentos da planilha do gerente ("Nota Fiscal")
        para as abas de mês desta planilha (Controle Financeiro).
     2) Copia os lançamentos FIXOS da aba do mês para o mês seguinte.
     3) Remove lançamentos duplicados.
     4) Reaplica a formatação padrão (aba MODELO desta planilha).

   Regras gerais:
     - A formatação SEMPRE vem da aba MODELO desta planilha
       (Controle Financeiro), inclusive quando os dados são
       gravados em outra planilha.
     - O script só escreve nas colunas A:K. De L em diante é
       território do dashboard e nunca é tocado.
     - Nenhuma linha é apagada com deleteRow: os dados são
       reescritos no bloco A:K, para não desalinhar o dashboard.

   Estrutura das colunas:
     A TIPO              H VALOR TOTAL
     B FORMA DE PAGAMENTO I VENCIMENTO
     C FORNECEDOR        J CONTÁBIL
     D NF                K STATUS
     E MOTIVO            L PARCELAS
     F SUBMOTIVO         M DESTINO/BAR
     G VALOR             N TERMINA / OBS
========================================================= */


/* ---------------------------------------------------------
   CONFIGURAÇÃO
--------------------------------------------------------- */

// Planilha do gerente, de onde vêm os boletos.
// O ID fica na URL da planilha, entre "/d/" e "/edit".
const PLANILHA_NOTA_FISCAL_ID = "COLE_AQUI_O_ID_DA_PLANILHA_NOTA_FISCAL";
const ABA_NOTA_FISCAL = "Nota Fiscal";

// Planilha secundária de destino (linhas com DESTINO/BAR = CACO).
// Deixe "" se todos os lançamentos ficam nesta planilha.
const PLANILHA_SECUNDARIA_ID = "1a9OmlhdVj29e6dFX3krjxerw-Ou0q7u_WAD8I18QNVA";
const DESTINO_SECUNDARIO = "CACO";

// Molde de formatação: cabeçalho na linha 1, linha 2 formatada
// com todos os chips, validações, cores e formatos de número.
const ABA_MODELO = "MODELO";
const LINHA_MODELO = 2;

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
const COL_PARCELAS = 12;
const COL_DESTINO = 13;
const COL_TERMINA = 14;

const PRIMEIRA_LINHA_DADOS = 2;
const ULTIMA_COLUNA_GRAVACAO = COL_STATUS; // K — daqui pra frente é dashboard


/* ---------------------------------------------------------
   MENU
--------------------------------------------------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Financeiro")
    .addItem("1) Importar lançamentos da Nota Fiscal", "importarNotaFiscal")
    .addItem("2) Copiar FIXOS para o próximo mês", "copiarFixosParaProximoMes")
    .addSeparator()
    .addItem("3) Remover duplicados (aba ativa)", "removerDuplicadosAbaAtiva")
    .addItem("4) Remover duplicados (todas as abas de mês)", "removerDuplicadosTodasAsAbas")
    .addSeparator()
    .addItem("5) Reformatar aba ativa pelo MODELO", "reformatarAbaAtiva")
    .addItem("6) Reformatar todas as abas de mês", "reformatarTodasAsAbasDeMes")
    .addToUi();
}


/* ---------------------------------------------------------
   1) IMPORTAR LANÇAMENTOS DA NOTA FISCAL

   Lê a aba "Nota Fiscal" da planilha do gerente e joga cada
   lançamento na aba do mês do vencimento. Se a coluna
   PARCELAS tiver um número maior que 1, o lançamento é
   repetido nos meses seguintes (uma parcela por mês).
--------------------------------------------------------- */

function importarNotaFiscal() {
  if (!PLANILHA_NOTA_FISCAL_ID || PLANILHA_NOTA_FISCAL_ID.indexOf("COLE_AQUI") === 0) {
    avisar(
      "Configure a constante PLANILHA_NOTA_FISCAL_ID no topo do script " +
      "com o ID da planilha do gerente."
    );
    return;
  }

  let ssOrigem;
  try {
    ssOrigem = SpreadsheetApp.openById(PLANILHA_NOTA_FISCAL_ID);
  } catch (e) {
    avisar("Não consegui abrir a planilha Nota Fiscal. Confira o ID e o compartilhamento.\n\n" + e);
    return;
  }

  const origem = buscarAba(ssOrigem, ABA_NOTA_FISCAL);
  if (!origem) {
    avisar(`Aba "${ABA_NOTA_FISCAL}" não encontrada na planilha do gerente.`);
    return;
  }

  const dados = lerLinhas(origem);
  if (!dados.length) {
    avisar(`A aba "${ABA_NOTA_FISCAL}" não tem lançamentos.`);
    return;
  }

  const lotes = {};
  let semFornecedor = 0;
  let semVencimento = 0;

  dados.forEach(linha => {
    const fornecedor = String(linha[COL_FORNECEDOR - 1] || "").trim();
    if (!fornecedor) {
      semFornecedor++;
      return;
    }

    const vencimento = normalizarData(linha[COL_VENCIMENTO - 1]);
    if (!vencimento) {
      semVencimento++;
      return;
    }

    const parcelas = quantidadeDeParcelas(linha[COL_PARCELAS - 1]);
    const ssDestino = planilhaDestino(linha);

    for (let p = 0; p < parcelas; p++) {
      const data = somarMeses(vencimento, p);
      const aba = obterOuCriarAbaMes(ssDestino, data);
      const chave = ssDestino.getId() + "|" + aba.getName();

      if (!lotes[chave]) lotes[chave] = { aba: aba, linhas: [] };
      lotes[chave].linhas.push(montarLinha(linha, data));
    }
  });

  let inseridos = 0;
  let duplicados = 0;

  Object.keys(lotes).forEach(chave => {
    const resultado = anexarSemDuplicar(lotes[chave].aba, lotes[chave].linhas);
    inseridos += resultado.inseridos;
    duplicados += resultado.duplicados;
  });

  avisar(
    `${inseridos} lançamento(s) importado(s) da Nota Fiscal.\n\n` +
    `Já existiam (duplicados ignorados): ${duplicados}\n` +
    `Ignorados sem fornecedor: ${semFornecedor}\n` +
    `Ignorados sem vencimento: ${semVencimento}`
  );
}


/* ---------------------------------------------------------
   2) COPIAR FIXOS PARA O PRÓXIMO MÊS

   Rode estando na aba do mês atual (ex: Setembro/26).
   Todas as linhas com TIPO = FIXO são copiadas para a aba
   do mês seguinte, com o mesmo dia de vencimento, o mês
   contábil ajustado e o STATUS em branco.
--------------------------------------------------------- */

function copiarFixosParaProximoMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = ss.getActiveSheet();
  const nomeOrigem = abaOrigem.getName();

  if (!ehAbaMes(nomeOrigem)) {
    avisar("Entre em uma aba de mês (ex: Setembro/26) antes de executar.");
    return;
  }

  const nomeDestino = somarMesesEmNomeAba(nomeOrigem, 1);
  const infoDestino = parseAbaMes(nomeDestino);
  const abaDestino = obterOuCriarAbaMes(ss, new Date(2000 + infoDestino.ano, infoDestino.mes, 1));

  const dados = lerLinhas(abaOrigem);
  const novas = [];
  let naoFixos = 0;

  dados.forEach(linha => {
    if (normalizarTexto(linha[COL_TIPO - 1]) !== "FIXO") {
      naoFixos++;
      return;
    }

    const fornecedor = String(linha[COL_FORNECEDOR - 1] || "").trim();
    if (!fornecedor) return;

    const vencimento = normalizarData(linha[COL_VENCIMENTO - 1]);
    const dia = vencimento ? vencimento.getDate() : 1;
    const ultimoDia = new Date(2000 + infoDestino.ano, infoDestino.mes + 1, 0).getDate();
    const novaData = new Date(2000 + infoDestino.ano, infoDestino.mes, Math.min(dia, ultimoDia));

    novas.push(montarLinha(linha, novaData));
  });

  const resultado = anexarSemDuplicar(abaDestino, novas);

  avisar(
    `${resultado.inseridos} fixo(s) copiado(s) de "${nomeOrigem}" para "${nomeDestino}".\n\n` +
    `Já existiam (duplicados ignorados): ${resultado.duplicados}\n` +
    `Linhas ignoradas por não serem FIXO: ${naoFixos}`
  );
}


/* ---------------------------------------------------------
   3 e 4) REMOVER DUPLICADOS

   Duas linhas são consideradas o mesmo lançamento quando
   coincidem TIPO, FORNECEDOR, NF, MOTIVO, SUBMOTIVO, VALOR,
   VALOR TOTAL e VENCIMENTO (ver chaveLancamento).

   Fica a primeira ocorrência. Se uma das cópias já tiver o
   STATUS preenchido (ex: PAGO) e a primeira não, a que fica
   é a preenchida — assim nenhuma baixa se perde.

   As linhas não são apagadas com deleteRow: o bloco A:K é
   reescrito compactado, então o dashboard em L:AF continua
   exatamente onde estava.
--------------------------------------------------------- */

function removerDuplicadosAbaAtiva() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!ehAbaMes(aba.getName())) {
    avisar("Entre em uma aba de mês (ex: Setembro/26) antes de executar.");
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
    "As linhas repetidas serão apagadas e os lançamentos restantes " +
    "reorganizados nas colunas A:K. O dashboard não é tocado. " +
    "Faça uma cópia da planilha antes se quiser garantia.",
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  let removidasTotal = 0;
  let abas = 0;

  ss.getSheets().forEach(aba => {
    if (!ehAbaMes(aba.getName())) return;
    removidasTotal += removerDuplicados(aba);
    abas++;
  });

  avisar(`${removidasTotal} linha(s) duplicada(s) removida(s) em ${abas} aba(s) de mês.`);
}

function removerDuplicados(aba) {
  const ultima = ultimaLinhaDados(aba);
  if (ultima < PRIMEIRA_LINHA_DADOS) return 0;

  const qtd = ultima - PRIMEIRA_LINHA_DADOS + 1;
  const range = aba.getRange(PRIMEIRA_LINHA_DADOS, 1, qtd, ULTIMA_COLUNA_GRAVACAO);
  const valores = range.getValues();

  const posicaoPorChave = {};
  const mantidas = [];
  let removidas = 0;

  valores.forEach(linha => {
    if (linhaVazia(linha)) return;

    const chave = chaveLancamento(linha);

    if (!(chave in posicaoPorChave)) {
      posicaoPorChave[chave] = mantidas.length;
      mantidas.push(limparTextos(linha));
      return;
    }

    removidas++;

    const indice = posicaoPorChave[chave];
    const statusMantido = String(mantidas[indice][COL_STATUS - 1] || "").trim();
    const statusDuplicado = String(linha[COL_STATUS - 1] || "").trim();

    if (!statusMantido && statusDuplicado) {
      mantidas[indice] = limparTextos(linha);
    }
  });

  if (!removidas && mantidas.length === qtd) return 0;

  const bloco = mantidas.slice();
  while (bloco.length < qtd) bloco.push(linhaEmBranco());

  aplicarFormatoModelo(range);
  range.setValues(bloco);

  return removidas;
}


/* ---------------------------------------------------------
   5 e 6) REFORMATAR PELO MODELO
--------------------------------------------------------- */

function reformatarAbaAtiva() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!ehAbaMes(aba.getName())) {
    avisar("Entre em uma aba de mês (ex: Setembro/26) antes de executar.");
    return;
  }

  const linhas = reformatarAba(aba);
  avisar(`${linhas} linha(s) reformatada(s) em "${aba.getName()}".`);
}

function reformatarTodasAsAbasDeMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const resposta = ui.alert(
    "Reformatar todas as abas de mês?",
    "A formatação da aba MODELO será reaplicada em todas as abas de mês. " +
    "Os valores são preservados.",
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  let linhas = 0;
  let abas = 0;

  ss.getSheets().forEach(aba => {
    if (!ehAbaMes(aba.getName())) return;
    linhas += reformatarAba(aba);
    abas++;
  });

  avisar(`${linhas} linha(s) reformatada(s) em ${abas} aba(s) de mês.`);
}

function reformatarAba(aba) {
  const ultima = ultimaLinhaDados(aba);
  if (ultima < PRIMEIRA_LINHA_DADOS) return 0;

  const qtd = ultima - PRIMEIRA_LINHA_DADOS + 1;
  const range = aba.getRange(PRIMEIRA_LINHA_DADOS, 1, qtd, ULTIMA_COLUNA_GRAVACAO);

  // o trim() faz o texto voltar a bater com a lista do dropdown,
  // que é o que traz o chip colorido de volta
  const valores = range.getValues().map(limparTextos);

  aplicarFormatoModelo(range);
  range.setValues(valores);

  return qtd;
}


/* ---------------------------------------------------------
   GRAVAÇÃO SEM DUPLICAR
--------------------------------------------------------- */

function anexarSemDuplicar(aba, linhas) {
  if (!linhas.length) return { inseridos: 0, duplicados: 0 };

  const chaves = {};
  const ultima = ultimaLinhaDados(aba);

  if (ultima >= PRIMEIRA_LINHA_DADOS) {
    aba.getRange(PRIMEIRA_LINHA_DADOS, 1, ultima - PRIMEIRA_LINHA_DADOS + 1, ULTIMA_COLUNA_GRAVACAO)
      .getValues()
      .forEach(linha => {
        if (!linhaVazia(linha)) chaves[chaveLancamento(linha)] = true;
      });
  }

  const novas = [];
  let duplicados = 0;

  linhas.forEach(linha => {
    const chave = chaveLancamento(linha);

    if (chaves[chave]) {
      duplicados++;
      return;
    }

    chaves[chave] = true; // evita duplicar dentro da própria importação
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

  aplicarFormatoModelo(range);
  range.setValues(linhas);
}

function montarLinha(linhaOrigem, vencimento) {
  const linha = linhaEmBranco();

  for (let c = 0; c < COL_CONTABIL; c++) {
    const valor = linhaOrigem[c];
    linha[c] = typeof valor === "string" ? valor.trim() : (valor === undefined ? "" : valor);
  }

  linha[COL_VENCIMENTO - 1] = vencimento;
  linha[COL_CONTABIL - 1] = MESES[vencimento.getMonth()].toUpperCase();
  linha[COL_STATUS - 1] = "";

  return linha;
}


/* ---------------------------------------------------------
   FORMATAÇÃO — sempre a partir do MODELO do CONTROLE FINANCEIRO

   Dentro da própria planilha dá pra usar copyTo, que leva
   tudo (cores, bordas, formatos, validações). Para a planilha
   secundária o copyTo não funciona entre arquivos, então a
   formatação é reaplicada propriedade por propriedade a
   partir do mesmo MODELO — o resultado visual é o mesmo,
   exceto bordas, que a API não deixa ler.
--------------------------------------------------------- */

let MODELO_CACHE = null;

function rangeModelo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modelo = buscarAba(ss, ABA_MODELO);

  if (!modelo) {
    throw new Error(
      `Aba "${ABA_MODELO}" não encontrada nesta planilha. ` +
      `Crie o molde (cabeçalho na linha 1 + linha 2 formatada) antes de rodar.`
    );
  }

  return modelo.getRange(LINHA_MODELO, 1, 1, ULTIMA_COLUNA_GRAVACAO);
}

function formatoModelo() {
  if (MODELO_CACHE) return MODELO_CACHE;

  const r = rangeModelo();

  MODELO_CACHE = {
    fundos: r.getBackgrounds()[0],
    cores: r.getFontColors()[0],
    fontes: r.getFontFamilies()[0],
    tamanhos: r.getFontSizes()[0],
    pesos: r.getFontWeights()[0],
    estilos: r.getFontStyles()[0],
    formatos: r.getNumberFormats()[0],
    alinhamentoH: r.getHorizontalAlignments()[0],
    alinhamentoV: r.getVerticalAlignments()[0],
    quebras: r.getWraps()[0],
    validacoes: r.getDataValidations()[0]
  };

  return MODELO_CACHE;
}

function aplicarFormatoModelo(range) {
  const mesmaPlanilha =
    range.getSheet().getParent().getId() === SpreadsheetApp.getActiveSpreadsheet().getId();

  if (mesmaPlanilha) {
    const modelo = rangeModelo();
    // uma linha de molde copiada sobre um bloco alto se repete em todas as linhas
    modelo.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    modelo.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    return;
  }

  const f = formatoModelo();
  const n = range.getNumRows();

  range.setBackgrounds(repetir(f.fundos, n));
  range.setFontColors(repetir(f.cores, n));
  range.setFontFamilies(repetir(f.fontes, n));
  range.setFontSizes(repetir(f.tamanhos, n));
  range.setFontWeights(repetir(f.pesos, n));
  range.setFontStyles(repetir(f.estilos, n));
  range.setNumberFormats(repetir(f.formatos, n));
  range.setHorizontalAlignments(repetir(f.alinhamentoH, n));
  range.setVerticalAlignments(repetir(f.alinhamentoV, n));
  range.setWraps(repetir(f.quebras, n));
  range.setDataValidations(repetir(f.validacoes, n));
}

function repetir(linha, quantidade) {
  const bloco = [];
  for (let i = 0; i < quantidade; i++) bloco.push(linha.slice());
  return bloco;
}


/* ---------------------------------------------------------
   ABAS DE MÊS
--------------------------------------------------------- */

function obterOuCriarAbaMes(ss, data) {
  const nome = nomeAbaPorData(data);
  const existente = buscarAba(ss, nome);

  if (existente) return existente;

  const modelo = buscarAba(ss, ABA_MODELO);
  let nova;

  if (modelo) {
    nova = modelo.copyTo(ss).setName(nome);

    const maxLinhas = nova.getMaxRows();
    if (maxLinhas >= PRIMEIRA_LINHA_DADOS) {
      nova.getRange(
        PRIMEIRA_LINHA_DADOS, 1,
        maxLinhas - PRIMEIRA_LINHA_DADOS + 1,
        ULTIMA_COLUNA_GRAVACAO
      ).clearContent();
    }
  } else {
    // planilha secundária sem MODELO: cria a aba e leva o cabeçalho
    nova = ss.insertSheet(nome);
    const cabecalho = rangeModelo().getSheet()
      .getRange(1, 1, 1, ULTIMA_COLUNA_GRAVACAO).getValues();
    nova.getRange(1, 1, 1, ULTIMA_COLUNA_GRAVACAO).setValues(cabecalho);
  }

  nova.setFrozenRows(1);
  return nova;
}

function planilhaDestino(linha) {
  const destino = normalizarTexto(linha[COL_DESTINO - 1]);

  if (destino === DESTINO_SECUNDARIO && PLANILHA_SECUNDARIA_ID) {
    return SpreadsheetApp.openById(PLANILHA_SECUNDARIA_ID);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}

// Última linha com FORNECEDOR preenchido dentro de A:K.
// Não usa getLastRow() porque o dashboard em L:AF empurraria
// o resultado para a linha 100 e deixaria um buraco nos dados.
function ultimaLinhaDados(aba) {
  const max = aba.getMaxRows();
  if (max < PRIMEIRA_LINHA_DADOS) return PRIMEIRA_LINHA_DADOS - 1;

  const coluna = aba
    .getRange(PRIMEIRA_LINHA_DADOS, COL_FORNECEDOR, max - PRIMEIRA_LINHA_DADOS + 1, 1)
    .getValues();

  for (let i = coluna.length - 1; i >= 0; i--) {
    if (String(coluna[i][0]).trim() !== "") return PRIMEIRA_LINHA_DADOS + i;
  }

  return PRIMEIRA_LINHA_DADOS - 1;
}

function lerLinhas(aba) {
  const ultima = ultimaLinhaDados(aba);
  if (ultima < PRIMEIRA_LINHA_DADOS) return [];

  const colunas = Math.max(Math.min(aba.getLastColumn(), COL_TERMINA), 1);

  return aba
    .getRange(PRIMEIRA_LINHA_DADOS, 1, ultima - PRIMEIRA_LINHA_DADOS + 1, colunas)
    .getValues()
    .map(linha => {
      const completa = linha.slice(0, COL_TERMINA);
      while (completa.length < COL_TERMINA) completa.push("");
      return completa;
    });
}


/* ---------------------------------------------------------
   CHAVE DO LANÇAMENTO (base da deduplicação)
--------------------------------------------------------- */

function chaveLancamento(linha) {
  return [
    normalizarTexto(linha[COL_TIPO - 1]),
    normalizarTexto(linha[COL_FORNECEDOR - 1]),
    normalizarTexto(linha[COL_NF - 1]),
    normalizarTexto(linha[COL_MOTIVO - 1]),
    normalizarTexto(linha[COL_SUBMOTIVO - 1]),
    chaveNumero(linha[COL_VALOR - 1]),
    chaveNumero(linha[COL_VALOR_TOTAL - 1]),
    chaveData(linha[COL_VENCIMENTO - 1])
  ].join("|");
}

function chaveNumero(valor) {
  if (valor === "" || valor === null || valor === undefined) return "0.00";
  if (typeof valor === "number") return valor.toFixed(2);

  let texto = String(valor).replace(/[R$\s ]/g, "");
  if (texto.indexOf(",") >= 0) texto = texto.replace(/\./g, "").replace(",", ".");

  const numero = Number(texto);
  return isNaN(numero) ? normalizarTexto(valor) : numero.toFixed(2);
}

function chaveData(valor) {
  const data = normalizarData(valor);
  if (!data) return normalizarTexto(valor);

  return Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM-dd");
}


/* ---------------------------------------------------------
   AUXILIARES
--------------------------------------------------------- */

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
  return linha.every(valor => String(valor === null || valor === undefined ? "" : valor).trim() === "");
}

function limparTextos(linha) {
  return linha.map(valor => (typeof valor === "string" ? valor.trim() : valor));
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

function nomeAbaPorData(data) {
  return `${MESES[data.getMonth()]}/${String(data.getFullYear()).slice(-2)}`;
}

function ehAbaMes(nome) {
  return parseAbaMes(nome) !== null;
}

function parseAbaMes(nome) {
  const partes = String(nome || "").replace(/ /g, " ").trim().split("/");
  if (partes.length !== 2) return null;

  const ano = Number(partes[1]);
  const mes = MESES.findIndex(m => normalizarTexto(m) === normalizarTexto(partes[0]));

  if (mes === -1 || isNaN(ano)) return null;

  return { mes: mes, ano: ano };
}

function somarMesesEmNomeAba(nomeAba, quantidade) {
  const info = parseAbaMes(nomeAba);
  if (!info) return nomeAba;

  let mes = info.mes + quantidade;
  let ano = info.ano;

  while (mes > 11) { mes -= 12; ano++; }
  while (mes < 0) { mes += 12; ano--; }

  return `${MESES[mes]}/${String(ano).padStart(2, "0")}`;
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
