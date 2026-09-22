/* PARTE 2 de 3 — fixos, duplicados e reformatação
   Cole este conteúdo num arquivo chamado 2-rotinas.gs */

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

  const proximoMes = somarMeses(new Date(2000 + info.ano, info.mes, 1), 1);
  const abaDestino = obterOuCriarAbaMes(ss, proximoMes);

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

  ss.getSheets().forEach(aba => {
    if (!ehAbaMes(aba.getName())) return;

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
    `"Remover duplicados" se estiver tudo certo.`
  );
}

function removerDuplicadosAbaAtiva() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!ehAbaMes(aba.getName())) {
    avisar("Entre na aba do mês (ex: Setembro26) antes de executar.");
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

  ss.getSheets().forEach(aba => {
    if (!ehAbaMes(aba.getName())) return;

    const n = removerDuplicados(aba);
    if (n) detalhe.push(`${aba.getName()}: ${n}`);
    removidas += n;
  });

  avisar(
    `${removidas} linha(s) duplicada(s) removida(s).\n\n` +
    (detalhe.length ? detalhe.join("\n") : "Nenhuma aba tinha duplicados.")
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

  aplicarFormatoModelo(range);
  range.setValues(bloco);

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

  avisar(`${reformatarAba(aba)} linha(s) reformatada(s) em "${aba.getName()}".`);
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

  ss.getSheets().forEach(aba => {
    if (!ehAbaMes(aba.getName())) return;
    linhas += reformatarAba(aba);
    abas++;
  });

  avisar(`${linhas} linha(s) reformatada(s) em ${abas} aba(s) de mês.`);
}

function reformatarAba(aba) {
  const ultima = ultimaLinhaDados(aba, COL_FORNECEDOR);
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



/* ===== FIM DA PARTE 2 DE 3 =====
   Se você não está vendo esta linha no fim do que colou,
   a colagem veio cortada: apague e cole de novo. */
