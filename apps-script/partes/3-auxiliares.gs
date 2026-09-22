/* PARTE 3 de 3 — gravação, formatação e auxiliares
   Cole este conteúdo num arquivo chamado 3-auxiliares.gs */

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

  aplicarFormatoModelo(range);
  range.setValues(linhas);
}


/* ---------------------------------------------------------
   FORMATAÇÃO — sempre a partir da aba Modelo desta planilha
--------------------------------------------------------- */

function aplicarFormatoModelo(range) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modelo = buscarAba(ss, ABA_MODELO);

  if (!modelo) {
    throw new Error(
      `Aba "${ABA_MODELO}" não encontrada. Ela é o molde de formatação: ` +
      `cabeçalho na linha 1 e linha 2 formatada.`
    );
  }

  const rangeModelo = modelo.getRange(LINHA_MODELO, 1, 1, ULTIMA_COLUNA_GRAVACAO);

  // uma linha de molde copiada sobre um bloco alto se repete em todas as linhas
  rangeModelo.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  rangeModelo.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
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

/* ===== FIM DA PARTE 3 DE 3 =====
   Se você não está vendo esta linha no fim do que colou,
   a colagem veio cortada: apague e cole de novo. */
