/**
 * Avaliador de elencos · Grupo B
 * Recebe as avaliações da página avaliacao.html e grava nesta planilha.
 *
 * Abas criadas pelo setup():
 *  - Codigos    : o código de cada patrono (só a organização vê)
 *  - Avaliacoes : a avaliação atual de cada atleta (o último envio vale)
 *  - Historico  : todos os envios, com data e hora
 *  - Resumo     : pontos e níveis por time, recalculados a cada envio
 */
const ABA_COD = 'Codigos';
const ABA_NOTAS = 'Avaliacoes';
const ABA_HIST = 'Historico';
const ABA_RES = 'Resumo';
const TIMES = ['EUA', 'PER', 'URU', 'MEX', 'ARG', 'CHI', 'PAR', 'BRA'];
const PONTOS = { 'Estrela': 3, 'Joga muito': 2, 'Agrega valor': 1 };

/** Rode UMA vez pelo editor (botão Executar) para montar a planilha. */
function setup() {
  const ss = SpreadsheetApp.getActive();

  let cod = ss.getSheetByName(ABA_COD) || ss.insertSheet(ABA_COD);
  if (cod.getLastRow() < 2) {
    cod.clear();
    cod.appendRow(['Time', 'Código do patrono']);
    TIMES.forEach(t => cod.appendRow([t, String(Math.floor(1000 + Math.random() * 9000))]));
    cod.getRange('B:B').setNumberFormat('@');
  }

  const notas = ss.getSheetByName(ABA_NOTAS) || ss.insertSheet(ABA_NOTAS);
  if (notas.getLastRow() < 1) notas.appendRow(['Time', 'Nº', 'Atleta', 'Nível', 'Pontos', 'Atualizado em']);

  const hist = ss.getSheetByName(ABA_HIST) || ss.insertSheet(ABA_HIST);
  if (hist.getLastRow() < 1) hist.appendRow(['Enviado em', 'Time', 'Atleta', 'Nível']);

  if (!ss.getSheetByName(ABA_RES)) ss.insertSheet(ABA_RES);
  atualizarResumo();
  [cod, notas, hist].forEach(sh => sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold'));
}

/** Recalcula a aba Resumo a partir da aba Avaliacoes (roda a cada envio). */
function atualizarResumo() {
  const ss = SpreadsheetApp.getActive();
  const res = ss.getSheetByName(ABA_RES);
  const linhas = ss.getSheetByName(ABA_NOTAS).getDataRange().getValues().slice(1);
  const tabela = [['Time', 'Avaliados', 'Estrela (A)', 'Joga muito (B)', 'Agrega valor (C)', 'Total de pontos', 'Média por atleta']];
  TIMES.forEach(t => {
    const doTime = linhas.filter(l => l[0] === t && l[3]);
    const conta = nivel => doTime.filter(l => l[3] === nivel).length;
    const total = doTime.reduce((s, l) => s + (PONTOS[l[3]] || 0), 0);
    tabela.push([t, doTime.length, conta('Estrela'), conta('Joga muito'), conta('Agrega valor'), total, doTime.length ? total / doTime.length : 0]);
  });
  res.clear();
  res.getRange(1, 1, tabela.length, 7).setValues(tabela);
  res.getRange(1, 1, 1, 7).setFontWeight('bold');
  res.getRange(2, 7, TIMES.length, 1).setNumberFormat('0.00');
}

function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const time = String(dados.time || '').toUpperCase();
    if (TIMES.indexOf(time) < 0) return resposta({ ok: false, erro: 'Time não encontrado.' });
    if (!codigoValido(time, dados.codigo)) return resposta({ ok: false, erro: 'Código incorreto para este time.' });

    if (dados.acao === 'carregar') return resposta({ ok: true, notas: notasDoTime(time) });
    if (dados.acao === 'enviar') return resposta(gravar(time, dados.notas || []));
    return resposta({ ok: false, erro: 'Ação desconhecida.' });
  } catch (err) {
    return resposta({ ok: false, erro: 'Erro no servidor: ' + err.message });
  }
}

function doGet() {
  return resposta({ ok: true, mensagem: 'Avaliador de elencos no ar.' });
}

function codigoValido(time, codigo) {
  const linhas = SpreadsheetApp.getActive().getSheetByName(ABA_COD).getDataRange().getValues();
  return linhas.some(l => String(l[0]).toUpperCase() === time && String(l[1]).trim() === String(codigo || '').trim());
}

function notasDoTime(time) {
  const linhas = SpreadsheetApp.getActive().getSheetByName(ABA_NOTAS).getDataRange().getValues().slice(1);
  const notas = {};
  linhas.forEach(l => { if (l[0] === time && l[3]) notas[l[2]] = l[3]; });
  return notas;
}

function gravar(time, notas) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(ABA_NOTAS);
    const agora = new Date();
    const todas = sh.getDataRange().getValues();
    const cabecalho = todas[0];
    const outras = todas.slice(1).filter(l => l[0] !== time);
    const novas = notas
      .filter(x => PONTOS[x.nivel])
      .map(x => [time, x.n, x.atleta, x.nivel, PONTOS[x.nivel], agora]);
    const final = [cabecalho].concat(outras, novas);
    sh.clearContents();
    sh.getRange(1, 1, final.length, 6).setValues(final);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');

    const hist = ss.getSheetByName(ABA_HIST);
    if (novas.length) hist.getRange(hist.getLastRow() + 1, 1, novas.length, 4).setValues(novas.map(l => [agora, time, l[2], l[3]]));

    atualizarResumo();
    const pontos = novas.reduce((s, l) => s + l[4], 0);
    return { ok: true, avaliados: novas.length, pontos: pontos };
  } finally {
    lock.releaseLock();
  }
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
