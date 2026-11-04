/**
 * Avaliador de elencos · Grupo B
 * Recebe as avaliações da página avaliacao.html e grava nesta planilha.
 *
 * Abas criadas pelo setup():
 *  - Codigos    : o código único dos patronos, na célula B2 (pode ser trocado à mão)
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
const POSICOES = ['Goleiro', 'Zagueiro', 'Lateral direito', 'Lateral esquerdo', 'Volante', 'Meia armador', 'Ponta direita', 'Ponta esquerda', 'Centroavante'];
const CAB_NOTAS = ['Time', 'Nº', 'Atleta', 'Nível', 'Pontos', 'Atualizado em', 'Posição'];
const CAB_HIST = ['Enviado em', 'Time', 'Atleta', 'Nível', 'Posição'];

/** Rode UMA vez pelo editor (botão Executar) para montar a planilha. */
function setup() {
  const ss = SpreadsheetApp.getActive();

  let cod = ss.getSheetByName(ABA_COD) || ss.insertSheet(ABA_COD);
  if (String(cod.getRange('A2').getValue()) !== 'TODOS') {
    cod.clear();
    cod.getRange('B:B').setNumberFormat('@');
    cod.appendRow(['Quem usa', 'Código dos patronos']);
    cod.appendRow(['TODOS', String(Math.floor(1000 + Math.random() * 9000))]);
  }

  const notas = ss.getSheetByName(ABA_NOTAS) || ss.insertSheet(ABA_NOTAS);
  if (notas.getLastRow() < 1) notas.appendRow(CAB_NOTAS);
  else notas.getRange(1, 1, 1, CAB_NOTAS.length).setValues([CAB_NOTAS]);

  const hist = ss.getSheetByName(ABA_HIST) || ss.insertSheet(ABA_HIST);
  if (hist.getLastRow() < 1) hist.appendRow(CAB_HIST);
  else hist.getRange(1, 1, 1, CAB_HIST.length).setValues([CAB_HIST]);

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
    if (!codigoValido(dados.codigo)) return resposta({ ok: false, erro: 'Código incorreto.' });

    if (dados.acao === 'carregar') return resposta(Object.assign({ ok: true }, todasAsNotas()));
    if (dados.acao === 'enviar') {
      const time = String(dados.time || '').toUpperCase();
      if (TIMES.indexOf(time) < 0) return resposta({ ok: false, erro: 'Time não encontrado.' });
      return resposta(gravar(time, dados.notas || []));
    }
    return resposta({ ok: false, erro: 'Ação desconhecida.' });
  } catch (err) {
    return resposta({ ok: false, erro: 'Erro no servidor: ' + err.message });
  }
}

function doGet() {
  return resposta({ ok: true, mensagem: 'Avaliador de elencos no ar.' });
}

function codigoValido(codigo) {
  const certo = String(SpreadsheetApp.getActive().getSheetByName(ABA_COD).getRange('B2').getValue()).trim();
  return certo !== '' && certo === String(codigo || '').trim();
}

/** Devolve níveis e posições de todos os times: { notas: { BRA: { 'NOME': 'Estrela' } }, posicoes: { BRA: { 'NOME': 'Volante' } } } */
function todasAsNotas() {
  const linhas = SpreadsheetApp.getActive().getSheetByName(ABA_NOTAS).getDataRange().getValues().slice(1);
  const notas = {}, posicoes = {};
  TIMES.forEach(t => { notas[t] = {}; posicoes[t] = {}; });
  linhas.forEach(l => {
    if (!notas[l[0]]) return;
    if (l[3]) notas[l[0]][l[2]] = l[3];
    if (l[6]) posicoes[l[0]][l[2]] = l[6];
  });
  return { notas: notas, posicoes: posicoes };
}

function gravar(time, notas) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(ABA_NOTAS);
    const agora = new Date();
    const todas = sh.getDataRange().getValues();
    const largura = CAB_NOTAS.length;
    const ajusta = l => { const r = l.slice(0, largura); while (r.length < largura) r.push(''); return r; };
    const outras = todas.slice(1).filter(l => l[0] !== time).map(ajusta);
    const novas = notas
      .map(x => ({ n: x.n, atleta: x.atleta, nivel: PONTOS[x.nivel] ? x.nivel : '', pos: POSICOES.indexOf(x.posicao) >= 0 ? x.posicao : '' }))
      .filter(x => x.nivel || x.pos)
      .map(x => [time, x.n, x.atleta, x.nivel, x.nivel ? PONTOS[x.nivel] : '', agora, x.pos]);
    const final = [CAB_NOTAS].concat(outras, novas);
    sh.clearContents();
    sh.getRange(1, 1, final.length, largura).setValues(final);
    sh.getRange(1, 1, 1, largura).setFontWeight('bold');

    const hist = ss.getSheetByName(ABA_HIST);
    if (novas.length) hist.getRange(hist.getLastRow() + 1, 1, novas.length, CAB_HIST.length).setValues(novas.map(l => [agora, time, l[2], l[3], l[6]]));

    atualizarResumo();
    const avaliados = novas.filter(l => l[3]).length;
    const pontos = novas.reduce((s, l) => s + (Number(l[4]) || 0), 0);
    return { ok: true, avaliados: avaliados, pontos: pontos };
  } finally {
    lock.releaseLock();
  }
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
