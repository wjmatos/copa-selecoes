/**
 * Votação da seleção da fase de grupos.
 * Cole este código no Apps Script da planilha (Extensões > Apps Script) e publique como "App da Web".
 * Abas usadas: Jogadores, Votos.
 */
var ABA_VOTOS = 'Votos';
var ABA_JOGADORES = 'Jogadores';
var ESQUEMAS = {
  '4-3-3': ['Atacante','Atacante','Atacante','Goleiro','Lateral','Lateral','Meia','Meia','Meia','Zagueiro','Zagueiro'],
  '4-4-2': ['Atacante','Atacante','Goleiro','Lateral','Lateral','Meia','Meia','Volante','Volante','Zagueiro','Zagueiro']
};

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Recebe um voto: { id, esquema, voto: [{posicao, jogador, time} x 11] }
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.tipo === 'enquete') return enquetePost_(d);
    if (d.tipo === 'foto') return fotoPost_(d);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var esperado = ESQUEMAS[d.esquema];
    if (!esperado || !d.id || String(d.id).length > 80 || !d.voto || d.voto.length !== 11) return json_({ ok: false, erro: 'voto invalido' });

    // posições corretas para o esquema
    var pos = d.voto.map(function (v) { return v.posicao; }).sort();
    if (pos.join('|') !== esperado.join('|')) return json_({ ok: false, erro: 'posicoes invalidas' });

    // jogadores existentes e sem repetição
    var lista = ss.getSheetByName(ABA_JOGADORES).getRange(2, 1, ss.getSheetByName(ABA_JOGADORES).getLastRow() - 1, 2).getValues();
    var validos = {};
    lista.forEach(function (r) { validos[r[0] + '|' + r[1]] = true; });
    var vistos = {};
    for (var i = 0; i < d.voto.length; i++) {
      var k = d.voto[i].time + '|' + d.voto[i].jogador;
      if (!validos[k] || vistos[k]) return json_({ ok: false, erro: 'jogador invalido ou repetido' });
      vistos[k] = true;
    }

    // um voto por ID (aparelho + esquema)
    var sh = ss.getSheetByName(ABA_VOTOS);
    var ult = sh.getLastRow();
    if (ult > 1) {
      var ids = sh.getRange(2, 2, ult - 1, 1).getValues();
      for (var j = 0; j < ids.length; j++) if (ids[j][0] === d.id) return json_({ ok: false, erro: 'duplicado' });
    }

    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
    var linhas = d.voto.map(function (v) { return [agora, d.id, d.esquema, v.posicao, v.jogador, v.time]; });
    sh.getRange(ult + 1, 1, linhas.length, 6).setNumberFormat('@').setValues(linhas);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Devolve o resultado parcial: os mais votados por posição, somando os votos de todos os esquemas
function doGet(e) {
  if (e && e.parameter && e.parameter.acao === 'mural') return json_({ fotos: muralLista_() });
  if (e && e.parameter && e.parameter.acao === 'noticias') return json_({ noticias: noticiasLista_() });
  if (e && e.parameter && e.parameter.acao === 'jogos') return json_({ jogos: jogosLista_() });
  if (e && e.parameter && e.parameter.acao === 'enquete') return json_({ jogos: enqueteResultado_() });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_VOTOS);
  var n = sh.getLastRow() - 1;
  var v = n > 0 ? sh.getRange(2, 1, n, 6).getValues() : [];
  var ids = {}, cont = {};
  v.forEach(function (r) {
    var id = r[1], esq = r[2], pos = r[3], jog = r[4], time = r[5];
    if (!id || !jog) return;
    ids[esq] = ids[esq] || {};
    ids[esq][id] = 1;
    var k = pos + '|' + jog + '|' + time;
    cont[k] = (cont[k] || 0) + 1;
  });
  var res = {};
  Object.keys(cont).forEach(function (k) {
    var p = k.split('|');
    res[p[0]] = res[p[0]] || [];
    res[p[0]].push({ jogador: p[1], time: p[2], votos: cont[k] });
  });
  Object.keys(res).forEach(function (pos) {
    res[pos].sort(function (a, b) { return b.votos - a.votos; });
    res[pos] = res[pos].slice(0, 8);
  });
  var tot = {}, soma = 0;
  Object.keys(ids).forEach(function (esq) { tot[esq] = Object.keys(ids[esq]).length; soma += tot[esq]; });
  return json_({ total: soma, totais: tot, resultado: res });
}

// ===== Enquete rápida por jogo (vitória da casa, empate ou vitória do visitante) =====
function abaEnquetes_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Enquetes');
  if (!sh) {
    sh = ss.insertSheet('Enquetes');
    sh.getRange(1, 1, 1, 4).setValues([['Data/hora', 'Voto (ID)', 'Jogo', 'Palpite']]);
  }
  return sh;
}

function enquetePost_(d) {
  var ok = { casa: 1, empate: 1, fora: 1 };
  var jogo = parseInt(d.jogo, 10);
  if (!d.id || String(d.id).length > 80 || !(jogo >= 1 && jogo <= 500) || !ok[d.palpite]) return json_({ ok: false, erro: 'enquete invalida' });
  var sh = abaEnquetes_();
  var ult = sh.getLastRow();
  if (ult > 1) {
    var ids = sh.getRange(2, 2, ult - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (ids[i][0] === d.id) return json_({ ok: false, erro: 'duplicado' });
  }
  var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  sh.getRange(ult + 1, 1, 1, 4).setNumberFormat('@').setValues([[agora, d.id, String(jogo), d.palpite]]);
  return json_({ ok: true });
}

function enqueteResultado_() {
  var sh = abaEnquetes_();
  var n = sh.getLastRow() - 1;
  var v = n > 0 ? sh.getRange(2, 1, n, 4).getValues() : [];
  var res = {};
  v.forEach(function (r) {
    var j = String(r[2]), p = r[3];
    if (!j || !p) return;
    res[j] = res[j] || { casa: 0, empate: 0, fora: 0 };
    if (res[j][p] !== undefined) res[j][p]++;
  });
  return res;
}


// Placares lancados na aba "Jogos" (colunas: Jogo, Data, Hora, Casa, Fora, Gols casa, Gols fora, Ao vivo?)
function jogosLista_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jogos');
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues(), out = [];
  v.forEach(function (r) {
    var n = Number(r[0]); if (!n) return;
    var gc = r[5], gf = r[6];
    if (gc === '' || gf === '' || isNaN(Number(gc)) || isNaN(Number(gf))) return;
    out.push({ n: n, gc: Number(gc), gf: Number(gf), vivo: String(r[7]).trim() !== '' });
  });
  return out;
}


// Aba "Noticias": Posicao | Etiqueta | Titulo | Texto | Times | Esconder?
function noticiasLista_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notícias');
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues(), out = [];
  v.forEach(function (r, i) {
    if (!String(r[2]).trim() || String(r[5]).trim() !== '') return;
    var p = Number(r[0]);
    out.push({ pos: p > 0 ? p : 0, ordem: i, tag: String(r[1]), t: String(r[2]), x: String(r[3]), times: String(r[4]) });
  });
  return out;
}


// ===== MURAL DE FOTOS =====
// Fotos chegam como "pendentes" (privadas, na pasta do Drive). Para aprovar, escreva ok na coluna Status da aba Mural.
var ABA_MURAL = 'Mural';
var EMAIL_AVISO = 'wjmatos@gmail.com';
var TIMES_OK = ['EUA', 'PERU', 'URUGUAI', 'MÉXICO', 'ARGENTINA', 'CHILE', 'PARAGUAI', 'BRASIL'];

function abaMural_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA_MURAL);
  if (!sh) {
    sh = ss.insertSheet(ABA_MURAL);
    sh.appendRow(['Data/hora', 'Time', 'Nome', 'Legenda', 'Status (escreva ok ou não)', 'ID do arquivo', 'Dispositivo', 'Publicada?']);
    sh.setFrozenRows(1);
    sh.getRange('A:H').setNumberFormat('@');
    sh.getRange('A1:H1').setFontWeight('bold');
    sh.getRange('E2:E1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ok', 'não'], true).setAllowInvalid(true).build());
  }
  return sh;
}

function pastaMural_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PASTA_MURAL');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var p = DriveApp.createFolder('Mural Copa de Seleções');
  props.setProperty('PASTA_MURAL', p.getId());
  return p;
}

// Recebe uma foto: { tipo:'foto', id, time, nome, legenda, img:'data:image/jpeg;base64,...' }
function fotoPost_(d) {
  var time = String(d.time || '').toUpperCase().trim();
  if (TIMES_OK.indexOf(time) < 0 || !d.id || String(d.id).length > 80) return json_({ ok: false, erro: 'dados invalidos' });
  var m = /^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/.exec(String(d.img || ''));
  if (!m || m[1].length > 1000000) return json_({ ok: false, erro: 'imagem invalida' });
  var cache = CacheService.getScriptCache(), k = 'foto_' + d.id, n = Number(cache.get(k) || 0);
  if (n >= 5) return json_({ ok: false, erro: 'limite' });
  cache.put(k, String(n + 1), 21600);
  var sh = abaMural_();
  var nome = String(d.nome || '').replace(/[\r\n]+/g, ' ').slice(0, 40).trim();
  var leg = String(d.legenda || '').replace(/[\r\n]+/g, ' ').slice(0, 140).trim();
  var blob = Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', 'foto_' + Date.now() + '.jpg');
  var f = pastaMural_().createFile(blob);
  var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  sh.appendRow([agora, time, nome, leg, '', f.getId(), String(d.id), '']);
  try {
    MailApp.sendEmail(EMAIL_AVISO, 'Nova foto no mural (' + time + ')',
      'Foto nova esperando aprovação.\n\nTime: ' + time + '\nNome: ' + (nome || '(sem nome)') + '\nLegenda: ' + (leg || '(sem legenda)') +
      '\n\nPara ver e aprovar, abra a aba Mural e escreva ok na coluna Status:\n' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + sh.getSheetId() +
      '\n\nA foto está aqui: ' + f.getUrl());
  } catch (err) {}
  return json_({ ok: true });
}

// Lista as fotos aprovadas. Libera o link da foto ao aprovar e o retira se voltar para não.
function muralLista_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_MURAL);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues(), out = [];
  v.forEach(function (r, i) {
    var st = String(r[4]).trim().toLowerCase(), id = String(r[5]).trim(), pub = String(r[7]).trim() === 'sim';
    if (!id) return;
    var ok = st === 'ok' || st === 'sim' || st === 'aprovada';
    try {
      if (ok && !pub) { DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); sh.getRange(i + 2, 8).setValue('sim'); pub = true; }
      if (!ok && pub) { DriveApp.getFileById(id).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); sh.getRange(i + 2, 8).setValue(''); }
    } catch (err) { return; }
    if (ok) out.push({ id: id, time: String(r[1]), nome: String(r[2]), leg: String(r[3]), data: String(r[0]) });
  });
  return out.reverse().slice(0, 200);
}
