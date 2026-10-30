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
