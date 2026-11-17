/******************************************************************
 * 🤖 ROBÔ STAGE ZONE — KN SRJ1 Cordovil
 *
 * Abas que o robô usa (ele cria sozinho em "Instalar robô"):
 *   • Bipagem HH        → o AppSheet do HH grava aqui (coluna A = data/hora, B = ID lido)
 *   • Colar Faltantes   → importe aqui o CSV de faltantes baixado do sistema
 *   • Painel            → endereços com faltantes, diminuindo a cada pacote achado
 *   • Faltantes nas gaiolas → lista de cada ID faltante, endereço e situação (o app Stage Zone lê esta aba)
 *   • Stage Zone        → bipagens das gaiolas (puxadas da nuvem do app Stage Zone)
 *   • App Endereços     → tabela do AppSheet: endereço, faltantes, achados, restam, status
 *   • App Reporte       → tabela do AppSheet: números do topo + texto pronto pro WhatsApp
 *
 * Rode pela conta pessoal (megakn2026) — na conta da empresa o Apps Script é bloqueado.
 ******************************************************************/

const SB_URL = 'https://wimqujskqgabjboihiei.supabase.co/rest/v1/stage_zone_snapshot';
const SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndpbXF1anNrcWdhYmpib2loaWVpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTI0MDI2NTgsImV4cCI6MjEwNzk3ODY1OH0.Nayqohgw12Ch5TAkfF3orYR1T-W0i7LqHWeJjAMgq7Y';

const ABA = {
  BIP: 'Bipagem HH',
  COLAR: 'Colar Faltantes',
  PAINEL: 'Painel',
  LISTA: 'Faltantes nas gaiolas',
  STAGE: 'Stage Zone',
  APP_END: 'App Endereços',
  APP_REP: 'App Reporte'
};
const BIP_CAB = ['Data/Hora', 'ID lido', 'Resultado', 'Endereço', 'Restam no endereço', 'Rota'];
const COR = { azul: '#0b2944', azul2: '#123b5c', verde: '#d9f7e8', verdeT: '#0a7a4b', vermelho: '#ffe1e5', vermelhoT: '#c4122f', amarelo: '#fff4cc', cinza: '#eef2f6' };

/* ======================= MENU ======================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🤖 Stage Zone')
    .addItem('▶ Atualizar painel agora', 'atualizarAgora')
    .addItem('☁ Puxar Stage Zone da nuvem', 'puxarStageMenu')
    .addSeparator()
    .addItem('🧹 Zerar bipagens (novo ciclo)', 'zerarBipagens')
    .addSeparator()
    .addItem('⚙ Instalar robô (1ª vez)', 'instalar')
    .addItem('⏸ Pausar robô', 'pausar')
    .addToUi();
}

/* ======================= INSTALAÇÃO ======================= */
function instalar() {
  const ss = SpreadsheetApp.getActive();
  const bip = aba_(ABA.BIP);
  if (bip.getLastRow() === 0) {
    bip.getRange(1, 1, 1, BIP_CAB.length).setValues([BIP_CAB]);
  }
  estiloCab_(bip, BIP_CAB.length);
  bip.setFrozenRows(1);
  bip.setColumnWidths(1, 6, 150); bip.setColumnWidth(3, 260);

  const col = aba_(ABA.COLAR);
  if (col.getLastRow() === 0) {
    col.getRange(1, 1).setValue('Importe aqui o CSV de faltantes: Arquivo › Importar › Fazer upload › "Substituir página atual" (com esta aba aberta).');
  }
  aba_(ABA.PAINEL); aba_(ABA.LISTA); aba_(ABA.STAGE); aba_(ABA.APP_END); aba_(ABA.APP_REP);

  // gatilhos: a cada 1 minuto (pega o que o AppSheet gravou) + quando alguém mexe na planilha
  pausar(true);
  ScriptApp.newTrigger('robo').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('aoMudar').forSpreadsheet(ss).onChange().create();

  try { puxarStage_(); } catch (e) { ss.toast('Não consegui puxar a Stage Zone: ' + e.message, '🤖 Robô', 10); }
  robo(true);
  ss.setActiveSheet(aba_(ABA.PAINEL));
  ss.toast('Robô instalado e ligado. Ele confere as bipagens a cada minuto.', '🤖 Stage Zone', 8);
}

function pausar(silencioso) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['robo', 'aoMudar'].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  if (silencioso !== true) SpreadsheetApp.getActive().toast('Robô pausado. Use "Instalar robô" para ligar de novo.', '🤖 Stage Zone', 6);
}

function aoMudar(e) {
  // importou faltantes ou mexeu em alguma aba → recalcula na hora
  robo(true);
}

function atualizarAgora() { robo(true); SpreadsheetApp.getActive().toast('Painel atualizado.', '🤖 Stage Zone', 4); }
function puxarStageMenu() {
  const n = puxarStage_();
  robo(true);
  SpreadsheetApp.getActive().toast('Stage Zone puxada da nuvem: ' + n + ' pacotes.', '🤖 Stage Zone', 6);
}

function zerarBipagens() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Zerar bipagens', 'Apagar todas as bipagens do HH e começar um ciclo novo?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const bip = aba_(ABA.BIP);
  if (bip.getLastRow() > 1) bip.getRange(2, 1, bip.getLastRow() - 1, bip.getMaxColumns()).clearContent().setBackground(null).setFontColor(null);
  PropertiesService.getDocumentProperties().deleteProperty('ultimaLinha');
  robo(true);
}

/* ======================= ROBÔ ======================= */
function robo(forcar) {
  forcar = forcar === true;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return;
  try {
    const props = PropertiesService.getDocumentProperties();
    const bip = aba_(ABA.BIP);
    const ultima = bip.getLastRow();
    const feita = Number(props.getProperty('ultimaLinha') || 1);
    if (!forcar && ultima <= feita) return;           // nada novo → sai rápido

    const B = base_();                                 // faltantes + stage zone
    const achados = new Map();                         // id -> hora em que foi bipado
    const emAuditoria = new Set();                     // endereços que já começaram a ser bipados no HH

    if (ultima > 1) {
      const n = ultima - 1;
      const rg = bip.getRange(2, 1, n, BIP_CAB.length);
      const v = rg.getValues();
      const out = [], bg = [], fc = [];

      // 1º: o que já foi processado conta como achado; toda gaiola com pacote bipado fica "em auditoria"
      for (let i = 0; i < n; i++) {
        const id = normId_(v[i][1]);
        if (id && String(v[i][2]).indexOf('✅') === 0) achados.set(id, v[i][0]);
        const pg = id && B.stage.get(id);
        if (pg) emAuditoria.add(pg.g);
      }
      // 2º: processa as linhas novas (sem resultado), em ordem
      const restam = restamPorEnd_(B, achados);
      let mudou = false;
      for (let i = 0; i < n; i++) {
        const row = v[i];
        let res = row[2], end = row[3], rst = row[4], rota = row[5];
        if (row[1] !== '' && row[2] === '') {
          mudou = true;
          const bruto = String(row[1]);
          const id = normId_(bruto);
          const p = id && B.stage.get(id);
          const f = id && B.falt.get(id);
          if (!id) {
            res = normCage_(bruto) ? '⚠ QR de gaiola — bipe o pacote' : '⚠ Leitura inválida';
            end = ''; rst = ''; rota = '';
          } else if (achados.has(id)) {
            res = '🔁 Já bipado antes';
            end = p ? p.g : ''; rst = p ? (restam.get(p.g) || 0) : ''; rota = f ? f.rota : '';
          } else if (f && p) {
            achados.set(id, row[0] || new Date());
            restam.set(p.g, Math.max(0, (restam.get(p.g) || 0) - 1));
            res = '✅ ACHOU — faltante'; end = p.g; rst = restam.get(p.g); rota = f.rota;
          } else if (f) {
            res = '🟡 Faltante, mas não estava em gaiola'; end = ''; rst = ''; rota = f.rota;
          } else {
            res = '❌ Não é faltante'; end = p ? p.g : ''; rst = ''; rota = '';
          }
          if (row[0] === '') row[0] = new Date();
        }
        out.push([row[0], row[1], res, end, rst, rota]);
        const r = String(res);
        bg.push(Array(6).fill(r.indexOf('✅') === 0 ? COR.verde : r.indexOf('❌') === 0 ? COR.vermelho : r ? COR.amarelo : null));
        fc.push(Array(6).fill(r.indexOf('✅') === 0 ? COR.verdeT : r.indexOf('❌') === 0 ? COR.vermelhoT : '#000000'));
      }
      if (mudou) { rg.setValues(out); rg.setBackgrounds(bg); rg.setFontColors(fc); }
    }

    B.emAuditoria = emAuditoria;
    painel_(B, achados);
    lista_(B, achados);
    props.setProperty('ultimaLinha', String(ultima));
  } finally {
    lock.releaseLock();
  }
}

/* ======================= BASE ======================= */
function base_() {
  // faltantes (último arquivo importado na aba Colar Faltantes)
  const falt = new Map();
  const col = aba_(ABA.COLAR);
  if (col.getLastRow() > 0) {
    const v = col.getDataRange().getDisplayValues();
    let hi = -1;
    for (let i = 0; i < Math.min(v.length, 15); i++) {
      if (v[i].some(x => /id do pacote|^id$|pacote/i.test(String(x).trim()))) { hi = i; break; }
    }
    const h = hi >= 0 ? v[hi].map(x => String(x).toLowerCase()) : [];
    let cId = h.findIndex(x => /id do pacote|^id$|pacote/.test(x));
    let cR = h.findIndex(x => /rota|route/.test(x));
    if (cId < 0) cId = 0;
    for (let i = hi + 1; i < v.length; i++) {
      const id = normId_(v[i][cId]);
      if (id) falt.set(id, { rota: cR >= 0 ? String(v[i][cR]).trim() : '' });
    }
  }
  // stage zone (última bipagem de cada pacote)
  const stage = new Map();
  const st = aba_(ABA.STAGE);
  if (st.getLastRow() > 1) {
    st.getRange(2, 1, st.getLastRow() - 1, 3).getDisplayValues().forEach(r => {
      const id = normId_(r[1]); if (id && r[0]) stage.set(id, { g: r[0], h: r[2] });
    });
  }
  // endereço -> faltantes que estão nele
  const porEnd = new Map();
  falt.forEach((f, id) => {
    const p = stage.get(id); if (!p) return;
    if (!porEnd.has(p.g)) porEnd.set(p.g, []);
    porEnd.get(p.g).push(id);
  });
  const props = PropertiesService.getDocumentProperties();
  return { falt, stage, porEnd, dataStage: props.getProperty('dataStage') || '' };
}

function emAud_(B, g) { return !!(B.emAuditoria && B.emAuditoria.has(g)); }
function status_(B, x) { return x.rst === 0 ? '✅ Auditada' : emAud_(B, x.g) ? '🟡 Em auditoria' : '🔴 A buscar'; }

function restamPorEnd_(B, achados) {
  const m = new Map();
  B.porEnd.forEach((ids, g) => m.set(g, ids.filter(i => !achados.has(i)).length));
  return m;
}

/* ======================= PAINEL ======================= */
function painel_(B, achados) {
  const sh = aba_(ABA.PAINEL);
  const L = [];
  B.porEnd.forEach((ids, g) => {
    const ach = ids.filter(i => achados.has(i)).length;
    const rotas = [...new Set(ids.map(i => B.falt.get(i).rota.replace(/_(AM|PM|SD|CPM)\d*$/i, '')).filter(Boolean))];
    L.push({ g, tot: ids.length, ach, rst: ids.length - ach, rotas });
  });
  L.sort((a, b) => (a.rst === 0) - (b.rst === 0) || b.rst - a.rst || a.g.localeCompare(b.g, 'pt', { numeric: true }));

  const tot = L.reduce((s, x) => s + x.tot, 0), ach = L.reduce((s, x) => s + x.ach, 0), rst = tot - ach;
  const abertos = L.filter(x => x.rst > 0).length;
  const pct = tot ? ach / tot : 0;
  // ordem: em auditoria primeiro, depois a buscar, auditadas no fim
  L.sort((a, b) => (a.rst === 0) - (b.rst === 0) || emAud_(B, b.g) - emAud_(B, a.g) || b.rst - a.rst || a.g.localeCompare(b.g, 'pt', { numeric: true }));

  sh.clear(); sh.clearConditionalFormatRules();
  sh.getRange('A1:F1').merge().setValue('📦 FALTANTES STAGE ZONE — BIPAGEM HH')
    .setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground(COR.azul).setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);
  sh.getRange('A2:F2').merge().setValue('SRJ1 · Cordovil' + (B.dataStage ? '  ·  Stage Zone de ' + B.dataStage : '') +
      '  ·  atualizado ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM HH:mm'))
    .setFontColor('#cfe3f2').setBackground(COR.azul2);

  const k = [['RESTAM BUSCAR', rst, COR.vermelhoT], ['JÁ ACHADOS', ach, COR.verdeT], ['ENDEREÇOS ABERTOS', abertos, '#0b5f99'], ['CONCLUÍDO', pct, '#0b5f99']];
  k.forEach((x, i) => {
    const c = sh.getRange(4, 1 + i); const n = sh.getRange(5, 1 + i);
    c.setValue(x[0]).setFontSize(9).setFontWeight('bold').setFontColor('#5b7488').setBackground(COR.cinza);
    n.setValue(x[1]).setFontSize(26).setFontWeight('bold').setFontColor(x[2]).setBackground(COR.cinza);
    if (i === 3) n.setNumberFormat('0%');
  });
  sh.setRowHeight(5, 46);

  const cab = ['ENDEREÇO', 'FALTANTES', 'ACHADOS', 'RESTAM', 'STATUS', 'ROTAS'];
  sh.getRange(7, 1, 1, 6).setValues([cab]);
  estiloCab_(sh, 6, 7);
  if (L.length) {
    const rows = L.map(x => [x.g, x.tot, x.ach, x.rst, status_(B, x), x.rotas.slice(0, 6).join(' · ') + (x.rotas.length > 6 ? ' +' + (x.rotas.length - 6) : '')]);
    const rg = sh.getRange(8, 1, rows.length, 6);
    rg.setValues(rows).setVerticalAlignment('middle');
    rg.setBackgrounds(L.map(x => Array(6).fill(x.rst === 0 ? COR.verde : emAud_(B, x.g) ? COR.amarelo : '#ffffff')));
    sh.getRange(8, 4, rows.length, 1).setFontWeight('bold').setFontSize(14)
      .setFontColors(L.map(x => [x.rst === 0 ? COR.verdeT : COR.vermelhoT]));
    sh.getRange(8, 1, rows.length, 1).setFontWeight('bold');
    sh.getRange(8, 2, rows.length, 3).setHorizontalAlignment('center');
  } else {
    sh.getRange(8, 1, 1, 6).merge().setValue(!B.falt.size ? 'Importe os faltantes na aba "Colar Faltantes".' : !B.stage.size ? 'Puxe a Stage Zone (menu 🤖 Stage Zone › Puxar Stage Zone da nuvem).' : 'Nenhum faltante nas gaiolas.')
      .setFontColor('#5b7488');
  }
  sh.setColumnWidth(1, 150); sh.setColumnWidths(2, 3, 105); sh.setColumnWidth(5, 130); sh.setColumnWidth(6, 260);
  sh.setFrozenRows(7);
  appTabs_(B, L, { tot, ach, rst, abertos, pct });
}

/* ======================= TABELAS DO APPSHEET ======================= */
function appTabs_(B, L, K) {
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM HH:mm');
  // App Endereços — 1 linha por endereço (cabeçalho na linha 1, chave = Endereço)
  const e = aba_(ABA.APP_END);
  const cabE = ['Endereço', 'Faltantes', 'Achados', 'Restam', 'Status', 'Rotas', 'Ordem', 'Atualizado'];
  const rowsE = L.map((x, i) => [x.g, x.tot, x.ach, x.rst, status_(B, x),
    x.rotas.slice(0, 6).join(' · ') + (x.rotas.length > 6 ? ' +' + (x.rotas.length - 6) : ''), i + 1, agora]);
  e.getRange(1, 1, 1, cabE.length).setValues([cabE]);
  if (rowsE.length) e.getRange(2, 1, rowsE.length, cabE.length).setValues(rowsE);
  limparSobra_(e, rowsE.length + 1, cabE.length);

  // App Reporte — 1 linha só (chave = "atual"), com o texto pronto pro grupo
  const aud = L.filter(x => x.rst === 0), abertos = L.filter(x => x.rst > 0);
  let t = '📦 *FALTANTES STAGE ZONE*' + (B.dataStage ? ' — ' + B.dataStage : '') + '\n🕒 Atualizado ' + agora + '\n\n' +
    '🔴 Faltam passar: *' + K.rst + '* em *' + abertos.length + '* endereço' + (abertos.length === 1 ? '' : 's') + '\n' +
    '✅ Encontrados no HH: *' + K.ach + '* (' + Math.round(K.pct * 100) + '%)\n' +
    (aud.length ? '🟢 Gaiolas auditadas: *' + aud.length + '*\n' : '') +
    (abertos.filter(x => emAud_(B, x.g)).length ? '🟡 Em auditoria agora: *' + abertos.filter(x => emAud_(B, x.g)).length + '*\n' : '') + '\n' +
    (abertos.length ? '*Endereços para buscar:*\n' + abertos.map(x => (emAud_(B, x.g) ? '🟡' : '▪') + ' *' + x.g + '* → ' + x.rst + ' faltante' + (x.rst > 1 ? 's' : '') + (x.ach ? ' _(' + x.ach + ' já achado' + (x.ach > 1 ? 's' : '') + ')_' : '')).join('\n')
      : '🎉 Todos os faltantes das gaiolas foram encontrados!') +
    (aud.length ? '\n\n*Gaiolas auditadas:*\n' + aud.map(x => '🟢 *' + x.g + '* (' + x.ach + ' achado' + (x.ach > 1 ? 's' : '') + ')').join('\n') : '');
  const r = aba_(ABA.APP_REP);
  const cabR = ['Chave', 'Faltam', 'Encontrados', 'Concluído', 'Endereços abertos', 'Auditadas', 'Atualizado', 'Texto'];
  r.getRange(1, 1, 2, cabR.length).setValues([cabR, ['atual', K.rst, K.ach, K.pct, abertos.length, aud.length, agora, t]]);
  r.getRange(2, 4).setNumberFormat('0%');
}

function lista_(B, achados) {
  const sh = aba_(ABA.LISTA);
  const rows = [];
  B.porEnd.forEach((ids, g) => ids.forEach(id => {
    const a = achados.get(id);
    rows.push([g, id, B.falt.get(id).rota, a ? '✅ Achado' : '🔴 Procurando', a ? (a instanceof Date ? Utilities.formatDate(a, 'America/Sao_Paulo', 'dd/MM HH:mm:ss') : String(a)) : '', B.stage.get(id).h, emAud_(B, g) ? 'SIM' : '']);
  }));
  rows.sort((x, y) => (x[3] > y[3] ? -1 : x[3] < y[3] ? 1 : 0) || x[0].localeCompare(y[0], 'pt', { numeric: true }));
  const cab = ['ENDEREÇO', 'ID DO PACOTE', 'ROTA', 'SITUAÇÃO', 'ACHADO EM', 'BIPADO NA GAIOLA', 'EM AUDITORIA'];
  sh.getRange(1, 1, 1, 7).setValues([cab]); estiloCab_(sh, 7);
  if (rows.length) {
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, 7).setValues(rows)
      .setBackgrounds(rows.map(r => Array(7).fill(r[3].indexOf('✅') === 0 ? COR.verde : r[6] ? COR.amarelo : '#ffffff')));
  }
  limparSobra_(sh, rows.length + 1, 7);
  sh.setFrozenRows(1); sh.setColumnWidths(1, 7, 140);
}

/* ======================= STAGE ZONE (nuvem) ======================= */
function puxarStage_() {
  const r = UrlFetchApp.fetch(SB_URL + '?id=eq.atual&select=data_ref,payload,atualizado_em', {
    headers: { apikey: SB_KEY, Authorization: 'Bearer ' + SB_KEY }, muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
  const row = JSON.parse(r.getContentText())[0];
  if (!row) throw new Error('ainda não tem Stage Zone salva na nuvem — suba a planilha no app Stage Zone uma vez.');
  const last = new Map();
  String(row.payload).split('\n').forEach(l => {
    const p = l.split('\t'); if (p.length < 3) return;
    const id = normId_(p[1]), g = normCage_(p[2]); if (id && g) last.set(id, [g, id, p[3] || '']);
  });
  const rows = [...last.values()];
  const sh = aba_(ABA.STAGE);
  sh.clear();
  sh.getRange(1, 1, 1, 3).setValues([['ENDEREÇO', 'ID DO PACOTE', 'HORA BIPADA']]); estiloCab_(sh, 3);
  if (rows.length) { sh.getRange(2, 2, rows.length, 1).setNumberFormat('@'); sh.getRange(2, 1, rows.length, 3).setValues(rows); }
  sh.setFrozenRows(1);
  const d = String(row.data_ref).split('-').reverse().join('/');
  const pr = PropertiesService.getDocumentProperties();
  pr.setProperty('dataStage', d); pr.setProperty('dataRef', String(row.data_ref));
  return rows.length;
}

/* ======================= UTILITÁRIOS ======================= */
// mesmas regras do app Stage Zone: aceita {"id":"48...","t":"lm"}, ignora QR de gaiola e chave de NF (44 dígitos)
function normId_(v) {
  v = String(v == null ? '' : v).trim();
  if (/"name"\s*:/.test(v) || /^\W*SR\d*-\d/i.test(v)) return '';
  let k = v.match(/^\D*(\d{8,15})\s*SR\d*-\d/i); if (k) return k[1];
  let m = v.match(/^\D*(\d{8,})\s*\{/); if (m) v = m[1];
  m = v.match(/"id"\s*:\s*"?(\d{8,})/); if (m && !/^\d/.test(v)) return m[1];
  let d = v.replace(/\D/g, '');
  if (d.length > 15 && d.length % 2 === 0 && d.slice(0, d.length / 2) === d.slice(d.length / 2)) d = d.slice(0, d.length / 2);
  return d.length >= 8 && d.length <= 15 ? d : '';
}
function normCage_(v) {
  v = String(v || '');
  const m = v.match(/"name"\s*:\s*"([^"]+)"/);
  let c = (m ? m[1] : v).toUpperCase().replace(/[^A-Z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  if (/^SR/.test(c)) c = c.replace(/[A-Z]+$/, '').replace(/-+$/, '');
  return /^SR\d+-\d+-\d+-\d{1,3}$/.test(c) ? c : '';
}
// apaga só as linhas que sobraram embaixo (em vez de limpar a aba inteira antes de escrever)
function limparSobra_(sh, ultimaUsada, ncol) {
  const fim = sh.getLastRow();
  if (fim > ultimaUsada) sh.getRange(ultimaUsada + 1, 1, fim - ultimaUsada, Math.max(ncol, sh.getLastColumn())).clearContent().setBackground(null);
}
function aba_(nome) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(nome) || ss.insertSheet(nome);
}
function estiloCab_(sh, ncol, linha) {
  sh.getRange(linha || 1, 1, 1, ncol).setBackground(COR.azul).setFontColor('#ffffff').setFontWeight('bold');
}
