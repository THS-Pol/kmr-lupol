// ============================================================
// KMR Statistik — Google Apps Script
// ============================================================
// Setup:
//  1. Google Sheets öffnen → Erweiterungen → Apps Script
//  2. Diesen Code einfügen, speichern
//  3. Bereitstellen → Neue Bereitstellung → Web-App
//     - Ausführen als: Ich (Thierry)
//     - Zugriff:      Jeder
//  4. URL kopieren → in KMR-App als SHEETS_URL eintragen
// ============================================================

const UEBERSICHT = 'Übersicht';

// ── ENTRY POINTS ────────────────────────────────────────────

function doPost(e) {
  try {
    let data;
    // 1. Iframe-Form POST: payload als Form-Feld
    if (e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else {
      // 2. Plain text body
      try { data = JSON.parse(e.postData.contents); }
      catch (_) { data = JSON.parse(e.parameter.data || '{}'); }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'saveDraft') {
      saveDraft(ss, data);
      return ok('Entwurf gespeichert');
    }
    if (data.action === 'savePyro') {
      savePyro(ss, data);
      return ok('Pyro gespeichert');
    }

    // Default: full KMR export
    updateUebersicht(ss, data);
    createMatchSheet(ss, data);
    return ok('Daten gespeichert');
  } catch (err) {
    return err_(err.toString());
  }
}

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'saveDraft') {
    try {
      const data = JSON.parse(e.parameter.data || '{}');
      if (!data.datum) data.datum = e.parameter.datum || '';
      saveDraft(ss, data);
      return ok('Entwurf gespeichert');
    } catch(err) {
      return err_('saveDraft Fehler: ' + err.toString());
    }
  }
  if (action === 'saveDraftField') {
    try {
      const datum = e.parameter.datum || '';
      const field = e.parameter.field || '';
      const data  = e.parameter.data  || '';
      saveDraftField(ss, datum, field, data);
      return ok('Gespeichert');
    } catch(err) {
      return err_('saveDraftField Fehler: ' + err.toString());
    }
  }
  if (action === 'loadDraft') {
    return loadDraft(ss, e.parameter.date);
  }
  if (action === 'getHistory') {
    return getHistory(ss, e.parameter.heim, e.parameter.gast);
  }
  if (action === 'loadPyro') {
    return loadPyro(ss, e.parameter.key);
  }
  if (action === 'addPyroEvent') {
    try {
      const event = JSON.parse(e.parameter.event || '{}');
      addPyroEvent(ss, e.parameter.key, event);
      return ok('Event gespeichert');
    } catch(err) { return err_('addPyroEvent: '+err.toString()); }
  }
  if (action === 'saveGame') {
    saveGameState(ss, e.parameter.session, e.parameter.state);
    return ok('OK');
  }
  if (action === 'loadGame') {
    return loadGameState(ss, e.parameter.session);
  }

  return ok('KMR Statistik API aktiv ✓');
}

// Helper: Zellwert sicher als ISO-Datum-String lesen
function cellToDateStr(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(val||'').trim();
}

// ── DRAFT SAVE (Field-by-Field) ──────────────────────────────
function saveDraftField(ss, datum, field, data) {
  let sh = ss.getSheetByName('Entwürfe');
  if (!sh) {
    sh = ss.insertSheet('Entwürfe');
    sh.getRange(1,1,1,4).setValues([['Datum','Feld','Gespeichert am','Daten (JSON)']])
      .setBackground('#2c3e50').setFontColor('#fff').setFontWeight('bold');
    sh.setColumnWidth(1,100).setColumnWidth(2,100).setColumnWidth(3,140).setColumnWidth(4,600);
    sh.setTabColor('#e67e22');
    // Spalte A als Text formatieren (verhindert Datumskonvertierung)
    sh.getRange('A:A').setNumberFormat('@');
  }
  const now = new Date().toLocaleString('de-CH');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const rows = sh.getRange(2,1,lastRow-1,2).getValues();
    for (let i=0; i<rows.length; i++) {
      if (cellToDateStr(rows[i][0])===datum && String(rows[i][1])===field) {
        sh.getRange(i+2,3,1,2).setValues([[now, data]]);
        return;
      }
    }
  }
  sh.appendRow([datum, field, now, data]);
}

// ── DRAFT LOAD (reassemble chunks) ──────────────────────────
function loadDraft(ss, date) {
  const sh = ss.getSheetByName('Entwürfe');
  if (!sh) return err_('Kein Entwurf-Sheet vorhanden');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return err_('Keine Entwürfe gespeichert');

  const numCols = Math.min(sh.getLastColumn(), 4);
  const rows = sh.getRange(2, 1, lastRow-1, numCols).getValues();

  const knownFields = ['main','spots','meta'];
  const isNewFormat = rows.some(r => cellToDateStr(r[0])===date && knownFields.includes(String(r[1])));

  if (isNewFormat) {
    const fields = {};
    rows.forEach(r => {
      if (cellToDateStr(r[0])===date && r[1] && r[numCols-1]) {
        fields[String(r[1])] = String(r[numCols-1]);
      }
    });
    if (fields['main']) {
      try {
        const mainData  = JSON.parse(fields['main']);
        const spotsArr  = fields['spots'] ? JSON.parse(fields['spots']) : [];
        const meta      = fields['meta']  ? JSON.parse(fields['meta'])  : {};
        const persCount = meta.persCount  || 0;
        const persons   = [];
        for (let i=0; i<persCount; i++) {
          if (fields['person_'+i]) persons.push(JSON.parse(fields['person_'+i]));
        }
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          data: {...mainData, spotterEinsatz: spotsArr, persons}
        })).setMimeType(ContentService.MimeType.JSON);
      } catch(err) { return err_('Fehler: '+err.toString()); }
    }
  }

  // Altes Format (3 Spalten)
  for (let i=rows.length-1; i>=0; i--) {
    if (cellToDateStr(rows[i][0])===date) {
      try {
        const jsonStr = String(rows[i][numCols-1]);
        const data = JSON.parse(jsonStr);
        if (data && typeof data === 'object') {
          return ContentService.createTextOutput(JSON.stringify({success:true, data}))
            .setMimeType(ContentService.MimeType.JSON);
        }
      } catch(e) {}
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    success:false, message:'Kein Entwurf für '+date
  })).setMimeType(ContentService.MimeType.JSON);
}

// ── HISTORY LOOKUP (letzte gleiche Paarung) ──────────────────
function getHistory(ss, heim, gast) {
  const sh = ss.getSheetByName(UEBERSICHT);
  if (!sh || sh.getLastRow() < 2)
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Keine Daten'}))
      .setMimeType(ContentService.MimeType.JSON);

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const data    = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();

  const hi = headers.indexOf('Heimteam');
  const gi = headers.indexOf('Gastteam');
  const get = (row, hdr) => { const i=headers.indexOf(hdr); return i>=0?row[i]:''; };

  // Search backwards for last match with same pairing
  let found = null;
  for (let i = data.length-1; i >= 0; i--) {
    const rowHeim = String(data[i][hi]||'').toLowerCase().trim();
    const rowGast = String(data[i][gi]||'').toLowerCase().trim();
    if (rowHeim === (heim||'').toLowerCase().trim() &&
        rowGast === (gast||'').toLowerCase().trim()) {
      found = data[i];
      break;
    }
  }

  if (!found)
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Keine frühere Begegnung gefunden'}))
      .setMimeType(ContentService.MimeType.JSON);

  // Map sheet columns → app state keys
  const result = {
    hFans:  get(found,'Heim-Fans'),
    gFans:  get(found,'Gäste-Fans'),
    riskH:  get(found,'Risk-Fans Heim'),
    riskG:  get(found,'Risk-Fans Gast'),
    anzS:   get(found,'Spotter'),
    anzP:   get(found,'Polizeikräfte gesamt'),
    anzSi:  get(found,'Private Si-Kräfte'),
    anzR:   get(found,'Rechtshilfe Szenenkenner'),
    fzAnz:  get(found,'Fahrzeuge Spotter'),
    ergebnis: get(found,'Ergebnis'),
  };

  return ContentService.createTextOutput(JSON.stringify({
    success:true, data:result, matchDate: String(get(found,'Datum'))
  })).setMimeType(ContentService.MimeType.JSON);
}

const ok  = msg => out({success: true,  message: msg});
const err_= msg => out({success: false, error:   msg});
const out = obj => ContentService
  .createTextOutput(JSON.stringify(obj))
  .setMimeType(ContentService.MimeType.JSON);

// ── PYRO COUNTER SYNC ────────────────────────────────────────
function addPyroEvent(ss, key, event) {
  let sh = ss.getSheetByName('PyroSync');
  if (!sh) {
    sh = ss.insertSheet('PyroSync');
    sh.getRange(1,1,1,4).setValues([['Key','Event-ID','Erfasst am','Event JSON']])
      .setBackground('#2c3e50').setFontColor('#fff').setFontWeight('bold');
    sh.setColumnWidth(1,150).setColumnWidth(2,160).setColumnWidth(3,140).setColumnWidth(4,600);
    sh.setTabColor('#e74c3c');
  }
  // Update existing event (for dupl flag changes) or append new
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const ids = sh.getRange(2,2,lastRow-1,1).getValues();
    for (let i=0; i<ids.length; i++) {
      if (String(ids[i][0]) === event.id) {
        sh.getRange(i+2, 3, 1, 2).setValues([[new Date().toLocaleString('de-CH'), JSON.stringify(event)]]);
        return;
      }
    }
  }
  sh.appendRow([key, event.id, new Date().toLocaleString('de-CH'), JSON.stringify(event)]);
}

function loadPyro(ss, key) {
  const sh = ss.getSheetByName('PyroSync');
  if (!sh || sh.getLastRow() < 2)
    return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.TEXT);
  const rows = sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
  const events = rows
    .filter(r => String(r[0]) === key)
    .map(r => { try { return JSON.parse(String(r[3])); } catch(e) { return null; } })
    .filter(Boolean);
  return ContentService.createTextOutput(JSON.stringify(events))
    .setMimeType(ContentService.MimeType.TEXT);
}

function saveGameState(ss, session, stateJson) {
  let sh = ss.getSheetByName('QuizGame');
  if (!sh) {
    sh = ss.insertSheet('QuizGame');
    sh.appendRow(['Session','State','Updated']);
    sh.setTabColor('#e74c3c');
  }
  const now = new Date().toLocaleString('de-CH');
  const last = sh.getLastRow();
  if (last > 1) {
    const rows = sh.getRange(2,1,last-1,1).getValues();
    for (let i=0; i<rows.length; i++) {
      if (String(rows[i][0])===session) {
        sh.getRange(i+2,2,1,2).setValues([[stateJson, now]]);
        return;
      }
    }
  }
  sh.appendRow([session, stateJson, now]);
}

function loadGameState(ss, session) {
  const sh = ss.getSheetByName('QuizGame');
  if (!sh || sh.getLastRow() < 2)
    return ContentService.createTextOutput('null').setMimeType(ContentService.MimeType.TEXT);
  const rows = sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  for (const row of rows) {
    if (String(row[0])===session)
      return ContentService.createTextOutput(String(row[1])).setMimeType(ContentService.MimeType.TEXT);
  }
  return ContentService.createTextOutput('null').setMimeType(ContentService.MimeType.TEXT);
}
  const sh = ss.getSheetByName('PyroSync');
  if (!sh || sh.getLastRow() < 2) return;
  const rows = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  const toDelete = [];
  for (let i=rows.length-1; i>=0; i--) {
    if (String(rows[i][0]) === key) toDelete.push(i+2);
  }
  toDelete.forEach(r => sh.deleteRow(r));
}

// ── ÜBERSICHT (Hauptblatt) ───────────────────────────────────

const HEADERS = [
  'Ereignis-ID','Datum','Wochentag','Anpfiff','Liga',
  'Heimteam','Gastteam','Ergebnis',
  'Zuschauer Total','Heim-Fans','Gäste-Fans','Risk-Fans Heim','Risk-Fans Gast',
  'PW Personen','PW Fahrzeuge',
  'Car Personen','Car Fahrzeuge','Car Angaben',
  'Zug Personen','Zug Verbindungen',
  'Spotter','Polizeikräfte inkl. Spotter','Rechtshilfe Szenenkenner','Total EZ Spotter','Einsatzleiter Spotter',
  'Personenkontrollen',
  'Festnahmen Heim','Festn. Heim Typ','Festnahmen Gast','Festn. Gast Typ',
  'Anzeigen Heim','Anz. Heim Art','Anzeigen Gast','Anz. Gast Art',
  'Schäden Personen Anzahl','Schäden Personen Beschr.',
  'Schäden Sachen Anzahl','Schäden Sachen CHF','Schäden Sachen Beschr.',
  'Schäden Polizei Anzahl','Schäden Polizei Beschr.',
  'Pyro HLF Heim','Pyro Rauch Heim','Pyro FW Heim','Pyro Böller Heim','Pyro Blitzer Heim','Pyro Total Heim',
  'Pyro HLF Gast','Pyro Rauch Gast','Pyro FW Gast','Pyro Böller Gast','Pyro Blitzer Gast','Pyro Total Gast',
  'Pyro Total',
  'Einsatz-Bewertung','Spotter-Rollen',
  'Erstellt durch','Ereignis-ID','Erfasst am'
];

function updateUebersicht(ss, d) {
  let sh = ss.getSheetByName(UEBERSICHT);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(UEBERSICHT);
  }

  // Immer Kopfzeile schreiben/aktualisieren (ohne Daten zu löschen)
  const firstCell = sh.getLastRow() > 0 ? String(sh.getRange(1,1).getValue()) : '';
  if (firstCell !== 'Ereignis-ID') {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS])
      .setBackground('#1a3560').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
    sh.setFrozenRows(1);
    sh.setRowHeight(1,32);
    sh.getRange(1,1).setNote('Automatisch befüllt durch KMR-App (Luzerner Polizei)');
  }

  const pyro = d.pyroEvents || [];
  const pc = (team, type, phases) => pyro.filter(e =>
    !e.dupl && e.team===team && e.type===type && (!phases||phases.includes(e.phase))
  ).length;

  // Manual pyro fallback
  const mpc = (team, type) => {
    const mp = (d.mPyro||{})[team]||{};
    return Object.values(mp).reduce((s,ph)=>s+(ph[type]||0),0);
  };
  const pct = (team, type) => pc(team,type,null) || mpc(team,type);

  const spots  = d.spotterEinsatz || [];
  const elSpot = spots.find(s=>s.el);
  const elName = elSpot ? `${elSpot.id} ${elSpot.n}` : '';
  const erster = spots[0] ? `${spots[0].id} ${spots[0].n}` : (d.eName||'');
  const rollenSummary = spots.filter(s=>s.rolle).map(s=>`${s.id}:${s.rolle}`).join(', ');

  const totZ   = (parseInt(d.hFans)||0)+(parseInt(d.gFans)||0);
  const datum  = d.datum||'';
  const wochentag = datum ? ['So','Mo','Di','Mi','Do','Fr','Sa'][new Date(datum+'T12:00:00').getDay()] : '';
  const ligaLabel = d.liga==='Andere'?(d.ligaAndere||'Andere'):(d.liga||'');

  const pHeim = ['HLF','RAUCH','FW','BOEL','BLITZ'].map(t=>pct('HEIM',t));
  const pGast = ['HLF','RAUCH','FW','BOEL','BLITZ'].map(t=>pct('GAST',t));
  const pHTot = pHeim.reduce((a,b)=>a+b,0);
  const pGTot = pGast.reduce((a,b)=>a+b,0);

  const bewertung = {
    'PROBLEMLOS':'Ruhiger Verlauf',
    'MITTLERE_PROBLEME':'Zwischenfälle',
    'SCHWERWIEGENDE_PROBLEME':'Gravierender Zwischenfall'
  }[d.einsatzBewertung||''] || d.einsatzBewertung||'';

  const row = [
    d.ereignisId||'', datum, wochentag, d.uhrzeit||'', ligaLabel,
    d.heim||'', d.gast||'',
    (d.tH!==''&&d.tG!=='') ? `${d.tH}:${d.tG}` : '',
    totZ||'', d.hFans||'', d.gFans||'', d.riskH||'', d.riskG||'',
    d.pwP||'', d.pwF||'',
    d.carP||'', d.carF||'', d.carI||'',
    d.zugP||'', d.zugZ||'',
    d.anzS||'', d.anzP||'', d.anzR||'', d.totEZ||'', elName,
    d.pkJ ? (d.pkAnz||'Ja') : 'Nein',
    d.fhAnz||0, [d.fhPG?'PolG':'',d.fhSP?'StPO':''].filter(Boolean).join('/')||'',
    d.fgAnz||0, [d.fgPG?'PolG':'',d.fgSP?'StPO':''].filter(Boolean).join('/')||'',
    d.azHAnz||0, d.azHArt||'', d.azGAnz||0, d.azGArt||'',
    d.dPJ?(d.dPAnz||'Ja'):'', d.dPB||'',
    d.dSJ?(d.dSAnz||'Ja'):'', d.dSJ?(d.dSSm||''):'', d.dSB||'',
    d.dPoJ?(d.dPoAnz||'Ja'):'', d.dPoB||'',
    ...pHeim, pHTot,
    ...pGast, pGTot,
    pHTot+pGTot,
    bewertung, rollenSummary,
    erster, d.ereignisId||'',
    new Date().toLocaleDateString('de-CH')
  ];

  const last = sh.getLastRow();
  let existRow = -1;
  if (last > 1) {
    const vals = sh.getRange(2,1,last-1,6).getValues();
    vals.forEach((r,i) => { if (cellToDateStr(r[1])===datum && String(r[4])===d.liga && String(r[5])===d.heim) existRow=i+2; });
  }

  if (existRow > 0) {
    sh.getRange(existRow,1,1,row.length).setValues([row]);
    colorRow(sh,existRow,row.length);
  } else {
    sh.appendRow(row);
    colorRow(sh,sh.getLastRow(),row.length);
  }
}

function colorRow(sh, r, len) {
  const bg = r%2===0 ? '#f4f7fb' : '#ffffff';
  sh.getRange(r,1,1,len).setBackground(bg).setFontSize(10).setVerticalAlignment('middle');
}

function initUebersicht(sh) {
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS])
    .setBackground('#1a3560').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setRowHeight(1,32);
  sh.getRange(1,1).setNote('Automatisch befüllt durch KMR-App (Luzerner Polizei)');
}

// ── MATCH SHEET (pro Spiel) ──────────────────────────────────

function createMatchSheet(ss, d) {
  const name = `${d.datum||'?'} ${(d.heim||'').substring(0,6)}-${(d.gast||'Gast').substring(0,6)}`;
  const safeName = name.replace(/[\\\/\?\*\[\]:]/g, '').substring(0, 50).trim();

  const old = ss.getSheetByName(safeName);
  if (old) ss.deleteSheet(old);

  const sh = ss.insertSheet(safeName);

  // Tab color based on team
  sh.setTabColor('#1a3560');

  let r = 1;
  r = writeTitle(sh, r, d);
  r = writeSection(sh, r, 'Spieldaten', spieldatenRows(d));
  r = writeSpotter(sh, r, d.spotterEinsatz || []);
  r = writeSection(sh, r, 'Zuschauer & Anreise', zuschauserRows(d));
  r = writeSection(sh, r, 'Personalaufgebot', personalRows(d));
  r = writeSachverhalt(sh, r, d);
  r = writeSection(sh, r, 'Polizeieinsatz & Einsatzmittel', einsatzRows(d));
  r = writeSzenenverhalten(sh, r, d);
  r = writePyro(sh, r, d);
  r = writeSection(sh, r, 'Anhaltungen, Anzeigen & Massnahmen', kontrollenRows(d));
  r = writeSection(sh, r, 'Verursachte Schäden', schaedenRows(d));
  r = writePersonen(sh, r, d.persons || []);
  r = writeAbschluss(sh, r, d);

  formatMatchSheet(sh);

  // Move after overview
  ss.moveActiveSheet(ss.getSheets().length);
}

function writeTitle(sh, r, d) {
  const title = `KMR — ${d.heim||''} vs. ${d.gast||''} — ${d.datum||''}  [${d.tH}:${d.tG}]  ${d.liga||''}`;
  sh.getRange(r, 1, 1, 6).merge().setValue(title)
    .setBackground('#0f1a2e').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sh.setRowHeight(r, 36);
  return r + 1;
}

function writeSection(sh, r, title, rows) {
  sh.getRange(r, 1, 1, 6).merge().setValue(title)
    .setBackground('#1a3560').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(r, 26);
  r++;
  rows.forEach(row => {
    if (row === null) { r++; return; }
    const range = sh.getRange(r, 1, 1, row.length);
    range.setValues([row]);
    // Alternate key/value styling
    for (let c = 0; c < row.length; c++) {
      if (c % 2 === 0 && row[c] !== '') {
        sh.getRange(r, c+1).setFontWeight('bold').setFontColor('#1a3560');
      }
    }
    r++;
  });
  return r + 1;
}

function writeSachverhalt(sh, r, d) {
  sh.getRange(r, 1, 1, 6).merge().setValue('Sachverhalt')
    .setBackground('#1a3560').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(r, 26); r++;
  [['⬆ Aufmarsch', d.aufm], ['⚽ Spielphase', d.spiel], ['⬇ Abmarsch', d.abm]].forEach(([lbl, txt]) => {
    if (!txt) return;
    sh.getRange(r, 1).setValue(lbl).setFontWeight('bold').setFontColor('#1a3560');
    sh.getRange(r, 2, 1, 5).merge().setValue(txt).setWrap(true);
    sh.setRowHeight(r, Math.max(40, Math.ceil(txt.length / 80) * 18));
    r++;
  });
  return r + 1;
}

function writeSzenenverhalten(sh, r, d) {
  sh.getRange(r, 1, 1, 6).merge().setValue('Szenenverhalten')
    .setBackground('#1a3560').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(r, 26); r++;
  sh.getRange(r, 1, 1, 6).setValues([['Team','Alkohol','Aufmarsch','Spielphase','Abmarsch','']])
    .setFontWeight('bold').setBackground('#e8eef7').setFontColor('#1a3560');
  r++;
  [[d.heim||'Heim','h'],[d.gast||'Gast','g']].forEach(([name, p]) => {
    sh.getRange(r, 1, 1, 6).setValues([[
      name,
      d[p+'Alk'] ? 'Ja':'Nein',
      d[p+'Auf'] ? 'Normal':'Aggressiv',
      d[p+'Sp']  ? 'Normal':'Aggressiv',
      d[p+'Ab']  ? 'Normal':'Aggressiv',
      ''
    ]]);
    if (!d[p+'Auf'] || !d[p+'Sp'] || !d[p+'Ab']) {
      sh.getRange(r, 1, 1, 6).setBackground('#ffeaea');
    }
    r++;
  });
  return r + 1;
}

function writePyro(sh, r, d) {
  const pyro = d.pyroEvents || [];
  const pc = (team, type, phases) => pyro.filter(e =>
    !e.dupl && e.team===team && e.type===type && (!phases || phases.includes(e.phase))
  ).length;

  sh.getRange(r, 1, 1, 6).merge().setValue('Pyrotechnik')
    .setBackground('#1a3560').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(r, 26); r++;

  sh.getRange(r, 1, 1, 6).setValues([['Team / Phase','','HLF','Rauch','Blitzer','Feuerwerk']])
    .setFontWeight('bold').setBackground('#e8eef7').setFontColor('#1a3560');
  r++;

  const phases = [
    ['Aufmarsch',['AUFM']], ['1. Halbzeit',['HZ1']],
    ['Pause',['PAUSE']],   ['2. Halbzeit',['HZ2']],
    ['Verlängerung',['VL']],['Abmarsch',['ABML']],
    ['TOTAL', null]
  ];

  ['HEIM','GAST'].forEach(team => {
    const tName = team==='HEIM' ? (d.heim||'Heim') : (d.gast||'Gast');
    const tColor = team==='HEIM' ? '#e8f5e9' : '#fde8e8';
    phases.forEach(([pLbl, pIds]) => {
      const vals = [pLbl==='TOTAL'?tName:'' , pLbl, pc(team,'HLF',pIds), pc(team,'RAUCH',pIds), pc(team,'BLITZ',pIds), pc(team,'FW',pIds)];
      sh.getRange(r, 1, 1, 6).setValues([vals]);
      if (pLbl === 'TOTAL') {
        sh.getRange(r, 1, 1, 6).setBackground(tColor).setFontWeight('bold');
      }
      r++;
    });
  });

  // Timeline (if pyro events exist)
  if (pyro.length > 0) {
    r++;
    sh.getRange(r, 1, 1, 6).merge().setValue('Pyro-Timeline (chronologisch)')
      .setBackground('#2c3e50').setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
    sh.setRowHeight(r, 22); r++;
    sh.getRange(r, 1, 1, 6).setValues([['Uhrzeit','Team','Typ','Phase','Minute','Spotter']])
      .setFontWeight('bold').setBackground('#e8eef7').setFontColor('#1a3560');
    r++;
    pyro.filter(e=>!e.dupl).sort((a,b)=>a.ts-b.ts).forEach(ev => {
      const t = new Date(ev.ts).toLocaleTimeString('de-CH',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
      const tColor = ev.team==='HEIM' ? '#e8f5e9' : '#fde8e8';
      sh.getRange(r, 1, 1, 6).setValues([[t, ev.team, ev.type, ev.phase, ev.minute!=='—'?ev.minute+"'":'—', ev.spotter]])
        .setBackground(tColor);
      r++;
    });
  }
  return r + 1;
}

function writeSpotter(sh, r, spots) {
  sh.getRange(r, 1, 1, 6).merge().setValue('Spotter im Einsatz')
    .setBackground('#1a3560').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(r, 26); r++;
  sh.getRange(r, 1, 1, 5).setValues([['LU-Nr.','Name','Von','Bis','Dauer']])
    .setFontWeight('bold').setBackground('#e8eef7').setFontColor('#1a3560');
  r++;
  if (!spots.length) { sh.getRange(r, 1).setValue('—'); r++; }
  spots.forEach(sp => {
    sh.getRange(r, 1, 1, 5).setValues([[sp.id||'', sp.n||'', sp.von||'', sp.bis||'', calcH_(sp.von, sp.bis)]]);
    r++;
  });
  return r + 1;
}

function writePersonen(sh, r, persons) {
  if (!persons.length) return r;
  sh.getRange(r, 1, 1, 6).merge().setValue(`Kontrollierte Personen (${persons.length})`)
    .setBackground('#1a3560').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(r, 26); r++;

  persons.forEach((p, i) => {
    sh.getRange(r, 1, 1, 6).setValues([[`Person ${i+1}`, '', '', '', '', '']])
      .setBackground('#e8eef7').setFontWeight('bold').setFontColor('#1a3560');
    r++;
    [
      ['Name', p.nm||'', 'Vorname', p.vnm||'', 'Geburtsdatum', p.geb||''],
      ['Heimatort', p.ho||'', 'Beruf', p.beruf||'', 'Adresse', p.adr||''],
      ['Telefon', p.tel||'', 'Fanzugehörigkeit', p.fanz||'', 'Kontrollort', p.ort||''],
      ['Kontrollzeit', p.zeit||'', 'Kontrolliert durch', p.durch||'', '', ''],
    ].forEach(rowVals => {
      sh.getRange(r, 1, 1, 6).setValues([rowVals]);
      for (let c = 0; c < 6; c+=2) { if (rowVals[c]) sh.getRange(r,c+1).setFontWeight('bold').setFontColor('#1a3560'); }
      r++;
    });

    const grund = [['gVM','Vermummung'],['gPR','Provokation'],['gSB','Sachbeschädigung'],['gGW','Gewalt g. Polizei'],['gSL','Schlägerei/KV'],['gPV','Präventiv'],['gPY','Pyro']].filter(([k])=>p[k]).map(([,v])=>v);
    const mass  = [['mPK','Pers.K.'],['mAZ','Anzeige'],['mFP','Festnahme PolG'],['mFS','Festnahme StPO'],['mWG','Wegweisung'],['mSV','SV'],['mRV','RV']].filter(([k])=>p[k]).map(([,v])=>v);
    sh.getRange(r, 1).setValue('Kontrollgrund:').setFontWeight('bold').setFontColor('#1a3560');
    sh.getRange(r, 2, 1, 5).merge().setValue(grund.join(', ')||'—');
    r++;
    sh.getRange(r, 1).setValue('Massnahme:').setFontWeight('bold').setFontColor('#1a3560');
    sh.getRange(r, 2, 1, 5).merge().setValue(mass.join(', ')||'—');
    r++;
    if (p.sv) {
      sh.getRange(r, 1).setValue('Sachverhalt:').setFontWeight('bold').setFontColor('#1a3560');
      sh.getRange(r, 2, 1, 5).merge().setValue(p.sv).setWrap(true);
      sh.setRowHeight(r, Math.max(36, Math.ceil(p.sv.length/80)*18));
      r++;
    }
    r++;
  });
  return r + 1;
}

function writeAbschluss(sh, r, d) {
  sh.getRange(r, 1, 1, 6).merge().setValue('Abschluss & Verteiler')
    .setBackground('#1a3560').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(r, 26); r++;
  [
    ['Erstellt durch', d.eName||'', 'Telefon', d.eTel||'', 'Mobile', d.eMob||''],
    ['E-Mail', d.eEml||'', 'Erstellt am', d.datum||'', '', ''],
    ['Geht an', 'Kommandant Adi Achermann · fussball.polizei@lu.ch · 2. Stv C Kripo Roland Stöckli · intervention.polizei@lu.ch', '', '', '', ''],
  ].forEach(rowVals => {
    sh.getRange(r, 1, 1, 6).setValues([rowVals]);
    sh.getRange(r, 1).setFontWeight('bold').setFontColor('#1a3560');
    r++;
  });
  return r;
}

function formatMatchSheet(sh) {
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 120);
  sh.setFrozenRows(1);
}

// ── ROW BUILDERS ─────────────────────────────────────────────

function spieldatenRows(d) {
  const totZ = (parseInt(d.hFans)||0)+(parseInt(d.gFans)||0);
  return [
    ['Liga', d.liga||'', 'Journal-Nr.', d.journalNr||'', 'Ereignis-ID', d.ereignisId||''],
    ['Datum', d.datum||'', 'Stadion', d.stadion||'', 'Kapazität', d.kap||''],
    ['Heimteam', d.heim||'', 'Ergebnis', `${d.tH||'?'}:${d.tG||'?'}`, 'Gastteam', d.gast||''],
  ];
}

function zuschauserRows(d) {
  const totZ = (parseInt(d.hFans)||0)+(parseInt(d.gFans)||0);
  const rows = [
    ['Heim-Fans', d.hFans||'', 'Gäste-Fans', d.gFans||'', 'Total', totZ||''],
    ['Risk-Fans Heim', d.riskH||'', 'Risk-Fans Gast', d.riskG||'', '', ''],
    ['Anreise PW', `${d.pwP||0} Pers. / ${d.pwF||0} Fz`, 'Car', `${d.carP||0} Pers. / ${d.carF||0} Fz`, 'Zug', `${d.zugP||0} Pers.`],
  ];
  if (d.carI) rows.push(['Car Angaben', d.carI, '', '', '', '']);
  return rows;
}

function personalRows(d) {
  const totKm = (parseInt(d.fzAnz)||0)*(parseInt(d.fzKm)||0);
  return [
    ['Spotter', d.anzS||'', 'Polizeikräfte inkl. Spotter', d.anzP||'', 'Private Si-Kräfte', d.anzSi||''],
    ['Fahrzeuge', d.fzAnz||'', 'Km pro Fahrzeug', d.fzKm||'', 'Total Km', totKm||''],
    ['Rechtshilfe', d.anzR||'', 'Total EZ Spotter', d.totEZ||'', '', ''],
    ['Spotter EZ', d.sVon?`${d.sVon}–${d.sBis} (${calcH_(d.sVon,d.sBis)})`:''  , 'Polizei EZ', d.pVon?`${d.pVon}–${d.pBis} (${calcH_(d.pVon,d.pBis)})`:'' , 'Si EZ', d.siVon?`${d.siVon}–${d.siBis} (${calcH_(d.siVon,d.siBis)})`:'' ],
  ];
}

function einsatzRows(d) {
  const emMap = {emOD:'Ordnungsdienst',emGu:'Gummi',emRZ:'Reizstoff',emPM:'PMS/TES',emWW:'Wasserwerfer',emSG:'Sperrgitter',emOB:'Observation',emHU:'Hund',emHK:'Helikopter',emSW:'Schusswaffe',emSP:'Spotter',emUP:'Uniformpolizei'};
  const active = Object.entries(emMap).filter(([k])=>d[k]).map(([,v])=>v);
  return [
    ['Präventiv', d.praev?'Ja':'Nein', 'Repressiv', d.repr?'Ja':'Nein', '', ''],
    ['Einsatzmittel', active.join(', ')||'—', '', '', '', ''],
  ];
}

function kontrollenRows(d) {
  return [
    ['Personenkontrollen', d.pkJ?(d.pkAnz||'Ja'):'Nein', '', '', '', ''],
    ['Festnahmen Heim', d.fhAnz||0, [d.fhPG?'PolG':'',d.fhSP?'StPO':''].filter(Boolean).join('/')||'—', 'Festnahmen Gast', d.fgAnz||0, [d.fgPG?'PolG':'',d.fgSP?'StPO':''].filter(Boolean).join('/')||'—'],
    ['Anzeigen Heim', d.azHAnz||0, d.azHArt||'', 'Anzeigen Gast', d.azGAnz||0, d.azGArt||''],
    ['SV', d.svJ?(d.svA||1):0, 'RV', d.rvJ?(d.rvA||1):0, 'MA', d.maJ?(d.maA||1):0],
    ['GW', d.gwJ?(d.gwA||1):0, '', '', '', ''],
  ];
}

function schaedenRows(d) {
  const rows = [
    ['Personen', d.dPJ?'Ja':'Nein', d.dPAnz||'', d.dPB||'', '', ''],
    ['Sachen', d.dSJ?'Ja':'Nein', d.dSSm?`CHF ${d.dSSm}`:'Nein', '', '', ''],
    ['Polizei', d.dPoJ?'Ja':'Nein', d.dPoAnz||'', '', '', ''],
  ];
  return rows;
}

// ── HELPERS ──────────────────────────────────────────────────

function getPyroHistory(ss) {
  const sh = ss.getSheetByName(UEBERSICHT);
  if (!sh || sh.getLastRow() < 2)
    return ContentService.createTextOutput(JSON.stringify({success:true,data:[]}))
      .setMimeType(ContentService.MimeType.JSON);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const rows    = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const get = (row,hdr) => { const i=headers.indexOf(hdr); return i>=0?row[i]:''; };
  const data = rows
    .filter(r => String(get(r,'Datum')))
    .map(r => {
      const hlf  = (parseInt(get(r,'Pyro HLF Heim'))||0)+(parseInt(get(r,'Pyro HLF Gast'))||0);
      const rauch= (parseInt(get(r,'Pyro Rauch Heim'))||0)+(parseInt(get(r,'Pyro Rauch Gast'))||0);
      const fw   = (parseInt(get(r,'Pyro FW Heim'))||0)+(parseInt(get(r,'Pyro FW Gast'))||0);
      const boel = (parseInt(get(r,'Pyro Böller Heim'))||0)+(parseInt(get(r,'Pyro Böller Gast'))||0);
      const blitz= (parseInt(get(r,'Pyro Blitzer Heim'))||0)+(parseInt(get(r,'Pyro Blitzer Gast'))||0);
      return {
        datum: String(get(r,'Datum')),
        heim:  String(get(r,'Heimteam')),
        gast:  String(get(r,'Gastteam')),
        total: hlf+rauch+fw+boel+blitz
      };
    });
  return ContentService.createTextOutput(JSON.stringify({success:true,data}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ──────────────────────────────────────────────────
function calcH_(von, bis) {
  const m = t => { const [h,mn] = t.split(':').map(Number); return h*60+mn; };
  let diff = m(bis)-m(von); if(diff<0) diff+=1440;
  const h=Math.floor(diff/60), mn=diff%60;
  return `${h}h${mn?' '+mn+'min':''}`;
}
