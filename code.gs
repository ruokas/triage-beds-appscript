/**
 * ED Dashboard - Apps Script backend (final)
 * - Logs to "VEIKSMŲ ŽURNALAS" (TimeISO | User | Action | Summary | From | To | Bed | Patient | Doctor | Triage | Status | Comment)
 * - Doctors from PACIENTŲ SKIRSTYMAS!B2:I2
 * - Ambulatorija: K=Doctor, L=Triage, M=Patient, N=Comment (rows 2..16 => AMB1..AMB15)
 * - Nurses from LENTA: IT A2+A3; Z1 A4:A9; Z2 A10:A17; Z3 A18:A26; Z4 A27:A35; Z5 A36:A40; Amb J2 (merged J2:J16)
 * - Zones: IT, 1..5, Ambulatorija
 * - LockService on all writes and reservations
 * - Undo support
 * - Robust bed lookup
 */

// -------- CONFIG --------
const SH_LENTA     = 'LENTA';
const SH_LOG       = 'VEIKSMŲ ŽURNALAS';
const SH_PACIENTU  = 'PACIENTŲ SKIRSTYMAS';

const STATUS_TEXT = 'Laukia apžiūros';

const BOARD_COLUMNS = {
  BED:    1,  // A: bed label (IT1, 1, P1, S4, 121A, ...)
  DOCTOR: 2,  // B
  TRIAGE: 5,  // E
  STATUS: 7,  // G
  COMMENT: 8  // H
};
const PATIENT_COL = 3; // C (patient) on hall board

const AMB_COLS = { // on LENTA
  DOCTOR:  11, // K
  TRIAGE:  12, // L
  PATIENT: 13, // M
  COMMENT: 14  // N
};

const ZONE_LAYOUT = {
  'Zona IT': ['IT1', 'IT2'],
  'Zona 1':  ['1','2','3','P1','P2','P3'],
  'Zona 2':  ['4','5','6','P4','P5','P6','S4','S5','S6'],
  'Zona 3':  ['7','8','9','P7','P8','P9','S7','S8','S9'],
  'Zona 4':  ['10','11','12','P10','P11','P12','S10','S11','S12'],
  'Zona 5':  ['13','14','15','16','17','121A','121B','IZO'],
  'Ambulatorija': Array.from({length:15}, (_,i)=>`AMB${i+1}`)
};

// -------- MENU --------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ED valdymas')
    .addItem('Atidaryti šoninę juostą', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('triageinfosidebar')
    .setTitle('ED lenta')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showHelpDialog() {
  SpreadsheetApp.getUi().alert('Pagalba', 'Trumpas aprašas rodomas šoninėje juostoje.', SpreadsheetApp.ButtonSet.OK);
}

// -------- USERNAME (actor) --------
function getUsername() {
  return PropertiesService.getUserProperties().getProperty('ED_USERNAME') || '';
}
function setUsername(name) {
  PropertiesService.getUserProperties().setProperty('ED_USERNAME', (name || '').trim());
  return { ok: true };
}
function _resolveUser_(provided) {
  const v = String(provided || '').trim();
  if (v) return v;                        // 1) explicit name from client
  const prop = getUsername();             // 2) UserProperties
  if (prop) return prop;
  try {
    const email = Session.getActiveUser().getEmail(); // 3) email
    if (email) return email;
  } catch(e) {}
  return 'anon:' + Session.getTemporaryActiveUserKey(); // 4) fallback
}

// -------- DOCTORS --------
function getDoctors() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_PACIENTU);
  if (!sh) return [];
  const vals = sh.getRange(2, 2, 1, 8).getValues()[0]; // B2:I2
  return vals.filter(v => v && String(v).trim()).map(v => String(v).trim());
}

// -------- NURSES --------
function getNursesFromBoard_Exact() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_LENTA);
  if (!sh) return {};
  function joinBlock(a1) {
    const range = sh.getRange(a1);
    const values = range.getValues().flat().map(v => String(v || '').trim()).filter(Boolean);
    return values.join(' ');
  }
  const nurses = {};
  const a2 = String(sh.getRange('A2').getValue() || '').trim();
  const a3 = String(sh.getRange('A3').getValue() || '').trim();
  nurses['Zona IT'] = [a2, a3].filter(Boolean).join(' ');
  nurses['Zona 1'] = joinBlock('A4:A9');
  nurses['Zona 2'] = joinBlock('A10:A17');
  nurses['Zona 3'] = joinBlock('A18:A26');
  nurses['Zona 4'] = joinBlock('A27:A35');
  nurses['Zona 5'] = joinBlock('A36:A40');
  nurses['Ambulatorija'] = String(sh.getRange('J2').getDisplayValue() || '').trim(); // merged J2:J16
  return nurses;
}

// -------- SNAPSHOT --------
function buildZonesPayload_(actor) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_LENTA);
  if (!sh) throw new Error('Nerastas lapas LENTA.');

  const nurses = getNursesFromBoard_Exact();
  const occupied = getOccupiedMap_(sh);
  const reservations = getReservationMap_();

  const zones = {};
  Object.keys(ZONE_LAYOUT).forEach(z => zones[z] = { name: z, beds: [] });

  for (const [zone, labels] of Object.entries(ZONE_LAYOUT)) {
    for (const label of labels) {
      const occ = occupied[label] || null;
      const res = reservations[label] || null;
      const reservedByMe    = !!(res && actor && res.actor === actor);
      const reservedByOther = !!(res && (!actor || res.actor !== actor));

      zones[zone].beds.push({
        label,
        occupied: !!occ,
        patient: occ ? occ.patient : '',
        reservedByMe,
        reservedByOther
      });
    }
  }
  return { beds: zones, nurses };
}

function getOccupiedMap_(sh) {
  const map = {};
  const lastRow = sh.getLastRow();
  if (lastRow >= 1) {
    const labels  = sh.getRange(1, BOARD_COLUMNS.BED, lastRow, 1).getValues();
    const doctors = sh.getRange(1, BOARD_COLUMNS.DOCTOR, lastRow, 1).getValues();
    const triages = sh.getRange(1, BOARD_COLUMNS.TRIAGE, lastRow, 1).getValues();
    const statuses= sh.getRange(1, BOARD_COLUMNS.STATUS, lastRow, 1).getValues();
    const comments= sh.getRange(1, BOARD_COLUMNS.COMMENT, lastRow, 1).getValues();
    const patients= sh.getRange(1, PATIENT_COL,       lastRow, 1).getValues();
    for (let r = 1; r <= lastRow; r++) {
      const label = String(labels[r-1][0] || '').trim();
      if (!label || label.startsWith('AMB')) continue;
      const patient = String(patients[r-1][0] || '').trim();
      const occupied = !!(patient ||
                          String(doctors[r-1][0]||'').trim() ||
                          String(triages[r-1][0]||'').trim() ||
                          String(statuses[r-1][0]||'').trim() ||
                          String(comments[r-1][0]||'').trim());
      if (occupied) {
        map[label] = {
          row: r,
          patient,
          doctor:  String(doctors[r-1][0]||'').trim(),
          triage:  String(triages[r-1][0]||'').trim(),
          status:  String(statuses[r-1][0]||'').trim(),
          comment: String(comments[r-1][0]||'').trim()
        };
      }
    }
  }
  // AMB rows 2..16
  for (let i = 1; i <= 15; i++) {
    const label = `AMB${i}`;
    const row = 1 + i;
    const patient = String(sh.getRange(row, AMB_COLS.PATIENT).getValue() || '').trim();
    const doctor  = String(sh.getRange(row, AMB_COLS.DOCTOR).getValue()  || '').trim();
    const triage  = String(sh.getRange(row, AMB_COLS.TRIAGE).getValue()  || '').trim();
    const comment = String(sh.getRange(row, AMB_COLS.COMMENT).getValue() || '').trim();
    if (patient || doctor || triage || comment) {
      map[label] = { row, patient, doctor, triage, status: '', comment };
    }
  }
  return map;
}

// -------- RESERVATIONS --------
const CACHE_PREFIX = 'ED_RES_';
const RES_TTL_SEC  = 180; // 3 min

function reserveBed(payload) {
  const { actor, bedLabel } = normalizePayload_(payload);
  if (!bedLabel) return { ok:false, msg:'Nenurodyta lova' };
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const cache = CacheService.getDocumentCache();
    const key = CACHE_PREFIX + bedLabel;
    const existing = readJson(cache.get(key));
    if (existing && existing.actor && existing.actor !== actor) {
      return { ok:false, msg:'Lova jau rezervuota kito naudotojo.' };
    }
    cache.put(key, JSON.stringify({ actor, ts: Date.now() }), RES_TTL_SEC);
    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}
function keepAliveReservation(payload) {
  const { actor, bedLabel } = normalizePayload_(payload);
  if (!bedLabel) return { ok:false };
  const cache = CacheService.getDocumentCache();
  const key = CACHE_PREFIX + bedLabel;
  const existing = readJson(cache.get(key));
  if (!existing || existing.actor !== actor) return { ok:false };
  cache.put(key, JSON.stringify({ actor, ts: Date.now() }), RES_TTL_SEC);
  return { ok:true };
}
function releaseReservation(payload) {
  const { actor, bedLabel } = normalizePayload_(payload);
  if (!bedLabel) return { ok:true };
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const cache = CacheService.getDocumentCache();
    const key = CACHE_PREFIX + bedLabel;
    const existing = readJson(cache.get(key));
    if (existing && existing.actor === actor) cache.remove(key);
  } finally {
    lock.releaseLock();
  }
  return { ok:true };
}
function getReservationMap_() {
  const cache = CacheService.getDocumentCache();
  const map = {};
  Object.values(ZONE_LAYOUT).flat().forEach(label => {
    const raw = cache.get(CACHE_PREFIX + label);
    const obj = readJson(raw);
    if (obj && obj.actor) map[label] = obj;
  });
  return map;
}

// -------- ACTIONS (assign/move/swap/discharge) + UNDO --------
function assignBed(payload) {
  const p = normalizePayload_(payload);
  if (!p.bedLabel || !p.patientName || !p.doctor) return { ok:false, msg:'Trūksta laukų' };

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_LENTA);
  if (!sh) return { ok:false, msg:'Nerastas LENTA' };

  try {
    const undo = {};
    if (p.bedLabel.startsWith('AMB')) {
      const idx = Number(p.bedLabel.replace('AMB',''));
      const row = 1 + idx;
      undo.type = 'assignAMB';
      undo.row  = row;
      undo.before = sh.getRange(row, AMB_COLS.DOCTOR, 1, 4).getValues()[0];
      sh.getRange(row, AMB_COLS.DOCTOR).setValue(p.doctor);
      sh.getRange(row, AMB_COLS.TRIAGE).setValue(p.triage || '');
      sh.getRange(row, AMB_COLS.PATIENT).setValue(p.patientName);
      sh.getRange(row, AMB_COLS.COMMENT).setValue(p.comment || '');
    } else {
      const row = _rowByBedFast_(sh, p.bedLabel);
      if (!row) {
        const all = sh.getRange(1, BOARD_COLUMNS.BED, Math.max(1, sh.getLastRow()), 1).getValues().map(v=>String(v[0]||'').trim()).filter(Boolean).slice(0,200);
        return { ok:false, msg:`Nerasta lova lentoje: "${p.bedLabel}". Patikrinkite A stulpelį. Pvz.: ${all.join(', ')}` };
      }
      const width = BOARD_COLUMNS.COMMENT - BOARD_COLUMNS.DOCTOR + 1;
      undo.type = 'assignHALL';
      undo.row  = row;
      undo.before = sh.getRange(row, BOARD_COLUMNS.DOCTOR, 1, width).getValues()[0];
      sh.getRange(row, BOARD_COLUMNS.DOCTOR).setValue(p.doctor);
      sh.getRange(row, BOARD_COLUMNS.TRIAGE).setValue(p.triage);
      sh.getRange(row, BOARD_COLUMNS.STATUS).setValue(STATUS_TEXT);
      sh.getRange(row, BOARD_COLUMNS.COMMENT).setValue(p.comment || '');
      sh.getRange(row, PATIENT_COL).setValue(p.patientName);
    }
    const token = saveUndo_(undo);
    _logAction_({
      action: 'ASSIGN',
      summary: `Įrašyta į ${p.bedLabel}: ${p.patientName} (${p.doctor}${p.triage?`, T${p.triage}`:''})`,
      bed: p.bedLabel, patient: p.patientName, doctor: p.doctor, triage: p.triage,
      status: STATUS_TEXT, comment: p.comment || '',
      userDisplayName: p.userDisplayName
    });
    releaseReservation({ actor:p.actor, bedLabel:p.bedLabel });
    return { ok:true, undoToken: token };
  } finally {
    lock.releaseLock();
  }
}

function movePatient(payload) {
  const p = normalizePayload_(payload);
  if (!p.fromBed || !p.toBed) return { ok:false, msg:'Trūksta laukų' };

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH_LENTA);
    if (!sh) return { ok:false, msg:'Nerastas LENTA' };

    const snapFrom = readBedSnapshot_(sh, p.fromBed);
    const snapTo   = readBedSnapshot_(sh, p.toBed);
    if (!snapFrom || !snapFrom.occupied) return { ok:false, msg:'Šaltinio lova tuščia' };
    if (snapTo && snapTo.occupied)       return { ok:false, msg:'Tikslinė lova užimta (naudokite „Swap“).' };

    const undo = {
      type: 'move',
      from: { bed:p.fromBed, row:snapFrom.row, before:snapFrom.before },
      to:   { bed:p.toBed,   row: snapTo ? snapTo.row : guessRowFromLabel_(p.toBed),
              before: snapTo ? snapTo.before : emptyBlockForBed_(p.toBed) }
    };
    writeSnapshotToBed_(sh, p.toBed, snapFrom);
    clearBed_(sh, p.fromBed);

    const token = saveUndo_(undo);
    _logAction_({
      action: 'MOVE',
      summary: `Perkelta: ${p.fromBed} → ${p.toBed} (${snapFrom.patient || ''})`,
      from: p.fromBed, to: p.toBed, bed: p.toBed, patient: snapFrom.patient || '',
      userDisplayName: p.userDisplayName
    });
    return { ok:true, undoToken: token };
  } finally {
    lock.releaseLock();
  }
}

function swapBeds(payload) {
  const p = normalizePayload_(payload);
  if (!p.bedA || !p.bedB) return { ok:false, msg:'Trūksta laukų' };

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH_LENTA);
    if (!sh) return { ok:false, msg:'Nerastas LENTA' };

    const snapA = readBedSnapshot_(sh, p.bedA);
    const snapB = readBedSnapshot_(sh, p.bedB);
    if (!snapA || !snapA.occupied) return { ok:false, msg:'Pirmoji lova tuščia' };
    if (!snapB || !snapB.occupied) return { ok:false, msg:'Antroji lova tuščia' };

    const undo = { type:'swap',
      A:{ bed:p.bedA, row:snapA.row, before:snapA.before },
      B:{ bed:p.bedB, row:snapB.row, before:snapB.before }
    };
    writeSnapshotToBed_(sh, p.bedA, snapB);
    writeSnapshotToBed_(sh, p.bedB, snapA);

    const token = saveUndo_(undo);
    _logAction_({
      action: 'SWAP',
      summary: `Sukeista: ${p.bedA} ↔ ${p.bedB}`,
      from: p.bedA, to: p.bedB, bed: `${p.bedA}↔${p.bedB}`,
      userDisplayName: p.userDisplayName
    });
    return { ok:true, undoToken: token };
  } finally {
    lock.releaseLock();
  }
}

function dischargePatient(payload) {
  const p = normalizePayload_(payload);
  if (!p.bedLabel) return { ok:false, msg:'Nenurodyta lova' };

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH_LENTA);
    if (!sh) return { ok:false, msg:'Nerastas LENTA' };

    const snap = readBedSnapshot_(sh, p.bedLabel);
    if (!snap || !snap.occupied) return { ok:false, msg:'Lova jau tuščia' };

    const undo = { type:'discharge', bed:p.bedLabel, row:snap.row, before:snap.before };
    clearBed_(sh, p.bedLabel);

    const token = saveUndo_(undo);
    _logAction_({
      action: 'DISCHARGE',
      summary: `Išrašyta iš ${p.bedLabel}: ${snap.patient || ''}`,
      bed: p.bedLabel, patient: snap.patient || '',
      userDisplayName: p.userDisplayName
    });
    return { ok:true, undoToken: token, msg:'Pacientas išrašytas.' };
  } finally {
    lock.releaseLock();
  }
}

// -------- UNDO STORAGE --------
function saveUndo_(obj) {
  const token = 'U' + Utilities.getUuid();
  PropertiesService.getDocumentProperties().setProperty(token, JSON.stringify(obj));
  return token;
}
function undoAssign(token) {
  if (!token) return { ok:false };
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(token);
  if (!raw) return { ok:false, msg:'Nebėra ką atšaukti.' };
  props.deleteProperty(token);

  const undo = JSON.parse(raw);
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_LENTA);
  if (!sh) return { ok:false };

  switch (undo.type) {
    case 'assignAMB':
      sh.getRange(undo.row, AMB_COLS.DOCTOR, 1, 4).setValues([undo.before]);
      break;
    case 'assignHALL': {
      const width = BOARD_COLUMNS.COMMENT - BOARD_COLUMNS.DOCTOR + 1;
      sh.getRange(undo.row, BOARD_COLUMNS.DOCTOR, 1, width).setValues([undo.before]);
      break;
    }
    case 'move':
      _applyBlock_(sh, undo.to);
      _applyBlock_(sh, undo.from);
      break;
    case 'swap':
      _applyBlock_(sh, undo.A);
      _applyBlock_(sh, undo.B);
      break;
    case 'discharge':
      _applyBlock_(sh, undo);
      break;
    default:
      return { ok:false };
  }
  return { ok:true };
}
function _applyBlock_(sh, blk) {
  if (!blk || !blk.row) return;
  if (blk.before && blk.before.length) {
    if (blk.bed && blk.bed.startsWith('AMB')) {
      sh.getRange(blk.row, AMB_COLS.DOCTOR, 1, 4).setValues([blk.before]);
    } else {
      const width = BOARD_COLUMNS.COMMENT - BOARD_COLUMNS.DOCTOR + 1;
      sh.getRange(blk.row, BOARD_COLUMNS.DOCTOR, 1, width).setValues([blk.before]);
    }
  } else {
    clearBed_(sh, blk.bed);
  }
}

// -------- ROW/SNAPSHOT HELPERS --------
function normalizeLabel_(s) {
  return String(s || '')
    .replace(/\u2019/g, "'")
    .replace(/\s+/g, ' ')
    .replace(/\u00A0/g, ' ')
    .trim()
    .toUpperCase();
}
function _rowByBedFast_(sh, bedLabel) {
  if (!bedLabel) return 0;
  const want = normalizeLabel_(bedLabel);
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return 0;

  const vals = sh.getRange(1, BOARD_COLUMNS.BED, lastRow, 1).getValues();
  for (let r = 1; r <= lastRow; r++) {
    const got = normalizeLabel_(vals[r-1][0]);
    if (got && got === want) return r;
  }
  try {
    const tf = sh.createTextFinder(bedLabel)
      .matchCase(false).matchEntireCell(true).useRegularExpression(false)
      .matchFormulaText(false).ignoreDiacritics(true);
    const hit = tf.findNext();
    if (hit && hit.getColumn() === BOARD_COLUMNS.BED) return hit.getRow();
  } catch(e) {}
  try {
    const esc = want.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    const tf2 = sh.createTextFinder(`^${esc}$`)
      .useRegularExpression(true).matchCase(false)
      .matchFormulaText(false).ignoreDiacritics(true);
    const hit2 = tf2.findNext();
    if (hit2 && hit2.getColumn() === BOARD_COLUMNS.BED) return hit2.getRow();
  } catch(e) {}
  return 0;
}
function guessRowFromLabel_(bedLabel) {
  if (bedLabel.startsWith('AMB')) {
    const idx = Number(bedLabel.replace('AMB',''));
    return 1 + idx;
  }
  return _rowByBedFast_(SpreadsheetApp.getActive().getSheetByName(SH_LENTA), bedLabel);
}
function emptyBlockForBed_(bedLabel) {
  if (bedLabel.startsWith('AMB')) {
    return ['', '', '', '']; // K..N
  } else {
    const width = BOARD_COLUMNS.COMMENT - BOARD_COLUMNS.DOCTOR + 1;
    return new Array(width).fill('');
  }
}
function readBedSnapshot_(sh, bedLabel) {
  if (bedLabel.startsWith('AMB')) {
    const row = 1 + Number(bedLabel.replace('AMB',''));
    const before = sh.getRange(row, AMB_COLS.DOCTOR, 1, 4).getValues()[0];
    const patient = String(sh.getRange(row, AMB_COLS.PATIENT).getValue() || '').trim();
    const occupied = before.some(v => String(v||'').trim()) || !!patient;
    return { bed:bedLabel, row, occupied: !!occupied, before, patient };
  } else {
    const row = _rowByBedFast_(sh, bedLabel);
    if (!row) return null;
    const width = BOARD_COLUMNS.COMMENT - BOARD_COLUMNS.DOCTOR + 1;
    const before = sh.getRange(row, BOARD_COLUMNS.DOCTOR, 1, width).getValues()[0];
    const patient = String(sh.getRange(row, PATIENT_COL).getValue() || '').trim();
    const occupied = patient || before.some(v => String(v||'').trim());
    return { bed:bedLabel, row, occupied: !!occupied, before, patient };
  }
}
function writeSnapshotToBed_(sh, bedLabel, snap) {
  if (bedLabel.startsWith('AMB')) {
    const row = 1 + Number(bedLabel.replace('AMB',''));
    if (snap && snap.before) {
      sh.getRange(row, AMB_COLS.DOCTOR, 1, 4).setValues([snap.before]);
    } else {
      sh.getRange(row, AMB_COLS.DOCTOR, 1, 4).clearContent();
    }
  } else {
    const row = _rowByBedFast_(sh, bedLabel);
    if (!row) return;
    const width = BOARD_COLUMNS.COMMENT - BOARD_COLUMNS.DOCTOR + 1;
    if (snap && snap.before) {
      sh.getRange(row, BOARD_COLUMNS.DOCTOR, 1, width).setValues([snap.before]);
      if (typeof snap.patient !== 'undefined') {
        sh.getRange(row, PATIENT_COL).setValue(snap.patient || '');
      }
    } else {
      sh.getRange(row, BOARD_COLUMNS.DOCTOR, 1, width).clearContent();
      sh.getRange(row, PATIENT_COL).clearContent();
    }
  }
}
function clearBed_(sh, bedLabel) {
  if (bedLabel.startsWith('AMB')) {
    const row = 1 + Number(bedLabel.replace('AMB',''));
    sh.getRange(row, AMB_COLS.DOCTOR, 1, 4).clearContent();
  } else {
    const row = _rowByBedFast_(sh, bedLabel);
    if (!row) return;
    const width = BOARD_COLUMNS.COMMENT - BOARD_COLUMNS.DOCTOR + 1;
    sh.getRange(row, BOARD_COLUMNS.DOCTOR, 1, width).clearContent();
    sh.getRange(row, PATIENT_COL).clearContent();
  }
}

// -------- LOGGING --------
function ensureLog_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SH_LOG);
  if (!sh) {
    sh = ss.insertSheet(SH_LOG);
    sh.appendRow(['TimeISO','User','Action','Summary','From','To','Bed','Patient','Doctor','Triage','Status','Comment']);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(['TimeISO','User','Action','Summary','From','To','Bed','Patient','Doctor','Triage','Status','Comment']);
  }
  return sh;
}
function _logAction_(o) {
  try {
    const sh = ensureLog_();
    const row = [
      new Date(), // store native Date (easier to format in Sheets)
      _resolveUser_(o && (o.userDisplayName || o.user)),
      o.action || '',
      o.summary || '',
      o.from || '',
      o.to || '',
      o.bed || '',
      o.patient || '',
      o.doctor || '',
      o.triage || '',
      o.status || '',
      o.comment || ''
    ];
    sh.appendRow(row);
  } catch(e) {
    console.error('log error', e);
  }
}
function getRecentActions(limit = 10) {
  const sh = ensureLog_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(limit, last - 1);
  const vals = sh.getRange(last - n + 1, 1, n, 12).getValues();
  return vals.map(r => ({
    ts: r[0], actor: r[1], action: r[2], summary: r[3]
  })).reverse();
}

// -------- SIDEBAR DATA --------
function sidebarGetAll(payload) {
  try {
    const actor = (payload && payload.actor)
      ? String(payload.actor).trim()
      : getUsername() || Session.getActiveUser().getEmail();
    const zonesPayload = buildZonesPayload_(actor);
    const doctors = getDoctors();
    const recent = getRecentActions(10);
    return { zonesPayload, doctors, recent, now: new Date() };
  } catch (e) {
    console.error('sidebarGetAll error:', e);
    return { zonesPayload: { beds: {}, nurses: {} }, doctors: [], recent: [], now: new Date(), error: String(e) };
  }
}

// -------- HELPERS --------
function normalizePayload_(payload) {
  if (!payload) return {};
  const p = Object.assign({}, payload);
  ['actor','bedLabel','patientName','triage','doctor','comment','fromBed','toBed','bedA','bedB','reason','userDisplayName']
    .forEach(k => { if (k in p && typeof p[k] === 'string') p[k] = p[k].trim(); });
  return p;
}
function readJson(str) {
  if (!str) return null;
  try { return JSON.parse(str); } catch (e) { return null; }
}
