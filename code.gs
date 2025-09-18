/** ===================== CONFIG ===================== **/
const SHEET_LENTA = 'LENTA';

// Pagrindinės LENTA stulpelių kolonos
const COLS = {
  bed:     'C',    // Lova (pagal šį stulpelį randame eilutę)
  doctor:  'D',    // Gydytojas
  triage:  'E',    // Triažas (1–5)
  patient: 'F',    // Pacientas (užimtumui)
  status:  'G',    // Automatiškai: "Laukia apžiūros"
  comment: 'H'     // Komentaras
};

// Ambulatorijos kolonos (K/L/M/N), eilutės 2..16
const AMB = {
  doctor:  'K',   // gydytojas
  triage:  'L',   // kategorija (triažas)
  patient: 'M',   // pacientas
  comment: 'N',   // komentaras
  firstRow: 2,
  lastRow: 16, // => AMB1..AMB15
};

// Gydytojų sąrašas iš "PACIENTŲ SKIRSTYMAS" (viena eilutė)
const DOCTORS_SHEET = 'PACIENTŲ SKIRSTYMAS';
const DOCTORS_RANGE = 'B2:I2';

// Rezervacijos TTL (sek.)
const RES_TTL_SEC = 5 * 60;
// UNDO saugojimas (sek.) – laikysime iki 10 min
const UNDO_TTL_SEC = 10 * 60;

// Veiksmų žurnalas
const LOG_SHEET = 'VEIKSMŲ ŽURNALAS';
const LOG_HEADERS = ['TimeISO','User','Action','Summary','From','To','Bed','Patient','Doctor','Triage','Status','Comment'];

/** Zonos (įskaitant Ambulatoriją) */
const AMB_BEDS = Array.from({length: (AMB.lastRow - AMB.firstRow + 1)}, (_,i)=>`AMB${i+1}`);
const ZONOS_LAYOUT = (() => {
  const ambRows = [];
  if (AMB_BEDS.length) {
    const midpoint = Math.ceil(AMB_BEDS.length / 2);
    ambRows.push(AMB_BEDS.slice(0, midpoint));
    if (midpoint < AMB_BEDS.length) {
      ambRows.push(AMB_BEDS.slice(midpoint));
    }
  }
  return [
    { name: "Zona IT", rows: [["IT1","IT2"]] },
    { name: "Zona 1", rows: [["1","2","3","P1","P2","P3"]] },
    { name: "Zona 2", rows: [["4","5","6","P4","P5","P6","S4","S5","S6"]] },
    { name: "Zona 3", rows: [["7","8","9","P7","P8","P9","S7","S8","S9"]] },
    { name: "Zona 4", rows: [["10","11","12","P10","P11","P12","S10","S11","S12"]] },
    { name: "Zona 5", rows: [["13","14","15","16","17","121A","121B","IZO"]] },
    {
      name: "Ambulatorija",
      rows: ambRows.length ? ambRows : [AMB_BEDS.slice()]
    }
  ];
})();
const ZONOS = ZONOS_LAYOUT.reduce((acc, zone) => {
  acc[zone.name] = zone.rows.flat();
  return acc;
}, {});
// Slaugytojų skaitymas iš lapo LENTA konkrečiose vietose.
// - Zona IT: A2 + A3 (abi sujungiamos " / " jei abi užpildytos)
// - Zona 1: A4–A9 (merged)  -> imame A4
// - Zona 2: A10–A17 (merged) -> A10
// - Zona 3: A18–A26 (merged) -> A18
// - Zona 4: A27–A35 (merged) -> A27
// - Zona 5: A36–A40 (merged) -> A36
// - Ambulatorija: J2 (merged per J2:J16)
function getNursesFromLenta_() {
  const sh = _sheet(SHEET_LENTA);
  const out = {};
  if (!sh) return out;

  const addresses = ["A2", "A3", "A4", "A10", "A18", "A27", "A36", "J2"];
  const values = sh
    .getRangeList(addresses)
    .getRanges()
    .map((range) => String(range.getDisplayValue() || "").trim());

  const [it1, it2, zona1, zona2, zona3, zona4, zona5, amb] = values;

  // Zona IT – sujungiam A2 ir A3, jei abu yra
  const it = [it1, it2].filter(Boolean).join(" / ");
  out["Zona IT"] = it || "Nenurodyta";

  // Merged zonos – imame viršutinį kairį langelį
  out["Zona 1"] = zona1 || "Nenurodyta";   // A4–A9
  out["Zona 2"] = zona2 || "Nenurodyta";   // A10–A17
  out["Zona 3"] = zona3 || "Nenurodyta";   // A18–A26
  out["Zona 4"] = zona4 || "Nenurodyta";   // A27–A35
  out["Zona 5"] = zona5 || "Nenurodyta";   // A36–A40

  // Ambulatorija – J2 (J2:J16 gali būti sujungtas)
  out["Ambulatorija"] = amb || "Nenurodyta";

  return out;
}


/** ===================== MENU / SIDEBAR ===================== **/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SMPS TRIAŽO LANGAS')
    .addItem('Atidaryti langą', 'rodytiLovuSkydeli')
    .addToUi();
}

function rodytiLovuSkydeli() {
  const html = HtmlService.createTemplateFromFile('triagesidebar')
    .evaluate()
    .setTitle('Lovų zonos');
  SpreadsheetApp.getUi().showSidebar(html);
}
function setUsername(name){ PropertiesService.getUserProperties().setProperty("username", name); }
function getUsername(){ return PropertiesService.getUserProperties().getProperty("username"); }
function getCurrentUserName(){ return _resolveUserIdentity_('').display; }

function logArrival(timestampSheet, row) {
  const user = getCurrentUserName();
  // ... jūsų kodas
  timestampSheet.appendRow([
    patientID, doctor, emergencyCategory, formattedTime, "", "", getWeekNumber(now),
    now.toLocaleString("default",{month:"long"}), now.getHours(), now.toLocaleString("default",{weekday:"long"}), user
  ]);
}



/** ===================== HELPERS ===================== **/
function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function _sheet(name) { return _ss().getSheetByName(name); }
function _colA1ToIndex(a1) {
  const s = String(a1).toUpperCase();
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}
function _findRowByValue(sheet, colA1, value) {
  const col = _colA1ToIndex(colA1);
  const last = sheet.getLastRow();
  if (last < 1) return -1;
  const rng = sheet.getRange(1, col, last, 1);
  const tf = rng.createTextFinder(String(value)).matchEntireCell(true);
  const cell = tf.findNext();
  return cell ? cell.getRow() : -1;
}

function _resolveUserIdentity_(userName) {
  let name = '';
  if (typeof userName === 'string') {
    name = userName.trim();
  }

  if (!name) {
    try {
      const stored = getUsername();
      if (stored) name = String(stored).trim();
    } catch (e) {}
  }

  if (!name) {
    try {
      const email = Session.getActiveUser().getEmail();
      if (email) name = String(email).trim();
    } catch (e) {}
  }

  let tempKey = '';
  try {
    tempKey = Session.getTemporaryActiveUserKey();
  } catch (e) {}

  const baseTag = name || 'anon';
  const tag = tempKey ? `${baseTag}#${tempKey}` : baseTag;
  const display = name || 'Nenurodytas';

  return { tag, display };
}

function _getUserTag(userName) {
  return _resolveUserIdentity_(userName).tag;
}

function _resKey(label) { return 'RES_BED_' + String(label); }

// AMB pagalba
function _isAmb(label){ return /^AMB(\d+)$/.test(String(label)); }
function _ambRow(label){
  const m = String(label).match(/^AMB(\d+)$/);
  if (!m) return -1;
  const n = Number(m[1]);
  return AMB.firstRow + (n - 1); // AMB1 -> 2, ..., AMB15 -> 16
}

// UNDO saugykla
function _storeUndo_(payloadObj) {
  const token = Utilities.getUuid();
  CacheService.getDocumentCache().put('UNDO_'+token, JSON.stringify(payloadObj), UNDO_TTL_SEC);
  return token;
}
function _readUndo_(token) {
  const raw = CacheService.getDocumentCache().get('UNDO_'+token);
  return raw ? JSON.parse(raw) : null;
}
function _clearUndo_(token) {
  CacheService.getDocumentCache().remove('UNDO_'+token);
}

/** Vienos eilutės rašymas į kelis stulpelius. */
function _writeRowPairs(sh, row, pairs) {
  const items = pairs
    .map(p => ({ idx: _colA1ToIndex(p.colA1), value: p.value }))
    .sort((a, b) => a.idx - b.idx);
  if (!items.length) return;

  let contiguous = true;
  for (let i = 1; i < items.length; i++) {
    if (items[i].idx !== items[0].idx + i) { contiguous = false; break; }
  }

  if (contiguous) {
    const startCol = items[0].idx;
    const valuesRow = items.map(it => it.value);
    sh.getRange(row, startCol, 1, items.length).setValues([valuesRow]);
  } else {
    items.forEach(it => sh.getRange(row, it.idx).setValue(it.value));
  }
}

/** ===================== AUDIT LOG ===================== **/
function _logSheet_() {
  let sh = _sheet(LOG_SHEET);
  if (!sh) {
    sh = _ss().insertSheet(LOG_SHEET);
    sh.getRange(1,1,1,LOG_HEADERS.length).setValues([LOG_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function _logAction_(o) {
  try {
    const sh = _logSheet_();
    const fallbackName = (o && o.userName) ? o.userName : (o && o.userTag ? String(o.userTag).split('#')[0] : '');
    const identity = (o && o.userIdentity) ? o.userIdentity : _resolveUserIdentity_(fallbackName);
    const userCell = identity && identity.display ? identity.display : 'Nenurodytas';
    const row = [
      new Date().toISOString(),
      userCell,
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
    const last = sh.getLastRow();
    if (last > 5000) sh.deleteRows(2, last - 5000);
  } catch(e) {}
}

/** ===================== DATA FOR SIDEBAR ===================== **/
function sidebarGetAll(payload) {
  const userName = (payload && typeof payload === 'object') ? payload.userName : '';
  const layout = ZONOS_LAYOUT.map(zone => ({
    name: zone.name,
    rows: zone.rows.map(row => row.slice())
  }));
  return {
    zonesPayload: getLiveZoneData(userName),
    layout,
    doctors: getDoctorsList_(),
    recent: _getRecentActions_(8),
    now: new Date().toISOString()
  };
}
function _getRecentActions_(limit) {
  const sh = _sheet(LOG_SHEET);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(limit || 8, last - 1);
  const start = last - n + 1;
  const values = sh.getRange(start, 1, n, LOG_HEADERS.length).getValues();
  return values.map(r => ({
    ts: r[0], user: r[1], action: r[2], summary: r[3],
    from: r[4], to: r[5], bed: r[6], patient: r[7], doctor: r[8], triage: r[9], status: r[10], comment: r[11]
  })).reverse();
}

/** Užimtumas (LENTA F; AMB M) ir slaugytojos, pac. vardas */
function getLiveZoneData(userName) {
  const sh = _sheet(SHEET_LENTA);
  const me = _getUserTag(userName);
  const cache = CacheService.getDocumentCache();

  // Visi rezervacijų raktai (vienkartinis užklausimas į Cache)
  const allBedLabels = Array.from(new Set(Object.values(ZONOS).flat()));
  const reservationKeys = allBedLabels.map(label => _resKey(label));
  const reservationOwners = (cache && typeof cache.getAll === 'function')
    ? (cache.getAll(reservationKeys) || {})
    : {};
  const reservationOwnersByLabel = allBedLabels.reduce((acc, label) => {
    acc[label] = reservationOwners[_resKey(label)];
    return acc;
  }, {});

  // Surenkam pacientus pagal lovą (Salė + AMB)
  const patientByBed = {};

  if (sh) {
    // Salė: C=bed, F=patient
    const bedCol = _colA1ToIndex(COLS.bed);
    const patientCol = _colA1ToIndex(COLS.patient);
    const last = sh.getLastRow();
    if (last >= 2) {
      const beds = sh.getRange(2, bedCol, last - 1, 1).getValues().flat();
      const patients = sh.getRange(2, patientCol, last - 1, 1).getValues().flat();
      for (let i = 0; i < beds.length; i++) {
        const b = String(beds[i] || '').trim();
        const p = String(patients[i] || '').trim();
        if (b) patientByBed[b] = p; // gali būti ir tuščias
      }
    }
    // AMB: AMB1..AMB15 → M=patient
    const ambPIdx = _colA1ToIndex(AMB.patient);
    const ambRows = AMB.lastRow - AMB.firstRow + 1;
    if (ambRows > 0) {
      const ambPats = sh.getRange(AMB.firstRow, ambPIdx, ambRows, 1).getValues().flat();
      ambPats.forEach((val, i) => {
        const label = `AMB${i+1}`;
        patientByBed[label] = String(val || '').trim();
      });
    }
  }

  // Sukuriam payloadą su pacientu
  const bedsPayload = {};
  Object.keys(ZONOS).forEach(z => {
    bedsPayload[z] = {
      beds: ZONOS[z].map(label => {
        const holder = reservationOwnersByLabel[label];
        const patient = patientByBed[label] || '';
        return {
          label,
          occupied: !!patient,
          patient, // <— pridėta
          reservedByMe: !!(holder && holder === me),
          reservedByOther: !!(holder && holder !== me)
        };
      })
    };
  });

  // Slaugytojų vardai iš LENTA
  const nurseMap = getNursesFromLenta_();
  const nurses = {};

  // Užtikriname, kad visoms zonoms būtų reikšmė (jei kas tuščia – "Nenurodyta")
  Object.keys(ZONOS).forEach(z => {
    nurses[z] = nurseMap[z] || "Nenurodyta";
  });
  // Ambulatorija (jei UI ją rodo atskirai)
  nurses["Ambulatorija"] = nurseMap["Ambulatorija"] || "Nenurodyta";

  return { beds: bedsPayload, nurses };
}


/** Gydytojų sąrašas */
function getDoctorsList_() {
  const sh = _sheet(DOCTORS_SHEET);
  if (!sh) return [];
  const rng = sh.getRange(DOCTORS_RANGE).getValues(); // 1 x N
  const vals = (rng[0] || []).map(v => String(v || '').trim()).filter(Boolean);
  const seen = {};
  const out = [];
  vals.forEach(v => { if (!seen[v]) { seen[v] = true; out.push(v); } });
  return out;
}

/** ===================== RESERVATIONS (su LockService) ===================== **/
function reserveBed(payload) {
  const data = (typeof payload === 'object' && payload !== null) ? payload : { bedLabel: payload };
  const bedLabel = data && data.bedLabel;
  if (!bedLabel) return { ok:false, msg:'Nenurodytas lovos žymuo.' };

  const userIdentity = _resolveUserIdentity_(data.userName);
  const me = userIdentity.tag;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { ok:false, msg:'Sistema užimta. Bandykite dar kartą.' };
  try {
    const sh = _sheet(SHEET_LENTA);
    if (sh) {
      let occupied = false;
      if (_isAmb(bedLabel)) {
        const row = _ambRow(bedLabel);
        const p = String(sh.getRange(row, _colA1ToIndex(AMB.patient)).getValue() || '').trim();
        occupied = !!p;
      } else {
        const row = _findRowByValue(sh, COLS.bed, bedLabel);
        if (row > 0) {
          const p = String(sh.getRange(row, _colA1ToIndex(COLS.patient)).getValue() || '').trim();
          occupied = !!p;
        }
      }
      if (occupied) return { ok:false, msg:'Lova jau užimta.' };
    }
    const cache = CacheService.getDocumentCache();
    const key = _resKey(bedLabel);
    const holder = cache.get(key);

    if (holder && holder !== me) return { ok:false, msg:'Lova jau rezervuota kito naudotojo.' };

    cache.put(key, me, RES_TTL_SEC);
    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}
function keepAliveReservation(payload) {
  const data = (typeof payload === 'object' && payload !== null) ? payload : { bedLabel: payload };
  const bedLabel = data && data.bedLabel;
  if (!bedLabel) return;
  const cache = CacheService.getDocumentCache();
  const key = _resKey(bedLabel);
  const holder = cache.get(key);
  const me = _getUserTag(data.userName);
  if (holder && holder === me) cache.put(key, me, RES_TTL_SEC);
}
function releaseReservation(payload) {
  const data = (typeof payload === 'object' && payload !== null) ? payload : { bedLabel: payload };
  const bedLabel = data && data.bedLabel;
  if (!bedLabel) return;
  CacheService.getDocumentCache().remove(_resKey(bedLabel));
}

/** ===================== WRITE / UNDO / MOVE / SWAP / DISCHARGE ===================== **/

function _normalizeTriage(t) {
  if (t === null || t === undefined) return '';
  const s = String(t).trim().toUpperCase();
  if (!s) return '';
  const map = { 'I':'1','II':'2','III':'3','IV':'4','V':'5' };
  if (map[s]) return Number(map[s]);
  const n = Number(s);
  return (n>=1 && n<=5) ? n : '';
}

/**
 * payload = { bedLabel, patientName, triage, doctor, comment }
 * LENTA: rašo D..H (status="Laukia apžiūros")
 * AMB:   rašo K..N (triažas neprivalomas)
 */
function assignBed(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { ok:false, msg:'Sistema užimta. Bandykite dar kartą.' };

  try {
    const { bedLabel, patientName, triage, doctor, comment, userName } = payload || {};
    const userIdentity = _resolveUserIdentity_(userName);
    const me = userIdentity.tag;
    if (!bedLabel || !patientName || !doctor) {
      return { ok:false, msg:'Trūksta laukų (lova, vardas, gydytojas).' };
    }
    const isAmb = _isAmb(bedLabel);

    const cache = CacheService.getDocumentCache();
    const key = _resKey(bedLabel);
    const holder = cache.get(key);
    if (holder && holder !== me) return { ok:false, msg:'Lova rezervuota kito naudotojo.' };

    const sh = _sheet(SHEET_LENTA);
    if (!sh) return { ok:false, msg:'Nerastas lapas „LENTA“.' };

    if (isAmb) {
      const row = _ambRow(bedLabel);
      if (row <= 0) return { ok:false, msg:'Nerasta AMB eilutė.' };
      const existing = String(sh.getRange(row, _colA1ToIndex(AMB.patient)).getValue() || '').trim();
      if (existing) return { ok:false, msg:'Lova jau užimta.' };

      const triageOut = _normalizeTriage(triage);

      // UNDO senos K..N
      const kIdx = _colA1ToIndex(AMB.doctor);
      const nIdx = _colA1ToIndex(AMB.comment);
      const width = nIdx - kIdx + 1;
      const oldVals = sh.getRange(row, kIdx, 1, width).getValues()[0];

      // Rašom K/L/M/N
      _writeRowPairs(sh, row, [
        { colA1: AMB.doctor,  value: doctor },
        { colA1: AMB.triage,  value: triageOut },
        { colA1: AMB.patient, value: patientName },
        { colA1: AMB.comment, value: comment || '' }
      ]);

      const undoToken = _storeUndo_({
        type: 'assign',
        sheet: SHEET_LENTA,
        row,
        startIdx: kIdx,
        values: oldVals
      });

      _logAction_({
        action: 'ASSIGN_AMB',
        userIdentity,
        summary: `AMB ${bedLabel}: ${patientName}${triageOut?` (T${triageOut})`:''}, gyd. ${doctor}`,
        bed: bedLabel, patient: patientName, doctor: doctor, triage: String(triageOut||''), comment: comment || ''
      });

      cache.remove(key);
      return { ok:true, undoToken };
    } else {
      // LENTA
      const row = _findRowByValue(sh, COLS.bed, bedLabel);
      if (row <= 0) return { ok:false, msg:'Nerasta lovos eilutė LENTA lape.' };
      const existing = String(sh.getRange(row, _colA1ToIndex(COLS.patient)).getValue() || '').trim();
      if (existing) return { ok:false, msg:'Lova jau užimta.' };

      const triageOut = _normalizeTriage(triage);
      if (triageOut === '') return { ok:false, msg:'Pasirinkite triažo kategoriją (1–5).' };

      // UNDO senos D..H
      const dIdx = _colA1ToIndex(COLS.doctor);
      const hIdx = _colA1ToIndex(COLS.comment);
      const oldVals = sh.getRange(row, dIdx, 1, hIdx - dIdx + 1).getValues()[0];

      // Rašom D..H (status visada "Laukia apžiūros")
      _writeRowPairs(sh, row, [
        { colA1: COLS.doctor,  value: doctor },
        { colA1: COLS.triage,  value: triageOut },
        { colA1: COLS.patient, value: patientName },
        { colA1: COLS.status,  value: "Laukia apžiūros" },
        { colA1: COLS.comment, value: comment || '' }
      ]);

      const undoToken = _storeUndo_({
        type: 'assign',
        sheet: SHEET_LENTA,
        row,
        startIdx: dIdx,
        values: oldVals
      });

      _logAction_({
        action: 'ASSIGN',
        userIdentity,
        summary: `Priskirta ${bedLabel}: ${patientName} (T${triageOut}), gyd. ${doctor}`,
        bed: bedLabel, patient: patientName, doctor: doctor, triage: String(triageOut), status: "Laukia apžiūros", comment: comment || ''
      });

      cache.remove(key);
      return { ok:true, undoToken };
    }
  } catch (err) {
    return { ok:false, msg: 'Klaida: ' + (err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

/** Undo (assign ar move) */
function undoAssign(payload) {
  const dataIn = (typeof payload === 'object' && payload !== null) ? payload : { token: payload };
  const undoToken = dataIn && dataIn.token;
  if (!undoToken) return { ok:false, msg:'Nėra undo žetono.' };
  const userIdentity = _resolveUserIdentity_(dataIn.userName);

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { ok:false, msg:'Sistema užimta. Bandykite dar kartą.' };

  try {
    const data = _readUndo_(undoToken);
    if (!data) return { ok:false, msg:'Nebegalima atšaukti (pasibaigė laikas).' };

    const sh = _sheet(data.sheet);
    if (!sh) return { ok:false, msg:'Nerastas lapas.' };

    if (data.type === 'assign') {
      if (data.startIdx && data.values) {
        sh.getRange(data.row, data.startIdx, 1, data.values.length).setValues([data.values]);
      } else if (data.cols && data.oldValues) {
        const dIdx = _colA1ToIndex(COLS.doctor);
        sh.getRange(data.row, dIdx, 1, data.oldValues.length).setValues([data.oldValues]);
      }
      _logAction_({ action: 'UNDO_ASSIGN', summary: `Atšaukta priskyrimas`, userIdentity });
      _clearUndo_(undoToken);
      return { ok:true };
    }

    if (data.type === 'move') {
      data.rows.forEach(r => {
        sh.getRange(r.row, r.startIdx, 1, r.values.length).setValues([r.values]);
      });
      _logAction_({ action: 'UNDO_MOVE', summary: `Atšauktas perkėlimas`, userIdentity });
      _clearUndo_(undoToken);
      return { ok:true };
    }

    return { ok:false, msg:'Nežinomas undo tipas.' };
  } catch (err) {
    return { ok:false, msg:'Klaida: ' + (err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Perkėlimas: leidžiama LENTA↔LENTA (D..H), AMB↔AMB (K..N) ir AMB→LENTA (triažas persikelia).
 * LENTA→AMB – neleidžiama.
 */
function movePatient(payload) {
  const { fromBed, toBed, userName } = payload || {};
  if (!fromBed || !toBed) return { ok:false, msg:'Nurodykite iš kur ir į kur perkelti.' };

  const fromIsAmb = /^AMB/.test(fromBed);
  const toIsAmb   = /^AMB/.test(toBed);

  const userIdentity = _resolveUserIdentity_(userName);

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { ok:false, msg:'Sistema užimta. Bandykite dar kartą.' };

  try {
    const sh = _sheet(SHEET_LENTA);
    if (!sh) return { ok:false, msg:'Nerastas lapas „LENTA“.' };

    // --- 1) AMB -> AMB ---
    if (fromIsAmb && toIsAmb) {
      const fromRow = _ambRow(fromBed);
      const toRow   = _ambRow(toBed);
      if (fromRow <= 0 || toRow <= 0) return { ok:false, msg:'Nerastos AMB eilutės.' };

      const pFrom = String(sh.getRange(fromRow, _colA1ToIndex(AMB.patient)).getValue() || '').trim();
      if (!pFrom) return { ok:false, msg:'Pradinė lova tuščia.' };

      const pTo = String(sh.getRange(toRow, _colA1ToIndex(AMB.patient)).getValue() || '').trim();
      if (pTo) return { ok:false, msg:'Tikslo lova užimta.' };

      const kIdx = _colA1ToIndex(AMB.doctor);
      const nIdx = _colA1ToIndex(AMB.comment);
      const width = nIdx - kIdx + 1;

      const fromVals  = sh.getRange(fromRow, kIdx, 1, width).getValues()[0];
      const toOldVals = sh.getRange(toRow,   kIdx, 1, width).getValues()[0];

      sh.getRange(toRow, kIdx, 1, width).setValues([fromVals]);
      sh.getRange(fromRow, kIdx, 1, width).clearContent();

      const undoToken = _storeUndo_({
        type: 'move',
        sheet: SHEET_LENTA,
        rows: [
          { row: fromRow, startIdx: kIdx, values: fromVals },
          { row: toRow,   startIdx: kIdx, values: toOldVals }
        ]
      });

      const doctor  = String(fromVals[0] || '').trim();
      const triage  = String(fromVals[1] || '').trim();
      const patient = String(fromVals[2] || '').trim();

      _logAction_({
        action: 'MOVE_AMB',
        userIdentity,
        summary: `AMB ${fromBed} → ${toBed}: ${patient}${triage?` (T${triage})`:''}, gyd. ${doctor}`,
        from: fromBed, to: toBed, bed: toBed, patient, doctor, triage
      });

      return { ok:true, undoToken };
    }

    // --- 2) LENTA -> LENTA ---
    if (!fromIsAmb && !toIsAmb) {
      const fromRow = _findRowByValue(sh, COLS.bed, fromBed);
      const toRow   = _findRowByValue(sh, COLS.bed, toBed);
      if (fromRow <= 0 || toRow <= 0) return { ok:false, msg:'Nerastos lovų eilutės.' };

      const pFrom = String(sh.getRange(fromRow, _colA1ToIndex(COLS.patient)).getValue() || '').trim();
      if (!pFrom) return { ok:false, msg:'Pradinė lova tuščia.' };

      const pTo = String(sh.getRange(toRow, _colA1ToIndex(COLS.patient)).getValue() || '').trim();
      if (pTo) return { ok:false, msg:'Tikslo lova užimta.' };

      const dIdx = _colA1ToIndex(COLS.doctor);
      const hIdx = _colA1ToIndex(COLS.comment);
      const width = hIdx - dIdx + 1;

      const fromVals  = sh.getRange(fromRow, dIdx, 1, width).getValues()[0];
      const toOldVals = sh.getRange(toRow,   dIdx, 1, width).getValues()[0];

      sh.getRange(toRow, dIdx, 1, width).setValues([fromVals]);
      sh.getRange(fromRow, dIdx, 1, width).clearContent();

      const undoToken = _storeUndo_({
        type: 'move',
        sheet: SHEET_LENTA,
        rows: [
          { row: fromRow, startIdx: dIdx, values: fromVals },
          { row: toRow,   startIdx: dIdx, values: toOldVals }
        ]
      });

      const doctor  = String(fromVals[0] || '').trim();
      const triage  = String(fromVals[1] || '').trim();
      const patient = String(fromVals[2] || '').trim();

      _logAction_({
        action: 'MOVE',
        userIdentity,
        summary: `Perkelta ${fromBed} → ${toBed}: ${patient} (T${triage}), gyd. ${doctor}`,
        from: fromBed, to: toBed, bed: toBed, patient, doctor, triage
      });

      return { ok:true, undoToken };
    }

    // --- 3) AMB -> LENTA: triage persikelia į E, status = "Laukia apžiūros"
    if (fromIsAmb && !toIsAmb) {
      const fromRow = _ambRow(fromBed);
      const toRow   = _findRowByValue(sh, COLS.bed, toBed);
      if (fromRow <= 0 || toRow <= 0) return { ok:false, msg:'Nerastos lovų eilutės.' };

      const pFrom = String(sh.getRange(fromRow, _colA1ToIndex(AMB.patient)).getValue() || '').trim();
      if (!pFrom) return { ok:false, msg:'Pradinė lova tuščia.' };

      const pTo = String(sh.getRange(toRow, _colA1ToIndex(COLS.patient)).getValue() || '').trim();
      if (pTo) return { ok:false, msg:'Tikslo lova užimta.' };

      const kIdx = _colA1ToIndex(AMB.doctor);
      const nIdx = _colA1ToIndex(AMB.comment);
      const widthAmb = nIdx - kIdx + 1;
      const ambVals = sh.getRange(fromRow, kIdx, 1, widthAmb).getValues()[0]; // [doctor, triage, patient, comment]

      const doctor  = String(ambVals[0] || '').trim();
      const triageA = _normalizeTriage(ambVals[1]);
      const patient = String(ambVals[2] || '').trim();
      const comment = String(ambVals[3] || '').trim();

      const dIdx = _colA1ToIndex(COLS.doctor);
      const eIdx = _colA1ToIndex(COLS.triage);
      const fIdx = _colA1ToIndex(COLS.patient);
      const gIdx = _colA1ToIndex(COLS.status);
      const hIdx = _colA1ToIndex(COLS.comment);

      const toOldVals = sh.getRange(toRow, dIdx, 1, hIdx - dIdx + 1).getValues()[0];

      sh.getRange(toRow, dIdx).setValue(doctor);
      if (triageA === '') sh.getRange(toRow, eIdx).clearContent(); else sh.getRange(toRow, eIdx).setValue(triageA);
      sh.getRange(toRow, fIdx).setValue(patient);
      sh.getRange(toRow, gIdx).setValue("Laukia apžiūros");
      sh.getRange(toRow, hIdx).setValue(comment);

      sh.getRange(fromRow, kIdx, 1, widthAmb).clearContent();

      const undoToken = _storeUndo_({
        type: 'move',
        sheet: SHEET_LENTA,
        rows: [
          { row: fromRow, startIdx: kIdx, values: ambVals },
          { row: toRow,   startIdx: dIdx, values: toOldVals }
        ]
      });

      _logAction_({
        action: 'MOVE_AMB_TO_LENTA',
        userIdentity,
        summary: `AMB ${fromBed} → ${toBed}: ${patient}${triageA?` (T${triageA})`:''}, gyd. ${doctor}; Status=Laukia apžiūros`,
        from: fromBed, to: toBed, bed: toBed, patient, doctor, triage: String(triageA||''), status: "Laukia apžiūros", comment
      });

      return { ok:true, undoToken };
    }

    // --- 4) LENTA -> AMB – neleidžiama ---
    return { ok:false, msg:'Negalima perkelti iš Salės į Ambulatoriją.' };

  } catch (err) {
    return { ok:false, msg:'Klaida: ' + (err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

/** SWAP: sukeitimas dviejų UŽIMTŲ lovų (Salė↔Salė arba AMB↔AMB) */
function swapBeds(payload){
  const { bedA, bedB, userName } = payload || {};
  if(!bedA || !bedB || bedA===bedB) return { ok:false, msg:"Neteisingi parametrai." };

  const aAmb = _isAmb(bedA), bAmb = _isAmb(bedB);
  if (aAmb !== bAmb) return { ok:false, msg:"Sukeitimas tarp zonų (Salė↔AMB) neleidžiamas." };

  const userIdentity = _resolveUserIdentity_(userName);

  const lock = LockService.getDocumentLock();
  if(!lock.tryLock(5000)) return { ok:false, msg:"Sistema užimta." };
  try{
    const sh = _sheet(SHEET_LENTA);

    if (aAmb) {
      const rowA=_ambRow(bedA), rowB=_ambRow(bedB);
      const k=_colA1ToIndex(AMB.doctor), n=_colA1ToIndex(AMB.comment), w=n-k+1;
      const A=sh.getRange(rowA,k,1,w).getValues()[0], B=sh.getRange(rowB,k,1,w).getValues()[0];
      if(!String(A[2]||'').trim() || !String(B[2]||'').trim()) return { ok:false, msg:"Abi lovos turi būti užimtos." };
      sh.getRange(rowA,k,1,w).setValues([B]);
      sh.getRange(rowB,k,1,w).setValues([A]);

      const undoToken=_storeUndo_({ type:'move', sheet:SHEET_LENTA, rows:[
        { row:rowA, startIdx:k, values:A }, { row:rowB, startIdx:k, values:B }
      ]});
      _logAction_({ action:"SWAP_AMB", summary:`AMB ${bedA} ⇄ ${bedB}`, userIdentity });
      return { ok:true, undoToken };
    } else {
      const rowA=_findRowByValue(sh, COLS.bed, bedA), rowB=_findRowByValue(sh, COLS.bed, bedB);
      const d=_colA1ToIndex(COLS.doctor), h=_colA1ToIndex(COLS.comment), w=h-d+1;
      const A=sh.getRange(rowA,d,1,w).getValues()[0], B=sh.getRange(rowB,d,1,w).getValues()[0];
      if(!String(A[2]||'').trim() || !String(B[2]||'').trim()) return { ok:false, msg:"Abi lovos turi būti užimtos." };
      sh.getRange(rowA,d,1,w).setValues([B]);
      sh.getRange(rowB,d,1,w).setValues([A]);

      const undoToken=_storeUndo_({ type:'move', sheet:SHEET_LENTA, rows:[
        { row:rowA, startIdx:d, values:A }, { row:rowB, startIdx:d, values:B }
      ]});
      _logAction_({ action:"SWAP", summary:`${bedA} ⇄ ${bedB}`, userIdentity });
      return { ok:true, undoToken };
    }
  } catch (err) {
    return { ok:false, msg:'Klaida: ' + (err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

/** Išrašyti / išvalyti paciento įrašą (su Undo) */
function dischargePatient(payload) {
  const { bedLabel, reason, userName } = payload || {};
  if (!bedLabel) return { ok:false, msg:"Nenurodyta lova." };

  const userIdentity = _resolveUserIdentity_(userName);

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { ok:false, msg:"Sistema užimta." };

  try {
    const sh = _sheet(SHEET_LENTA);
    if (!sh) return { ok:false, msg:'Nerastas lapas „LENTA“.' };

    if (_isAmb(bedLabel)) {
      // --- AMB ---
      const row = _ambRow(bedLabel);
      if (row <= 0) return { ok:false, msg:`Nerasta AMB eilutė (${bedLabel}).` };

      const kIdx = _colA1ToIndex(AMB.doctor);
      const lIdx = _colA1ToIndex(AMB.triage);
      const mIdx = _colA1ToIndex(AMB.patient);
      const nIdx = _colA1ToIndex(AMB.comment);

      const existingPatient = String(sh.getRange(row, mIdx).getValue() || '').trim();
      if (!existingPatient) return { ok:false, msg:"Lova tuščia (AMB)." };

      // Išsisaugom seną reikšmių bloką K..N
      const old = sh.getRange(row, kIdx, 1, nIdx - kIdx + 1).getValues()[0];

      // (neprivaloma) archyvuokite prieš trynimą:
      // _archive_({ bed: bedLabel, patient: old[2], doctor: old[0], triage: old[1], status: '', comment: old[3] });

      sh.getRange(row, kIdx, 1, nIdx - kIdx + 1).clearContent();

      const undoToken = _storeUndo_({ type:'assign', sheet:SHEET_LENTA, row, startIdx:kIdx, values: old });
      _logAction_({ action:"DISCHARGE_AMB", summary:`AMB ${bedLabel}: ${reason||'išrašyta'}`, bed:bedLabel, comment:reason||'', userIdentity });
      return { ok:true, undoToken, msg:"Išrašyta iš Ambulatorijos." };
    } else {
      // --- LENTA ---
      // Greitas paieškos kelias (jei turite indeksą), su „fallback“
      // --- LENTA ---
      let row = _findRowByValue(sh, COLS.bed, bedLabel);   // <— tik lėtas, bet saugus paieškos būdas
      if (row <= 0) return { ok:false, msg:`Nerasta lova (${bedLabel}).` };


      const dIdx = _colA1ToIndex(COLS.doctor);
      const eIdx = _colA1ToIndex(COLS.triage);
      const fIdx = _colA1ToIndex(COLS.patient);
      const gIdx = _colA1ToIndex(COLS.status);
      const hIdx = _colA1ToIndex(COLS.comment);

      const existingPatient = String(sh.getRange(row, fIdx).getValue() || '').trim();
      if (!existingPatient) return { ok:false, msg:"Lova tuščia." };

      // Išsisaugom D..H bloką „undo“ ir (neprivalomai) archyvui
      const old = sh.getRange(row, dIdx, 1, hIdx - dIdx + 1).getValues()[0];
      // _archive_({ bed: bedLabel, patient: old[2], doctor: old[0], triage: old[1], status: old[3], comment: old[4] });

      // Šalinam visą turinį D..H (gyd., triažas, pacientas, statusas, komentaras)
      sh.getRange(row, dIdx, 1, hIdx - dIdx + 1).clearContent();

      const undoToken = _storeUndo_({ type:'assign', sheet:SHEET_LENTA, row, startIdx:dIdx, values: old });
      _logAction_({ action:"DISCHARGE", summary:`${bedLabel}: ${reason||'išrašyta'}`, bed:bedLabel, comment:reason||'', userIdentity });
      return { ok:true, undoToken, msg:"Išrašyta iš Salės." };
    }
  } catch (err) {
    return { ok:false, msg:'Klaida išrašant: ' + (err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

