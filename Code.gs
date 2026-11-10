/**
 * ALL INDIA POLICE DIRECTORY - Backend API (Google Apps Script)
 * Front-end: https://allindiapolicedirectory.github.io (calls this via fetch)
 *
 * Sheet ID  : 1ghy26XqLN3VHMvVmiruLS_Vjg_H3hXyr7kAR1zhpAFQ
 * Sheet Name: AIPD
 * Columns   : A SNO | B Name | C Email | D Mobile | E User Type | F State | G District | H Station
 *
 * Deploy as Web app → Execute as: Me → Who has access: Anyone
 */

// ============ CONFIGURATION ============
const SHEET_ID = '1ghy26XqLN3VHMvVmiruLS_Vjg_H3hXyr7kAR1zhpAFQ';
const SHEET_NAME = 'AIPD';

// ============ WEB APP ENTRY POINTS ============
/** GET ?action=ping|states|districts|stations|search */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) {
    return json_({ success: true, message: 'AIPD API is running' });
  }
  return handleApi_(p.action, p);
}

/** POST body (text/plain): { "action": "update", "payload": { rowIndex, name, ... } } */
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ success: false, error: 'Invalid JSON body' });
  }
  return handleApi_(body.action, body);
}

function handleApi_(action, p) {
  try {
    let res;
    switch (String(action || '')) {
      case 'ping':      res = { success: true }; break;
      case 'states':    res = getStates(); break;
      case 'districts': res = getDistricts(p.state); break;
      case 'stations':  res = getStations(p.state, p.district); break;
      case 'search':
        res = searchRecords({
          state: p.state, district: p.district, station: p.station,
          name: p.name, mobile: p.mobile
        });
        break;
      case 'update':
        res = updateRecord(typeof p.payload === 'string' ? JSON.parse(p.payload) : p.payload);
        break;
      default:
        res = { success: false, error: 'Unknown action: ' + action };
    }
    return json_(res);
  } catch (err) {
    return json_({ success: false, error: err.message });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found');
  return sheet;
}

// ============ DROPDOWNS ============
function getStates() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('AIPD_STATES');
    if (cached) return { success: true, states: JSON.parse(cached) };

    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, states: [] };

    const values = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    const unique = new Set();
    values.forEach(r => {
      const v = String(r[0] || '').trim();
      if (v) unique.add(v);
    });
    const states = Array.from(unique).sort();
    cache.put('AIPD_STATES', JSON.stringify(states), 600); // 10 minutes
    return { success: true, states: states };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getDistricts(state) {
  try {
    if (!state) return { success: true, districts: [] };
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, districts: [] };

    const data = sheet.getRange(2, 6, lastRow - 1, 2).getValues();
    const s0 = String(state).trim().toLowerCase();
    const unique = new Set();
    data.forEach(r => {
      const s = String(r[0] || '').trim().toLowerCase();
      const d = String(r[1] || '').trim();
      if (s === s0 && d) unique.add(d);
    });
    return { success: true, districts: Array.from(unique).sort() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getStations(state, district) {
  try {
    if (!state || !district) return { success: true, stations: [] };
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, stations: [] };

    const data = sheet.getRange(2, 6, lastRow - 1, 3).getValues();
    const s0 = String(state).trim().toLowerCase();
    const d0 = String(district).trim().toLowerCase();
    const unique = new Set();
    data.forEach(r => {
      const s = String(r[0] || '').trim().toLowerCase();
      const d = String(r[1] || '').trim().toLowerCase();
      const st = String(r[2] || '').trim();
      if (s === s0 && d === d0 && st) unique.add(st);
    });
    return { success: true, stations: Array.from(unique).sort() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============ SEARCH ============
function searchRecords(filters) {
  try {
    filters = filters || {};
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, results: [], incomplete: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

    const fState = String(filters.state || '').trim().toLowerCase();
    const fDistrict = String(filters.district || '').trim().toLowerCase();
    const fStation = String(filters.station || '').trim().toLowerCase();
    const fName = String(filters.name || '').trim().toLowerCase();
    const fMobile = String(filters.mobile || '').trim().toLowerCase();

    if (!(fState || fDistrict || fStation || fName || fMobile)) {
      return { success: false, error: 'Please enter at least one search criterion.' };
    }

    const results = [];
    const incomplete = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rec = {
        rowIndex: i + 2,
        sno: row[0],
        name: String(row[1] || ''),
        email: String(row[2] || ''),
        mobile: String(row[3] || ''),
        userType: String(row[4] || ''),
        state: String(row[5] || ''),
        district: String(row[6] || ''),
        station: String(row[7] || '')
      };

      if (fState && rec.state.toLowerCase().indexOf(fState) === -1) continue;
      if (fDistrict && rec.district.toLowerCase().indexOf(fDistrict) === -1) continue;
      if (fStation && rec.station.toLowerCase().indexOf(fStation) === -1) continue;

      if (fMobile) {
        const md = rec.mobile.replace(/\D/g, '');
        const qd = fMobile.replace(/\D/g, '');
        if (qd && md.indexOf(qd) === -1) continue;
      }

      if (fName && !fuzzyNameMatch(rec.name.toLowerCase(), fName)) continue;

      const missing = [];
      if (!rec.name.trim()) missing.push('Name');
      if (!rec.email.trim()) missing.push('Email');
      if (!rec.mobile.trim()) missing.push('Mobile');
      if (!rec.userType.trim()) missing.push('User Type');
      if (!rec.state.trim()) missing.push('State');
      if (!rec.district.trim()) missing.push('District');
      if (!rec.station.trim()) missing.push('Station');
      if (missing.length) {
        incomplete.push({ rowIndex: rec.rowIndex, sno: rec.sno, name: rec.name || '(no name)', missing: missing });
      }

      results.push(rec);
    }

    return { success: true, results: results, incomplete: incomplete, totalMatched: results.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** Fuzzy name match: substring, word match, or shuffled letters (NAGESWAR ↔ EGANRSWA) */
function fuzzyNameMatch(name, query) {
  if (!query) return true;
  if (!name) return false;
  if (name.indexOf(query) !== -1) return true;

  const words = name.split(/\s+/);
  for (let i = 0; i < words.length; i++) {
    if (words[i].indexOf(query) !== -1) return true;
    if (query.indexOf(words[i]) !== -1 && words[i].length >= 3) return true;
  }

  if (query.length >= 4) {
    const q = query.replace(/[^a-z]/g, '').split('').sort().join('');
    if (!q) return false;
    const n = name.replace(/[^a-z]/g, '').split('').sort().join('');
    if (n.indexOf(q) !== -1) return true;
    for (let i = 0; i < words.length; i++) {
      const w = words[i].replace(/[^a-z]/g, '').split('').sort().join('');
      if (w === q || isSubMultiset(q, w)) return true;
    }
  }
  return false;
}

function isSubMultiset(small, large) {
  if (small.length > large.length) return false;
  const count = {};
  for (let i = 0; i < large.length; i++) count[large[i]] = (count[large[i]] || 0) + 1;
  for (let i = 0; i < small.length; i++) {
    if (!count[small[i]]) return false;
    count[small[i]]--;
  }
  return true;
}

// ============ UPDATE (complete missing fields) ============
function updateRecord(payload) {
  const lock = LockService.getScriptLock();
  try {
    if (!payload || !payload.rowIndex) return { success: false, error: 'Invalid payload' };
    const row = parseInt(payload.rowIndex, 10);
    if (isNaN(row) || row < 2) return { success: false, error: 'Invalid row index' };

    const missing = [];
    if (!String(payload.name || '').trim()) missing.push('Name');
    if (!String(payload.email || '').trim()) missing.push('Email');
    if (!String(payload.mobile || '').trim()) missing.push('Mobile');
    if (!String(payload.userType || '').trim()) missing.push('User Type');
    if (!String(payload.state || '').trim()) missing.push('State');
    if (!String(payload.district || '').trim()) missing.push('District');
    if (!String(payload.station || '').trim()) missing.push('Station');
    if (missing.length) return { success: false, error: 'Still missing: ' + missing.join(', ') };

    lock.waitLock(10000);
    const sheet = getSheet_();
    if (row > sheet.getLastRow()) return { success: false, error: 'Row no longer exists' };

    // Columns B..H (Name..Station); SNO in column A is untouched
    sheet.getRange(row, 2, 1, 7).setValues([[
      payload.name, payload.email, payload.mobile, payload.userType,
      payload.state, payload.district, payload.station
    ]]);
    CacheService.getScriptCache().remove('AIPD_STATES');
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
