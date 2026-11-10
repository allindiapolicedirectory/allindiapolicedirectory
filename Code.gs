/**
 * ALL INDIA POLICE DIRECTORY - Backend API (Google Apps Script)
 * Front-end: GitHub Pages (static index.html) calls this web app via fetch().
 *
 * Sheet ID  : 1ghy26XqLN3VHMvVmiruLS_Vjg_H3hXyr7kAR1zhpAFQ
 * Sheet Name: AIPD
 * Columns   : A SNO | B Name | C Email | D Mobile | E User Type | F State | G District | H Station
 *
 * SECURITY: Set a Script Property named ACCESS_KEY
 *   (Project Settings ▸ Script Properties ▸ Add property)
 * Every API call must send this key, otherwise it is rejected.
 */

// ============ CONFIGURATION ============
const SHEET_ID = '1ghy26XqLN3VHMvVmiruLS_Vjg_H3hXyr7kAR1zhpAFQ';
const SHEET_NAME = 'AIPD';
const REQUIRED_FIELDS = ['Name', 'Email', 'Mobile', 'User Type', 'State', 'District', 'Station'];

// ============ WEB APP ENTRY POINTS ============
/**
 * GET  ?action=ping|states|districts|stations|search&key=...
 * With no ?action the old Apps Script UI (Index.html) is still served, if present.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('All India Police Directory')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return handleApi_(p.action, p);
}

/**
 * POST body (sent as text/plain to avoid CORS preflight):
 *   { "action": "update", "key": "...", "payload": { rowIndex, name, ... } }
 */
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
    if (!checkKey_(p.key)) {
      return json_({ success: false, code: 401, error: 'Unauthorized: invalid access key' });
    }
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

function checkKey_(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('ACCESS_KEY');
  if (!expected) return false; // locked until ACCESS_KEY is configured
  return String(key || '') === expected;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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

    const stateValues = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    const unique = new Set();
    stateValues.forEach(row => {
      const v = String(row[0] || '').trim();
      if (v) unique.add(v);
    });
    const states = Array.from(unique).sort();
    cache.put('AIPD_STATES', JSON.stringify(states), 600); // 10 min
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
    const stateLower = String(state).trim().toLowerCase();
    const unique = new Set();
    data.forEach(row => {
      const s = String(row[0] || '').trim().toLowerCase();
      const d = String(row[1] || '').trim();
      if (s === stateLower && d) unique.add(d);
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
    const stateLower = String(state).trim().toLowerCase();
    const districtLower = String(district).trim().toLowerCase();
    const unique = new Set();
    data.forEach(row => {
      const s = String(row[0] || '').trim().toLowerCase();
      const d = String(row[1] || '').trim().toLowerCase();
      const st = String(row[2] || '').trim();
      if (s === stateLower && d === districtLower && st) unique.add(st);
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
      const record = {
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

      if (fState && record.state.toLowerCase().indexOf(fState) === -1) continue;
      if (fDistrict && record.district.toLowerCase().indexOf(fDistrict) === -1) continue;
      if (fStation && record.station.toLowerCase().indexOf(fStation) === -1) continue;

      if (fMobile) {
        const mobileDigits = record.mobile.replace(/\D/g, '');
        const queryDigits = fMobile.replace(/\D/g, '');
        if (queryDigits && mobileDigits.indexOf(queryDigits) === -1) continue;
      }

      if (fName && !fuzzyNameMatch(record.name.toLowerCase(), fName)) continue;

      const missing = [];
      if (!record.name.trim()) missing.push('Name');
      if (!record.email.trim()) missing.push('Email');
      if (!record.mobile.trim()) missing.push('Mobile');
      if (!record.userType.trim()) missing.push('User Type');
      if (!record.state.trim()) missing.push('State');
      if (!record.district.trim()) missing.push('District');
      if (!record.station.trim()) missing.push('Station');

      if (missing.length > 0) {
        incomplete.push({
          rowIndex: record.rowIndex,
          sno: record.sno,
          name: record.name || '(no name)',
          missing: missing
        });
      }
      results.push(record);
    }

    return { success: true, results: results, incomplete: incomplete, totalMatched: results.length };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

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
    const queryLetters = query.replace(/[^a-z]/g, '').split('').sort().join('');
    if (!queryLetters) return false;
    const nameLetters = name.replace(/[^a-z]/g, '').split('').sort().join('');
    if (nameLetters.indexOf(queryLetters) !== -1) return true;
    for (let i = 0; i < words.length; i++) {
      const wordLetters = words[i].replace(/[^a-z]/g, '').split('').sort().join('');
      if (wordLetters === queryLetters) return true;
      if (isSubMultiset(queryLetters, wordLetters)) return true;
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

// ============ UPDATE ============
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
    if (missing.length > 0) return { success: false, error: 'Still missing: ' + missing.join(', ') };

    lock.waitLock(10000);
    const sheet = getSheet_();
    if (row > sheet.getLastRow()) return { success: false, error: 'Row no longer exists' };

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
