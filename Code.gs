/**
 * FRESHDROP backend (Google Apps Script) — API for Google Sheets FRESHDROP_DATABASE.
 * Sheets: TAGS (one row per tag), SCANS (one row per scan; never deleted or overwritten).
 * Run setup() once to create the sheets and the 24 tag rows.
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 * API: GET ?action=getTag|getHistory&tag_id=...   POST {action:"activate"|"saveScan", ...}
 */
const SPREADSHEET_ID = '';  // leave empty if this script is opened from inside FRESHDROP_DATABASE (Extensions > Apps Script)

const TAG_COLS = ['tag_id', 'status', 'sample_id', 'container_id', 'food_category', 'storage_condition',
  'measured_temperature', 'activation_timestamp', 'notes'];
const KEYS = ['active', 'r1', 'r2', 'r3', 'r4'];
const RAW = [];
KEYS.forEach(function (k) { ['r', 'g', 'b'].forEach(function (c) { RAW.push(k + '_raw_' + c); }); });
const CORR = ['active_corr_r', 'active_corr_g', 'active_corr_b', 'active_L', 'active_a', 'active_b', 'delta_e'];
const SCAN_COLS = ['scan_id', 'tag_id', 'sample_id', 'scan_type', 'timestamp', 'server_timestamp', 'imaging_condition']
  .concat(RAW, CORR, ['delta_e_metric', 'calibration_method', 'calibration_version']);
const TEXT_COLS = ['timestamp', 'server_timestamp', 'activation_timestamp'];   // stored as plain text to keep ISO strings intact
const STORAGE = ['Refrigerated', 'Room-temperature experimental condition', 'Other experimental condition'];

function db_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, cols) {
  var ss = db_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    cols.forEach(function (c, i) { if (TEXT_COLS.indexOf(c) >= 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@'); });
  }
  return sh;
}

/** Run once from the editor. Safe to run again: it only adds missing tag rows. */
function setup() {
  var tags = sheet_('TAGS', TAG_COLS); sheet_('SCANS', SCAN_COLS);
  var have = tags.getRange(2, 1, Math.max(tags.getLastRow() - 1, 1), 1).getValues().map(function (r) { return r[0]; });
  for (var i = 1; i <= 24; i++) {
    var id = 'FD-P-' + ('00' + i).slice(-3);
    if (have.indexOf(id) < 0) tags.appendRow([id, 'NOT ACTIVATED', '', '', '', '', '', '', '']);
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function validTag_(id) { return /^FD-P-0(0[1-9]|1\d|2[0-4])$/.test(String(id)); }
function rows_(sh, cols) {
  var n = sh.getLastRow() - 1; if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().map(function (r, i) {
    var o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = r[j] instanceof Date ? r[j].toISOString() : r[j]; });
    return o;
  });
}
function findTag_(id) {
  var r = rows_(sheet_('TAGS', TAG_COLS), TAG_COLS).filter(function (x) { return x.tag_id === id; })[0];
  if (!r) throw new Error('Tag not found in database. Run setup() first.');
  return r;
}
function strip_(o) { var c = {}; for (var k in o) if (k !== '_row') c[k] = o[k]; return c; }

function doGet(e) {
  try {
    var a = e.parameter.action, id = e.parameter.tag_id;
    if (!validTag_(id)) throw new Error('Invalid tag_id.');
    if (a === 'getTag') return out_({ ok: true, tag: strip_(findTag_(id)) });
    if (a === 'getHistory') {
      var scans = rows_(sheet_('SCANS', SCAN_COLS), SCAN_COLS).filter(function (r) { return r.tag_id === id; }).map(strip_);
      scans.sort(function (x, y) { return String(x.timestamp) < String(y.timestamp) ? -1 : 1; });
      return out_({ ok: true, scans: scans });
    }
    throw new Error('Unknown action.');
  } catch (err) { return out_({ ok: false, error: String(err.message) }); }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    if (!validTag_(d.tag_id)) throw new Error('Invalid tag_id.');
    if (d.action === 'activate') return out_(activate_(d));
    if (d.action === 'saveScan') return out_(saveScan_(d));
    throw new Error('Unknown action.');
  } catch (err) { return out_({ ok: false, error: String(err.message) }); }
  finally { lock.releaseLock(); }
}

function text_(v, name, required, max) {
  v = (v === undefined || v === null) ? '' : String(v).trim();
  if (required && !v) throw new Error(name + ' is required.');
  if (v.length > max) throw new Error(name + ' is too long.');
  return v;
}
function iso_(v, name) { if (isNaN(Date.parse(v))) throw new Error(name + ' is not a valid timestamp.'); return String(v); }

/** Validates a scan payload and returns a row array in SCAN_COLS order. Only known columns are accepted. */
function scanRow_(s, tagId, sampleId, type) {
  if (!s || s.scan_type !== type) throw new Error('Scan type must be ' + type + '.');
  var o = { scan_id: Utilities.getUuid(), tag_id: tagId, sample_id: sampleId, scan_type: type,
    timestamp: iso_(s.timestamp, 'timestamp'), server_timestamp: new Date().toISOString(),
    imaging_condition: text_(s.imaging_condition, 'imaging_condition', false, 200),
    delta_e_metric: text_(s.delta_e_metric, 'delta_e_metric', true, 20),
    calibration_method: text_(s.calibration_method, 'calibration_method', true, 50),
    calibration_version: text_(s.calibration_version, 'calibration_version', true, 50) };
  RAW.concat(CORR).forEach(function (c) {
    var n = Number(s[c]);
    if (s[c] === '' || s[c] === null || !isFinite(n)) throw new Error(c + ' must be a number.');
    if (RAW.indexOf(c) >= 0 && (n < 0 || n > 255)) throw new Error(c + ' must be 0-255.');
    o[c] = n;
  });
  return SCAN_COLS.map(function (c) { return o[c]; });
}

function activate_(d) {
  var tag = findTag_(d.tag_id);
  if (tag.status === 'ACTIVE') throw new Error('Tag is already activated.');
  var sample = text_(d.sample_id, 'Sample ID', true, 60), cont = text_(d.container_id, 'Container ID', true, 60);
  if (STORAGE.indexOf(d.storage_condition) < 0) throw new Error('Invalid storage condition.');
  var temp = '';
  if (d.measured_temperature !== '' && d.measured_temperature !== undefined && d.measured_temperature !== null) {
    temp = Number(d.measured_temperature); if (!isFinite(temp)) throw new Error('Temperature must be a number.');
  }
  var row = scanRow_(d.scan, d.tag_id, sample, 'BASELINE');
  sheet_('SCANS', SCAN_COLS).appendRow(row);            // write the baseline first ...
  var tags = sheet_('TAGS', TAG_COLS);                  // ... and only then flip the status to ACTIVE
  tags.getRange(tag._row, 1, 1, TAG_COLS.length).setValues([[d.tag_id, 'ACTIVE', sample, cont, 'Poultry / Chicken',
    d.storage_condition, temp, iso_(d.activation_timestamp, 'activation_timestamp'), text_(d.notes, 'Notes', false, 500)]]);
  return { ok: true };
}

function saveScan_(d) {
  var tag = findTag_(d.tag_id);
  if (tag.status !== 'ACTIVE') throw new Error('Tag is not activated.');
  sheet_('SCANS', SCAN_COLS).appendRow(scanRow_(d.scan, d.tag_id, tag.sample_id, 'CURRENT'));
  return { ok: true };
}
