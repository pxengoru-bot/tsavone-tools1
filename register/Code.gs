/**
 * Tsavone Consulting – Training Register
 *
 * Paste this file into Extensions > Apps Script of the Google Sheet
 * "Tsavone Training Register", together with Verify.html.
 * See INSTALL.md for step-by-step instructions.
 *
 * Uses only built-in Google services: SpreadsheetApp, DriveApp, DocumentApp,
 * GmailApp (for drafts), HtmlService and Utilities. QR codes and the logo are built in.
 */

/* ============================== CONSTANTS ============================== */

var REGISTER_SHEET = 'Register';
var SETTINGS_SHEET = 'Settings';

var COLUMNS = [
  'certificate_id', 'participant_name', 'organisation', 'email', 'phone',
  'course_code', 'course_title', 'start_date', 'end_date', 'location', 'format',
  'attendance_percent', 'pre_test_score', 'post_test_score', 'result',
  'certificate_type', 'issued_date', 'status', 'notes'
];

var FORMATS = ['Open enrolment', 'In-house', 'Blended', 'Field-based', 'Training of trainers'];
var STATUSES = ['Draft', 'Issued', 'Revoked'];
var CERT_TYPES = ['Completion', 'Competence'];

var RESULT_PASS = 'Pass';
var RESULT_FAIL = 'Did not complete';

var DEFAULT_SETTINGS = [
  ['pass_attendance_percent', 80, 'Minimum attendance (%) to pass'],
  ['pass_post_test_score', 60, 'Minimum post-test score to pass'],
  ['competence_post_test_score', 75, 'Post-test score at or above this earns a Certificate of Competence'],
  ['certificate_folder_name', 'Tsavone Certificates', 'Google Drive folder (in My Drive) where certificate PDFs are saved'],
  ['logo_drive_file_id', '', 'Drive file ID of the logo (PNG or JPG). Filled in by setup; replace to use another logo'],
  ['signatory_title', 'Managing Director', 'Title printed under the signature line (no personal name)'],
  ['sender_email_name', 'Tsavone Consulting Training', 'Sender name shown on certificate emails'],
  ['verify_url', '', 'Link of the verification page. Fills itself in the first time you open the published page (INSTALL.md step 3)'],
  ['course_sheet_id', '10V8_jXPTFwRl2y940v_AtmMDPU8RGQBI6yH5qDy6N4I', 'ID of the public "Tsavone Courses" sheet (code, title, days)'],
  ['hours_per_day', 7, 'Course hours printed on certificates = days x this number']
];

var NAVY = '#193591';
var NAVY_DARK = '#0c1a4f';
var ORANGE = '#F17B19';
var ORANGE_DARK = '#B04E00';
var LIGHT_BLUE = '#EFFCFF';
var GREY = '#5a6385';

var TIME_LIMIT_MS = 5 * 60 * 1000; // stop safely before Google's 6-minute limit

/* ============================== MENU ============================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tsavone')
    .addItem('Fill course titles from codes', 'menuFillCourseTitles')
    .addItem('Calculate results', 'menuCalculateResults')
    .addItem('Create certificate IDs', 'menuCreateCertificateIds')
    .addSeparator()
    .addItem('Generate certificates', 'menuGenerateCertificates')
    .addItem('Email certificates (create Gmail drafts)', 'menuEmailCertificates')
    .addSeparator()
    .addItem('Course summary', 'menuCourseSummary')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Testing')
      .addItem('Add test participants', 'menuAddTestParticipants')
      .addItem('Remove test participants', 'menuRemoveTestParticipants'))
    .addItem('Run setup again', 'setup')
    .addToUi();
}

/* ============================== SETUP ============================== */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(REGISTER_SHEET) || ss.getSheets()[0];
  sh.setName(REGISTER_SHEET);

  // Header row: write it if empty, otherwise add any missing columns at the end
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  if (existing.join('') === '') {
    sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  } else {
    var lower = existing.map(function (h) { return h.toLowerCase(); });
    var missing = COLUMNS.filter(function (c) { return lower.indexOf(c) === -1; });
    if (missing.length) sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }

  var width = sh.getLastColumn();
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, width)
    .setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold')
    .setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 36);

  var cols = headerMap_(sh);
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  var rule = function (list) {
    return SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  };
  sh.getRange(2, cols.format + 1, rows, 1).setDataValidation(rule(FORMATS));
  sh.getRange(2, cols.status + 1, rows, 1).setDataValidation(rule(STATUSES));
  sh.getRange(2, cols.certificate_type + 1, rows, 1).setDataValidation(rule(CERT_TYPES));
  ['start_date', 'end_date', 'issued_date'].forEach(function (c) {
    sh.getRange(2, cols[c] + 1, rows, 1).setNumberFormat('d mmm yyyy');
  });
  ['attendance_percent', 'pre_test_score', 'post_test_score'].forEach(function (c) {
    sh.getRange(2, cols[c] + 1, rows, 1).setNumberFormat('0');
  });
  sh.getRange(2, cols.certificate_id + 1, rows, 1).setNumberFormat('@');
  sh.getRange(2, cols.course_code + 1, rows, 1).setNumberFormat('@');

  var widths = { certificate_id: 170, participant_name: 190, organisation: 170, email: 200, course_title: 260, notes: 260, location: 140, format: 140 };
  Object.keys(widths).forEach(function (c) { if (cols[c] !== undefined) sh.setColumnWidth(cols[c] + 1, widths[c]); });

  // Settings tab: add any missing settings, never overwrite your values
  var st = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
  if (st.getLastRow() === 0) st.getRange(1, 1, 1, 3).setValues([['setting', 'value', 'what it does']]);
  var have = st.getLastRow() > 1 ? st.getRange(2, 1, st.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  var add = DEFAULT_SETTINGS.filter(function (s) { return have.indexOf(s[0]) === -1; });
  if (add.length) st.getRange(st.getLastRow() + 1, 1, add.length, 3).setValues(add);
  st.setFrozenRows(1);
  st.getRange(1, 1, 1, 3).setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold');
  st.setColumnWidth(1, 220); st.setColumnWidth(2, 380); st.setColumnWidth(3, 480);
  st.getRange(2, 2, Math.max(st.getLastRow() - 1, 1), 1).setNumberFormat('@');

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  var logoNote = ensureLogo_(ss);
  onOpen();
  ss.setActiveSheet(sh);
  ss.toast('Setup complete. ' + logoNote + ' Next: publish the verification page (INSTALL.md step 3).', 'Tsavone', 10);
}

/** Saves the built-in logo (bottom of this file) to Drive and fills in Settings > logo_drive_file_id, if it is empty. */
function ensureLogo_(ss) {
  var settings = readSettings_(ss);
  if (String(settings.logo_drive_file_id || '').trim()) return 'Logo already set.';
  if (typeof TSAVONE_LOGO_PNG_BASE64 === 'undefined') return 'No built-in logo found; add a logo file ID in Settings.';
  var name = 'Tsavone logo for certificates.png';
  var root = DriveApp.getRootFolder();
  var existing = root.getFilesByName(name);
  var file = existing.hasNext() ? existing.next()
    : root.createFile(Utilities.newBlob(Utilities.base64Decode(TSAVONE_LOGO_PNG_BASE64), 'image/png', name));
  setSetting_(ss, 'logo_drive_file_id', file.getId());
  return 'Logo saved to your Drive.';
}

/** Writes one value into the Settings tab (adds the row if needed). */
function setSetting_(ss, key, value) {
  var st = ss.getSheetByName(SETTINGS_SHEET);
  if (!st) return;
  var last = st.getLastRow();
  var keys = last > 1 ? st.getRange(2, 1, last - 1, 1).getValues() : [];
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === key) { st.getRange(i + 2, 2).setValue(value); return; }
  }
  st.getRange(last + 1, 1, 1, 2).setValues([[key, value]]);
}

/* ============================== HELPERS: SHEET ============================== */

function headerMap_(sh) {
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {};
  head.forEach(function (h, i) {
    var k = String(h).trim().toLowerCase();
    if (k && map[k] === undefined) map[k] = i;
  });
  return map;
}

function registerSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(REGISTER_SHEET);
  if (!sh) throw new Error('No "Register" tab found. Run Tsavone > Run setup again first.');
  var cols = headerMap_(sh);
  var missing = COLUMNS.filter(function (c) { return cols[c] === undefined; });
  if (missing.length) throw new Error('These columns are missing from the Register tab: ' + missing.join(', ') + '. Run Tsavone > Run setup again.');
  return { ss: ss, sh: sh, cols: cols };
}

/** Reads all data rows as objects. rowNumber is the sheet row (2 = first data row). */
function readRows_(reg) {
  var last = reg.sh.getLastRow();
  if (last < 2) return [];
  var values = reg.sh.getRange(2, 1, last - 1, reg.sh.getLastColumn()).getValues();
  return values.map(function (v, i) {
    var o = { rowNumber: i + 2 };
    Object.keys(reg.cols).forEach(function (k) { o[k] = v[reg.cols[k]]; });
    return o;
  }).filter(function (o) {
    return String(o.participant_name || '').trim() !== '' || String(o.certificate_id || '').trim() !== '';
  });
}

function setCell_(reg, rowNumber, col, value) {
  reg.sh.getRange(rowNumber, reg.cols[col] + 1).setValue(value);
}

/** Writes one column for many rows at once (fast). updates = {rowNumber: value} */
function writeColumn_(reg, col, updates) {
  var rowsNums = Object.keys(updates).map(Number);
  if (!rowsNums.length) return;
  var min = Math.min.apply(null, rowsNums), max = Math.max.apply(null, rowsNums);
  var range = reg.sh.getRange(min, reg.cols[col] + 1, max - min + 1, 1);
  var vals = range.getValues();
  rowsNums.forEach(function (r) { vals[r - min][0] = updates[r]; });
  range.setValues(vals);
}

function getSettings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  return readSettings_(ss);
}

function readSettings_(ss) {
  var s = {};
  DEFAULT_SETTINGS.forEach(function (d) { s[d[0]] = d[1]; });
  var st = ss.getSheetByName(SETTINGS_SHEET);
  if (st && st.getLastRow() > 1) {
    st.getRange(2, 1, st.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]).trim();
      if (k) s[k] = typeof r[1] === 'string' ? r[1].trim() : r[1];
    });
  }
  ['pass_attendance_percent', 'pass_post_test_score', 'competence_post_test_score', 'hours_per_day'].forEach(function (k) {
    s[k] = Number(s[k]);
  });
  return s;
}

/** Rows the person has selected on the Register tab (data rows only). */
function selectedRowNumbers_(reg) {
  var active = reg.ss.getActiveSheet();
  if (active.getName() !== REGISTER_SHEET) return [];
  var list = active.getActiveRangeList();
  if (!list) return [];
  var seen = {};
  list.getRanges().forEach(function (r) {
    for (var i = r.getRow(); i < r.getRow() + r.getNumRows(); i++) if (i >= 2) seen[i] = true;
  });
  return Object.keys(seen).map(Number).sort(function (a, b) { return a - b; });
}

function ui_() { return SpreadsheetApp.getUi(); }

/* ============================== HELPERS: PURE LOGIC ============================== */
/* These functions do not touch Google services, so they can be tested outside Apps Script. */

function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }

function isBlank_(v) { return v === null || v === undefined || String(v).trim() === ''; }

/** Number from a cell, or null. Accepts 85, "85", "85%", or 0.85 for attendance. */
function toNumber_(v, isPercent) {
  if (isBlank_(v)) return null;
  var n = typeof v === 'number' ? v : parseFloat(String(v).replace('%', '').trim());
  if (!isFinite(n)) return null;
  if (isPercent && n > 0 && n <= 1 && String(v).indexOf('%') === -1) n = n * 100;
  return n;
}

/** Works out result and certificate type from attendance and post-test score. */
function computeResult_(attendance, post, settings) {
  if (attendance === null || post === null) return null;
  var pass = attendance >= settings.pass_attendance_percent && post >= settings.pass_post_test_score;
  return {
    result: pass ? RESULT_PASS : RESULT_FAIL,
    certificate_type: pass ? (post >= settings.competence_post_test_score ? 'Competence' : 'Completion') : ''
  };
}

/** Replaces any earlier "Improvement: ..." note, keeping the rest of the notes. */
function improvementNote_(notes, pre, post) {
  var parts = String(notes || '').split(' | ').map(function (p) { return p.trim(); })
    .filter(function (p) { return p && p.indexOf('Improvement:') !== 0; });
  if (pre !== null && post !== null) {
    var diff = Math.round((post - pre) * 10) / 10;
    parts.push('Improvement: ' + (diff > 0 ? '+' : '') + diff + ' points (pre ' + pre + ', post ' + post + ')');
  }
  return parts.join(' | ');
}

/** Cleans a course code for use in IDs and folder names, e.g. " t1.1 " -> "T1.1". */
function cleanCode_(code) {
  return String(code || '').trim().toUpperCase().replace(/\s+/g, '');
}

/** Next unique certificate ID, e.g. TSV-2026-T1.1-0001. usedIds is an object of IDs already taken. */
function nextCertificateId_(year, code, usedIds) {
  var prefix = 'TSV-' + year + '-' + cleanCode_(code) + '-';
  var max = 0;
  Object.keys(usedIds).forEach(function (id) {
    if (id.indexOf(prefix) === 0) {
      var n = parseInt(id.slice(prefix.length), 10);
      if (isFinite(n) && n > max) max = n;
    }
  });
  var next = prefix + ('000' + (max + 1)).slice(-Math.max(4, String(max + 1).length));
  usedIds[next] = true;
  return next;
}

/**
 * "10–14 March 2027", "28 February – 3 March 2027", "30 December 2026 – 2 January 2027" or "10 March 2027".
 * fmt(date, pattern) formats a date (Utilities.formatDate in Apps Script).
 */
function formatDateRange_(start, end, fmt) {
  var s = isDate_(start) ? start : null;
  var e = isDate_(end) ? end : null;
  if (!s && !e) return '';
  if (!s) s = e;
  if (!e) e = s;
  var day = function (d) { return fmt(d, 'd'); };
  var month = function (d) { return fmt(d, 'MMMM'); };
  var year = function (d) { return fmt(d, 'yyyy'); };
  if (fmt(s, 'yyyy-MM-dd') === fmt(e, 'yyyy-MM-dd')) return day(s) + ' ' + month(s) + ' ' + year(s);
  if (year(s) !== year(e)) return day(s) + ' ' + month(s) + ' ' + year(s) + ' – ' + day(e) + ' ' + month(e) + ' ' + year(e);
  if (month(s) !== month(e)) return day(s) + ' ' + month(s) + ' – ' + day(e) + ' ' + month(e) + ' ' + year(e);
  return day(s) + '–' + day(e) + ' ' + month(e) + ' ' + year(e);
}

/** Course days, from the course list or else counted from the dates (inclusive). */
function courseDays_(courseInfo, start, end) {
  if (courseInfo && Number(courseInfo.days) > 0) return Number(courseInfo.days);
  if (isDate_(start) && isDate_(end)) {
    var d = Math.round((Date.UTC(end.getFullYear(), end.getMonth(), end.getDate()) - Date.UTC(start.getFullYear(), start.getMonth(), start.getDate())) / 86400000) + 1;
    return d > 0 ? d : null;
  }
  return isDate_(start) ? 1 : null;
}

function average_(list) {
  var nums = list.filter(function (n) { return n !== null && isFinite(n); });
  if (!nums.length) return null;
  return nums.reduce(function (a, b) { return a + b; }, 0) / nums.length;
}

function round1_(n) { return n === null ? null : Math.round(n * 10) / 10; }

/* ============================== COURSE LIST ============================== */

/** Reads the public course sheet: { "T1.1": {title, days}, ... } */
function loadCourses_(settings) {
  var id = String(settings.course_sheet_id || '').trim();
  if (!id) throw new Error('Settings: course_sheet_id is empty.');
  var sheet = SpreadsheetApp.openById(id).getSheets()[0];
  var values = sheet.getDataRange().getValues();
  var head = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var iCode = head.indexOf('code'), iTitle = head.indexOf('title'), iDays = head.indexOf('days');
  if (iCode === -1 || iTitle === -1) throw new Error('The course sheet needs "code" and "title" columns.');
  var map = {};
  values.slice(1).forEach(function (r) {
    var code = cleanCode_(r[iCode]);
    if (code) map[code] = { title: String(r[iTitle]).trim(), days: iDays === -1 ? null : Number(r[iDays]) || null };
  });
  return map;
}

/* ============================== MENU: FILL TITLES ============================== */

function menuFillCourseTitles() {
  var reg = registerSheet_();
  var courses = loadCourses_(getSettings_());
  var updates = {}, unknown = {};
  readRows_(reg).forEach(function (r) {
    var code = cleanCode_(r.course_code);
    if (!code || !isBlank_(r.course_title)) return;
    if (courses[code]) updates[r.rowNumber] = courses[code].title;
    else unknown[code] = true;
  });
  writeColumn_(reg, 'course_title', updates);
  var msg = Object.keys(updates).length + ' course title(s) filled in.';
  if (Object.keys(unknown).length) msg += '\n\nThese codes were not found in the course list: ' + Object.keys(unknown).join(', ');
  ui_().alert('Fill course titles', msg, ui_().ButtonSet.OK);
}

/* ============================== MENU: CALCULATE RESULTS ============================== */

function menuCalculateResults() {
  var reg = registerSheet_();
  var settings = getSettings_();
  var results = {}, types = {}, notes = {};
  var done = 0, skippedLocked = 0, missing = [];
  readRows_(reg).forEach(function (r) {
    var status = String(r.status || '').trim();
    if (status === 'Issued' || status === 'Revoked') { skippedLocked++; return; }
    var att = toNumber_(r.attendance_percent, true);
    var pre = toNumber_(r.pre_test_score);
    var post = toNumber_(r.post_test_score);
    var res = computeResult_(att, post, settings);
    if (!res) { missing.push(r.rowNumber); return; }
    results[r.rowNumber] = res.result;
    types[r.rowNumber] = res.certificate_type;
    notes[r.rowNumber] = improvementNote_(r.notes, pre, post);
    done++;
  });
  writeColumn_(reg, 'result', results);
  writeColumn_(reg, 'certificate_type', types);
  writeColumn_(reg, 'notes', notes);
  var msg = done + ' row(s) calculated.';
  if (skippedLocked) msg += '\n' + skippedLocked + ' row(s) already Issued or Revoked were left unchanged.';
  if (missing.length) msg += '\nMissing attendance or post-test score on row(s): ' + missing.join(', ');
  msg += '\n\nRules used: attendance ≥ ' + settings.pass_attendance_percent + '% and post-test ≥ ' + settings.pass_post_test_score +
    ' = Pass. Post-test ≥ ' + settings.competence_post_test_score + ' = Certificate of Competence.';
  ui_().alert('Calculate results', msg, ui_().ButtonSet.OK);
}

/* ============================== MENU: CERTIFICATE IDS ============================== */

function menuCreateCertificateIds() {
  var reg = registerSheet_();
  var created = assignIds_(reg, readRows_(reg), null);
  ui_().alert('Create certificate IDs', created + ' new certificate ID(s) created for passing participants.', ui_().ButtonSet.OK);
}

/** Gives an ID to passing rows without one. onlyRows: optional list of row numbers. Returns count. */
function assignIds_(reg, rows, onlyRows) {
  var used = {};
  rows.forEach(function (r) { var id = String(r.certificate_id || '').trim(); if (id) used[id] = true; });
  var tz = reg.ss.getSpreadsheetTimeZone();
  var updates = {};
  rows.forEach(function (r) {
    if (onlyRows && onlyRows.indexOf(r.rowNumber) === -1) return;
    if (!isBlank_(r.certificate_id) || String(r.result).trim() !== RESULT_PASS) return;
    if (String(r.status).trim() === 'Revoked' || isBlank_(r.course_code)) return;
    var year = isDate_(r.start_date) ? Utilities.formatDate(r.start_date, tz, 'yyyy') : Utilities.formatDate(new Date(), tz, 'yyyy');
    var id = nextCertificateId_(year, r.course_code, used);
    updates[r.rowNumber] = id;
    r.certificate_id = id;
  });
  writeColumn_(reg, 'certificate_id', updates);
  return Object.keys(updates).length;
}

/* ============================== MENU: GENERATE CERTIFICATES ============================== */

function menuGenerateCertificates() {
  var ui = ui_();
  var reg = registerSheet_();
  var settings = getSettings_();
  var verifyBase = verifyBaseUrl_(settings);
  if (!verifyBase) {
    ui.alert('Verification page not set up',
      'Certificates need a QR code that links to your verification page.\n\nPlease deploy the web app and paste its link into Settings > verify_url (INSTALL.md step 5), then try again.',
      ui.ButtonSet.OK);
    return;
  }

  var rows = readRows_(reg);
  var selected = selectedRowNumbers_(reg).filter(function (n) { return rows.some(function (r) { return r.rowNumber === n; }); });
  var targets;
  if (selected.length) {
    var choice = ui.alert('Generate certificates',
      'You have ' + selected.length + ' row(s) selected.\n\nYES = make certificates for the selected rows only (existing certificates for these rows are replaced).\nNO = make certificates for ALL passing participants who have no issued date yet.',
      ui.ButtonSet.YES_NO_CANCEL);
    if (choice === ui.Button.CANCEL || choice === ui.Button.CLOSE) return;
    targets = choice === ui.Button.YES ? selected : null;
  } else {
    targets = null;
  }

  var list = rows.filter(function (r) {
    if (targets) return targets.indexOf(r.rowNumber) !== -1;
    // All mode: only passing participants without a certificate yet (others are expected, not "skipped")
    return isBlank_(r.issued_date) && String(r.result).trim() === RESULT_PASS && String(r.status).trim() !== 'Revoked';
  });
  var skipped = [];
  list = list.filter(function (r) {
    if (String(r.result).trim() !== RESULT_PASS) { skipped.push(r.rowNumber + ' (result is not Pass)'); return false; }
    if (String(r.status).trim() === 'Revoked') { skipped.push(r.rowNumber + ' (revoked)'); return false; }
    if (isBlank_(r.course_code)) { skipped.push(r.rowNumber + ' (no course code)'); return false; }
    return true;
  });
  if (!list.length) {
    ui.alert('Generate certificates', 'There are no certificates to make.' + (skipped.length ? '\n\nSkipped rows: ' + skipped.join('; ') : '') +
      '\n\nTip: run "Calculate results" first so passing rows show "Pass".', ui.ButtonSet.OK);
    return;
  }

  assignIds_(reg, rows, list.map(function (r) { return r.rowNumber; }));

  var courses = {};
  try { courses = loadCourses_(settings); } catch (e) { /* hours fall back to the dates */ }
  var logo = loadLogo_(settings);
  var root = folderPath_([String(settings.certificate_folder_name || 'Tsavone Certificates')]);
  var tz = reg.ss.getSpreadsheetTimeZone();
  var started = Date.now();
  var made = 0, stoppedEarly = false, errors = [];

  for (var i = 0; i < list.length; i++) {
    if (Date.now() - started > TIME_LIMIT_MS) { stoppedEarly = true; break; }
    var r = list[i];
    try {
      var issued = isDate_(r.issued_date) ? r.issued_date : new Date();
      var info = certificateInfo_(r, courses, settings, tz, issued, verifyBase);
      var folder = folderPath_([root.getName(), cleanCode_(r.course_code)], root);
      var pdf = makeCertificatePdf_(info, settings, logo.blob);
      removeExisting_(folder, info.id);
      folder.createFile(pdf.setName(info.id + ' - ' + safeFileName_(info.name) + '.pdf'));
      setCell_(reg, r.rowNumber, 'issued_date', issued);
      setCell_(reg, r.rowNumber, 'status', 'Issued');
      SpreadsheetApp.flush();
      made++;
    } catch (e) {
      errors.push('Row ' + r.rowNumber + ': ' + e.message);
    }
  }

  var msg = made + ' certificate(s) saved to Google Drive > ' + root.getName() + '.\nQR codes link to: ' + verifyBase;
  if (stoppedEarly) msg += '\n\nGoogle limits each run to 6 minutes, so I stopped early. Run "Generate certificates" again (choose NO) to make the rest.';
  if (!logo.blob) msg += '\n\nNote: ' + logo.message;
  if (skipped.length) msg += '\n\nSkipped rows: ' + skipped.join('; ');
  if (errors.length) msg += '\n\nProblems:\n' + errors.join('\n');
  ui.alert('Generate certificates', msg, ui.ButtonSet.OK);
}

function verifyBaseUrl_(settings) {
  var url = String(settings.verify_url || '').trim();
  if (!url) {
    try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
  }
  return url.replace(/[?#].*$/, '');
}

function verifyLink_(base, id) {
  return base + '?id=' + encodeURIComponent(id);
}

function certificateInfo_(r, courses, settings, tz, issued, verifyBase) {
  var code = cleanCode_(r.course_code);
  var course = courses[code] || null;
  var fmt = function (d, p) { return Utilities.formatDate(d, tz, p); };
  var days = courseDays_(course, r.start_date, r.end_date);
  var type = String(r.certificate_type).trim() === 'Competence' ? 'Competence' : 'Completion';
  var id = String(r.certificate_id).trim();
  return {
    id: id,
    name: String(r.participant_name).trim(),
    code: code,
    title: String(r.course_title || (course && course.title) || code).trim(),
    dates: formatDateRange_(r.start_date, r.end_date, fmt),
    location: String(r.location || '').trim(),
    format: String(r.format || '').trim(),
    hours: days ? days * (settings.hours_per_day || 7) : null,
    type: type,
    issued: fmt(issued, 'd MMMM yyyy'),
    verifyUrl: verifyLink_(verifyBase, id)
  };
}

function loadLogo_(settings) {
  var id = String(settings.logo_drive_file_id || '').trim();
  if (!id) return { blob: null, message: 'No logo set in Settings > logo_drive_file_id, so the company name was printed instead.' };
  try {
    var blob = DriveApp.getFileById(id).getBlob();
    var type = blob.getContentType();
    if (['image/png', 'image/jpeg', 'image/gif'].indexOf(type) === -1) {
      return { blob: null, message: 'The logo file must be PNG or JPG (yours is ' + type + '). The company name was printed instead.' };
    }
    return { blob: blob, message: '' };
  } catch (e) {
    return { blob: null, message: 'Could not open the logo file (' + e.message + '). The company name was printed instead.' };
  }
}

/** Finds or creates nested folders in My Drive. names[0] is the top folder. */
function folderPath_(names, rootFolder) {
  var folder = rootFolder || null;
  var start = 0;
  if (folder) start = 1; else folder = DriveApp.getRootFolder();
  for (var i = start; i < names.length; i++) {
    var it = folder.getFoldersByName(names[i]);
    folder = it.hasNext() ? it.next() : folder.createFolder(names[i]);
  }
  return folder;
}

function removeExisting_(folder, id) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf(id + ' - ') === 0) f.setTrashed(true);
  }
}

function findCertificateFile_(settings, code, id) {
  var root = folderPath_([String(settings.certificate_folder_name || 'Tsavone Certificates')]);
  var it = root.getFoldersByName(cleanCode_(code));
  if (!it.hasNext()) return null;
  var files = it.next().getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf(id + ' - ') === 0 && !f.isTrashed()) return f;
  }
  return null;
}

function safeFileName_(s) { return String(s).replace(/[\\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim(); }

/* ---------- Certificate design (Google Doc, exported as PDF) ---------- */

var A4_LANDSCAPE = { width: 841.89, height: 595.28 };
var FONT_HEAD = 'Montserrat';
var FONT_BODY = 'Source Sans 3';

function makeCertificatePdf_(info, settings, logoBlob) {
  var doc = DocumentApp.create('Certificate ' + info.id + ' (temporary)');
  var docFile = DriveApp.getFileById(doc.getId());
  try {
    var body = doc.getBody();
    var margin = 20;
    body.setPageWidth(A4_LANDSCAPE.width).setPageHeight(A4_LANDSCAPE.height);
    body.setMarginTop(margin).setMarginBottom(margin).setMarginLeft(margin).setMarginRight(margin);
    var usable = A4_LANDSCAPE.width - margin * 2;

    // Light-blue border band: an outer one-cell table with a light-blue background
    var outer = body.appendTable([['']]);
    plainTable_(outer, LIGHT_BLUE, usable);
    var band = outer.getCell(0, 0);
    band.setBackgroundColor(LIGHT_BLUE);
    padCell_(band, 16, 16, 18, 18);

    // White card inside the band: orange bar, content, navy bar
    var inner = band.appendTable([[''], [''], ['']]);
    var innerWidth = usable - 36;
    plainTable_(inner, '#ffffff', innerWidth);
    barCell_(inner, 0, ORANGE);
    barCell_(inner, 2, NAVY);
    var c = inner.getCell(1, 0);
    c.setBackgroundColor('#ffffff');
    padCell_(c, 18, 10, 40, 40);
    // Fill the page: page height minus margins, band padding, bars and a safety allowance for Docs spacing
    inner.getRow(1).setMinimumHeight(Math.round(A4_LANDSCAPE.height - margin * 2 - 32 - 14 - 55));
    c.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

    // Logo or name
    var top = c.getChild(0).asParagraph();
    top.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(0);
    if (logoBlob) {
      var img = top.appendInlineImage(logoBlob);
      var h = 66, w = Math.round(img.getWidth() * h / img.getHeight());
      if (w > 260) { w = 260; h = Math.round(img.getHeight() * w / img.getWidth()); }
      img.setHeight(h).setWidth(w);
    } else {
      top.appendText('TSAVONE CONSULTING');
      styleText_(top, FONT_HEAD, 18, NAVY, true);
    }

    var heading = info.type === 'Competence' ? 'CERTIFICATE OF COMPETENCE' : 'CERTIFICATE OF COMPLETION';
    para_(c, heading, { font: FONT_HEAD, size: 26, color: NAVY, bold: true, before: 12 });
    para_(c, '————', { font: FONT_HEAD, size: 12, color: ORANGE, bold: true, before: 2 });
    para_(c, 'This certifies that', { font: FONT_BODY, size: 13, color: GREY, before: 8 });
    para_(c, info.name, { font: FONT_HEAD, size: 32, color: NAVY_DARK, bold: true, before: 6, after: 2 });
    para_(c, info.type === 'Competence'
      ? 'has completed the course and been assessed as competent in'
      : 'has successfully completed the course', { font: FONT_BODY, size: 13, color: GREY, before: 6 });
    para_(c, info.title, { font: FONT_HEAD, size: 19, color: NAVY, bold: true, before: 6 });
    var details = [info.dates, info.location, info.format, info.hours ? info.hours + ' course hours' : '']
      .filter(function (x) { return x; }).join('   ·   ');
    if (details) para_(c, details, { font: FONT_BODY, size: 12, color: '#2c3557', before: 8 });

    // Bottom row: signature | certificate ID and issue date | QR code
    var bottom = c.appendTable([['', '', '']]);
    var colW = innerWidth - 80;
    plainTable_(bottom, '#ffffff', null);
    bottom.setColumnWidth(0, Math.round(colW * 0.38));
    bottom.setColumnWidth(1, Math.round(colW * 0.37));
    bottom.setColumnWidth(2, colW - Math.round(colW * 0.38) - Math.round(colW * 0.37));
    for (var i = 0; i < 3; i++) {
      var cell = bottom.getCell(0, i);
      padCell_(cell, 0, 0, 4, 4);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.BOTTOM);
    }
    var sig = bottom.getCell(0, 0);
    firstPara_(sig, '______________________________', { font: FONT_BODY, size: 12, color: NAVY, before: 24 });
    para_(sig, String(settings.signatory_title || 'Managing Director'), { font: FONT_HEAD, size: 11, color: NAVY_DARK, bold: true, before: 4 });
    para_(sig, 'Tsavone Consulting', { font: FONT_BODY, size: 10, color: GREY });

    var mid = bottom.getCell(0, 1);
    firstPara_(mid, 'CERTIFICATE ID', { font: FONT_HEAD, size: 8, color: ORANGE_DARK, bold: true, before: 18 });
    para_(mid, info.id, { font: FONT_HEAD, size: 12, color: NAVY_DARK, bold: true, before: 1 });
    para_(mid, 'DATE OF ISSUE', { font: FONT_HEAD, size: 8, color: ORANGE_DARK, bold: true, before: 8 });
    para_(mid, info.issued, { font: FONT_BODY, size: 11, color: NAVY_DARK, before: 1 });

    var qrCell = bottom.getCell(0, 2);
    var qrP = qrCell.getChild(0).asParagraph();
    qrP.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(10).setSpacingAfter(0);
    var qr = qrP.appendInlineImage(qrPngBlob(info.verifyUrl, 'qr.png'));
    qr.setWidth(78).setHeight(78);
    para_(qrCell, 'Scan to verify', { font: FONT_BODY, size: 8, color: GREY, before: 0 });

    para_(c, 'Tsavone Consulting  |  www.tsavone.com  |  info@tsavone.com', { font: FONT_BODY, size: 9, color: GREY, before: 10 });

    // Docs keeps empty paragraphs around tables: make them tiny so everything stays on one page
    shrinkEmptyParagraphs_(body);
    shrinkEmptyParagraphs_(band);
    shrinkEmptyParagraphs_(c);

    doc.saveAndClose();
    return docFile.getAs('application/pdf');
  } finally {
    docFile.setTrashed(true);
  }
}

function plainTable_(table, colour, width) {
  table.setBorderWidth(0);
  table.setBorderColor(colour);
  if (width) table.setColumnWidth(0, width);
}

function padCell_(cell, top, bottom, left, right) {
  cell.setPaddingTop(top).setPaddingBottom(bottom).setPaddingLeft(left).setPaddingRight(right);
}

function barCell_(table, row, colour) {
  var cell = table.getCell(row, 0);
  cell.setBackgroundColor(colour);
  padCell_(cell, 0, 0, 0, 0);
  tinyParagraph_(cell.getChild(0).asParagraph());
  table.getRow(row).setMinimumHeight(7);
}

function tinyParagraph_(p) {
  p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  var attrs = {};
  attrs[DocumentApp.Attribute.FONT_SIZE] = 1;
  p.setAttributes(attrs);
}

function shrinkEmptyParagraphs_(container) {
  for (var i = 0; i < container.getNumChildren(); i++) {
    var el = container.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var p = el.asParagraph();
      if (p.getText() === '' && p.getNumChildren() === 0) tinyParagraph_(p);
    }
  }
}

function styleText_(p, font, size, colour, bold) {
  var t = p.editAsText();
  t.setFontSize(size).setForegroundColor(colour).setBold(!!bold);
  try { t.setFontFamily(font); } catch (e) { t.setFontFamily('Arial'); }
}

function para_(container, text, o) {
  var p = container.appendParagraph(String(text));
  formatPara_(p, o);
  return p;
}

/** Uses the empty first paragraph that every new table cell starts with. */
function firstPara_(cell, text, o) {
  var p = cell.getChild(0).asParagraph();
  p.setText(String(text));
  formatPara_(p, o);
  return p;
}

function formatPara_(p, o) {
  p.setAlignment(o.align || DocumentApp.HorizontalAlignment.CENTER);
  p.setSpacingBefore(o.before || 0).setSpacingAfter(o.after || 0).setLineSpacing(1.05);
  if (p.getText()) styleText_(p, o.font, o.size, o.color, o.bold);
}

/* ============================== MENU: EMAIL DRAFTS ============================== */

function menuEmailCertificates() {
  var ui = ui_();
  var reg = registerSheet_();
  var settings = getSettings_();
  var selected = selectedRowNumbers_(reg);
  if (!selected.length) {
    ui.alert('Email certificates', 'First select the participant rows on the Register tab (click the row numbers on the left), then choose this menu item again.', ui.ButtonSet.OK);
    return;
  }
  var tz = reg.ss.getSpreadsheetTimeZone();
  var verifyBase = verifyBaseUrl_(settings);
  var courses = {};
  try { courses = loadCourses_(settings); } catch (e) { /* not needed for drafts */ }
  var rows = readRows_(reg).filter(function (r) { return selected.indexOf(r.rowNumber) !== -1; });
  var made = 0, skipped = [];
  rows.forEach(function (r) {
    var email = String(r.email || '').trim();
    if (String(r.status).trim() !== 'Issued') { skipped.push(r.rowNumber + ' (status is not Issued)'); return; }
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { skipped.push(r.rowNumber + ' (no valid email)'); return; }
    var file = findCertificateFile_(settings, r.course_code, String(r.certificate_id).trim());
    if (!file) { skipped.push(r.rowNumber + ' (certificate PDF not found; generate it first)'); return; }
    var info = certificateInfo_(r, courses, settings, tz, isDate_(r.issued_date) ? r.issued_date : new Date(), verifyBase);
    var msg = certificateEmail_(info);
    GmailApp.createDraft(email, msg.subject, msg.text, {
      htmlBody: msg.html,
      name: String(settings.sender_email_name || 'Tsavone Consulting Training'),
      attachments: [file.getBlob()]
    });
    made++;
  });
  var text = made + ' Gmail draft(s) created. Open Gmail > Drafts to review and send them.';
  if (skipped.length) text += '\n\nSkipped rows: ' + skipped.join('; ');
  ui.alert('Email certificates', text, ui.ButtonSet.OK);
}

function certificateEmail_(info) {
  var first = info.name.split(/\s+/)[0] || info.name;
  var kind = info.type === 'Competence' ? 'Certificate of Competence' : 'Certificate of Completion';
  var when = [info.dates, info.location].filter(function (x) { return x; }).join(', ');
  var subject = 'Your Tsavone certificate: ' + info.title;
  var hasLink = info.verifyUrl.indexOf('http') === 0;
  var lines = [
    'Dear ' + first + ',',
    '',
    'Congratulations on completing ' + info.title + (when ? ' (' + when + ')' : '') + '. Your ' + kind + ' is attached.',
    '',
    'Certificate ID: ' + info.id,
    hasLink ? 'Anyone, including an employer, can confirm it at: ' + info.verifyUrl : '',
    '',
    'Thank you for training with Tsavone Consulting. We would be glad to hear how you are putting your action plan into practice.',
    '',
    'Kind regards,',
    'Tsavone Consulting Training',
    'info@tsavone.com | +256 704 579 186 | www.tsavone.com'
  ];
  var esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var html =
    '<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.55;color:#2c3557">' +
    '<p>Dear ' + esc(first) + ',</p>' +
    '<p>Congratulations on completing <b>' + esc(info.title) + '</b>' + (when ? ' (' + esc(when) + ')' : '') + '. Your ' + kind + ' is attached.</p>' +
    '<p style="background:#EFFCFF;border-left:4px solid #F17B19;padding:10px 14px">Certificate ID: <b>' + esc(info.id) + '</b>' +
    (hasLink ? '<br>Anyone, including an employer, can confirm it here: <a href="' + esc(info.verifyUrl) + '" style="color:#193591">verify certificate</a>' : '') + '</p>' +
    '<p>Thank you for training with Tsavone Consulting. We would be glad to hear how you are putting your action plan into practice.</p>' +
    '<p>Kind regards,<br><b style="color:#193591">Tsavone Consulting Training</b><br>' +
    '<span style="color:#5a6385">info@tsavone.com | +256 704 579 186 | www.tsavone.com</span></p></div>';
  return { subject: subject, text: lines.join('\n'), html: html };
}

/* ============================== MENU: COURSE SUMMARY ============================== */

function menuCourseSummary() {
  var ui = ui_();
  var reg = registerSheet_();
  var tz = reg.ss.getSpreadsheetTimeZone();
  var rows = readRows_(reg).filter(function (r) { return !isBlank_(r.course_code); });
  var runs = courseRuns_(rows, function (d) { return isDate_(d) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : ''; });
  if (!runs.length) { ui.alert('Course summary', 'No participants with a course code were found.', ui.ButtonSet.OK); return; }
  var shown = runs.slice(0, 30);
  var list = shown.map(function (run, i) {
    return (i + 1) + '.  ' + run.code + '   ' + (run.start ? Utilities.formatDate(run.startDate, tz, 'd MMM yyyy') : '(no start date)') + '   – ' + run.rows.length + ' participant(s)';
  }).join('\n');
  var resp = ui.prompt('Course summary', 'Which course? Type the number:\n\n' + list, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var n = parseInt(resp.getResponseText(), 10);
  if (!(n >= 1 && n <= shown.length)) { ui.alert('Please type one of the numbers shown.'); return; }
  var run = shown[n - 1];
  var stats = courseStats_(run.rows, reg.cols);
  writeSummarySheet_(reg.ss, run, stats, tz);
}

/** Groups rows by course code + start date, newest first. */
function courseRuns_(rows, dayKey) {
  var map = {};
  rows.forEach(function (r) {
    var code = cleanCode_(r.course_code), start = dayKey(r.start_date);
    var key = code + '|' + start;
    if (!map[key]) map[key] = { code: code, start: start, startDate: r.start_date, rows: [] };
    map[key].rows.push(r);
  });
  return Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) { return a.start < b.start ? 1 : a.start > b.start ? -1 : a.code.localeCompare(b.code); });
}

/** Works out the numbers for a completion report. */
function courseStats_(rows, cols) {
  var att = [], pre = [], post = [], imp = [];
  var pass = 0, fail = 0, issued = 0, competence = 0, revoked = 0;
  var byOrg = {}, bySex = {};
  var sexCol = cols.sex !== undefined ? 'sex' : cols.gender !== undefined ? 'gender' : null;
  rows.forEach(function (r) {
    var a = toNumber_(r.attendance_percent, true), p = toNumber_(r.pre_test_score), q = toNumber_(r.post_test_score);
    att.push(a); pre.push(p); post.push(q);
    if (p !== null && q !== null) imp.push(q - p);
    var res = String(r.result || '').trim();
    if (res === RESULT_PASS) pass++; else if (res === RESULT_FAIL) fail++;
    var st = String(r.status || '').trim();
    if (st === 'Issued') issued++;
    if (st === 'Revoked') revoked++;
    if (st === 'Issued' && String(r.certificate_type).trim() === 'Competence') competence++;
    var org = String(r.organisation || '').trim() || '(not given)';
    byOrg[org] = (byOrg[org] || 0) + 1;
    if (sexCol) { var s = String(r[sexCol] || '').trim() || '(not given)'; bySex[s] = (bySex[s] || 0) + 1; }
  });
  return {
    participants: rows.length, pass: pass, fail: fail, issued: issued, competence: competence, revoked: revoked,
    passRate: rows.length ? pass / rows.length : null,
    avgAttendance: round1_(average_(att)), avgPre: round1_(average_(pre)), avgPost: round1_(average_(post)), avgImprovement: round1_(average_(imp)),
    byOrg: byOrg, bySex: sexCol ? bySex : null, sexLabel: sexCol
  };
}

function writeSummarySheet_(ss, run, s, tz) {
  var first = run.rows[0];
  var title = String(first.course_title || run.code).trim();
  var dateText = formatDateRange_(first.start_date, first.end_date, function (d, p) { return Utilities.formatDate(d, tz, p); });
  var name = ('Summary ' + run.code + ' ' + (run.start || 'no date')).slice(0, 90);
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet(name);
  var pct = function (x) { return x === null ? '–' : Math.round(x * 100) + '%'; };
  var val = function (x, suffix) { return x === null ? '–' : x + (suffix || ''); };

  var out = [];
  out.push(['Course summary', '']);
  out.push([run.code + ' – ' + title, '']);
  out.push([[dateText, String(first.location || '').trim(), String(first.format || '').trim()].filter(String).join(' · '), '']);
  out.push(['', '']);
  out.push(['Indicator', 'Result']);
  var tableStart = out.length;
  out.push(['Participants', s.participants]);
  out.push(['Met completion requirements (Pass)', s.pass]);
  out.push(['Did not complete', s.fail]);
  out.push(['Pass rate', pct(s.passRate)]);
  out.push(['Average attendance', val(s.avgAttendance, '%')]);
  out.push(['Average pre-test score', val(s.avgPre)]);
  out.push(['Average post-test score', val(s.avgPost)]);
  out.push(['Average improvement (points)', s.avgImprovement === null ? '–' : (s.avgImprovement > 0 ? '+' : '') + s.avgImprovement]);
  out.push(['Certificates issued', s.issued]);
  out.push(['…of which Certificates of Competence', s.competence]);
  if (s.revoked) out.push(['Certificates revoked', s.revoked]);
  var tableEnd = out.length;

  var sentence = s.participants + ' participant' + (s.participants === 1 ? '' : 's') + ' took part; ' + s.pass + ' (' + pct(s.passRate) + ') met the completion requirements.';
  if (s.avgPre !== null && s.avgPost !== null) sentence += ' Average test scores rose from ' + s.avgPre + ' to ' + s.avgPost + ' (' + (s.avgImprovement > 0 ? '+' : '') + s.avgImprovement + ' points).';
  if (s.avgAttendance !== null) sentence += ' Average attendance was ' + s.avgAttendance + '%.';
  out.push(['', '']);
  out.push(['Summary sentence for the completion report', '']);
  out.push([sentence, '']);

  var blocks = [];
  if (s.bySex) blocks.push([s.sexLabel === 'gender' ? 'Participants by gender' : 'Participants by sex', s.bySex]);
  blocks.push(['Participants by organisation', s.byOrg]);
  var blockHeads = [];
  blocks.forEach(function (b) {
    out.push(['', '']);
    blockHeads.push(out.length + 1);
    out.push([b[0], 'Number']);
    Object.keys(b[1]).sort(function (x, y) { return b[1][y] - b[1][x] || x.localeCompare(y); })
      .forEach(function (k) { out.push([k, b[1][k]]); });
  });

  sh.getRange(1, 1, out.length, 2).setValues(out);
  sh.setColumnWidth(1, 360); sh.setColumnWidth(2, 160);
  sh.getRange(1, 1).setFontSize(11).setFontColor(ORANGE_DARK).setFontWeight('bold');
  sh.getRange(2, 1).setFontSize(16).setFontColor(NAVY).setFontWeight('bold');
  sh.getRange(3, 1).setFontColor(GREY);
  [tableStart].concat(blockHeads).forEach(function (r) {
    sh.getRange(r, 1, 1, 2).setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold');
  });
  sh.getRange(tableStart + 1, 1, tableEnd - tableStart, 2).setBorder(true, true, true, true, true, true, '#dbe3f0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(tableStart + 1, 2, tableEnd - tableStart, 1).setHorizontalAlignment('right');
  var sentRow = tableEnd + 3;
  sh.getRange(sentRow - 1, 1).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange(sentRow, 1, 1, 2).merge().setWrap(true).setBackground(LIGHT_BLUE);
  sh.setRowHeight(sentRow, 60);
  ss.setActiveSheet(sh);
}

/* ============================== MENU: TESTING ============================== */

var TEST_NOTE = 'TEST ROW - delete after testing';

/** Email address for test rows: your own address with +test1, +test2 ... (arrives in your inbox). */
function testEmail_(n) {
  var me = '';
  try { me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || ''; } catch (e) { me = ''; }
  var at = me.indexOf('@');
  return at > 0 ? me.slice(0, at) + '+test' + n + me.slice(at) : 'test' + n + '@example.com';
}

/** Adds three test participants (two pass, one does not) so every menu item can be tried. */
function menuAddTestParticipants() {
  var reg = registerSheet_();
  var rows = readRows_(reg);
  if (rows.some(function (r) { return String(r.notes || '').indexOf(TEST_NOTE) === 0; })) {
    ui_().alert('Add test participants', 'Test participants are already in the Register. Use Tsavone > Testing > Remove test participants first if you want to start again.', ui_().ButtonSet.OK);
    return;
  }
  var start = new Date(2027, 2, 10), end = new Date(2027, 2, 14);
  var people = [
    ['Test Participant One', 'Sample Organisation A', 100, 45, 82],
    ['Test Participant Two', 'Sample Organisation B', 90, 50, 68],
    ['Test Participant Three', 'Sample Organisation B', 60, 40, 55]
  ];
  var width = reg.sh.getLastColumn();
  var values = people.map(function (p, i) {
    var row = new Array(width).fill('');
    var set = function (col, v) { row[reg.cols[col]] = v; };
    set('participant_name', p[0]); set('organisation', p[1]); set('email', testEmail_(i + 1));
    set('phone', '+25670000000' + (i + 1)); set('course_code', 'T1.1');
    set('start_date', start); set('end_date', end); set('location', 'Kampala, Uganda'); set('format', 'Open enrolment');
    set('attendance_percent', p[2]); set('pre_test_score', p[3]); set('post_test_score', p[4]);
    set('status', 'Draft'); set('notes', TEST_NOTE);
    return row;
  });
  var first = Math.max(reg.sh.getLastRow(), 1) + 1;
  reg.sh.getRange(first, 1, values.length, width).setValues(values);
  reg.sh.setRowHeights(first, values.length, 21);
  reg.sh.getRange(first, 1, values.length, width).setWrap(false);
  ui_().alert('Add test participants',
    '3 test participants added (rows ' + first + '–' + (first + 2) + '). Their emails go to your own inbox.\n\nNow run, in order:\n1. Fill course titles from codes\n2. Calculate results\n3. Create certificate IDs\n4. Generate certificates (choose NO)\n\nWhen finished, use Tsavone > Testing > Remove test participants.',
    ui_().ButtonSet.OK);
}

/** Removes test rows, their certificate PDFs, their Gmail drafts and their summary tab. */
function menuRemoveTestParticipants() {
  var ui = ui_();
  var reg = registerSheet_();
  var settings = getSettings_();
  var tests = readRows_(reg).filter(function (r) { return String(r.notes || '').indexOf(TEST_NOTE) === 0; });
  if (!tests.length) { ui.alert('Remove test participants', 'No test participants found.', ui.ButtonSet.OK); return; }
  var ok = ui.alert('Remove test participants', 'Delete the ' + tests.length + ' test row(s), their certificate PDFs, their Gmail drafts and the test summary tab?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  var pdfs = 0, drafts = 0, tabs = 0;
  var emails = {};
  tests.forEach(function (r) {
    emails[String(r.email || '').trim().toLowerCase()] = true;
    var id = String(r.certificate_id || '').trim();
    if (id) {
      try { var f = findCertificateFile_(settings, r.course_code, id); if (f) { f.setTrashed(true); pdfs++; } } catch (e) { /* folder may not exist */ }
    }
  });
  try {
    GmailApp.getDrafts().forEach(function (d) {
      var m = d.getMessage();
      var to = String(m.getTo() || '').toLowerCase();
      if (m.getSubject().indexOf('Your Tsavone certificate') === 0 && Object.keys(emails).some(function (e) { return e && to.indexOf(e) !== -1; })) {
        d.deleteDraft(); drafts++;
      }
    });
  } catch (e) { /* Gmail not available: drafts stay */ }

  tests.map(function (r) { return r.rowNumber; }).sort(function (a, b) { return b - a; })
    .forEach(function (n) { reg.sh.deleteRow(n); });

  // Summary tabs for the test course run, unless real participants share it
  var tz = reg.ss.getSpreadsheetTimeZone();
  var remaining = readRows_(reg);
  var runKeys = {};
  tests.forEach(function (r) {
    if (isDate_(r.start_date)) runKeys[cleanCode_(r.course_code) + ' ' + Utilities.formatDate(r.start_date, tz, 'yyyy-MM-dd')] = true;
  });
  Object.keys(runKeys).forEach(function (k) {
    var stillUsed = remaining.some(function (r) {
      return isDate_(r.start_date) && cleanCode_(r.course_code) + ' ' + Utilities.formatDate(r.start_date, tz, 'yyyy-MM-dd') === k;
    });
    var sh = reg.ss.getSheetByName('Summary ' + k);
    if (sh && !stillUsed) { reg.ss.deleteSheet(sh); tabs++; }
  });
  reg.ss.setActiveSheet(reg.sh);
  ui.alert('Remove test participants',
    'Removed ' + tests.length + ' test row(s), ' + pdfs + ' certificate PDF(s) (moved to Drive Bin), ' + drafts + ' Gmail draft(s) and ' + tabs + ' summary tab(s).',
    ui.ButtonSet.OK);
}

/* ============================== VERIFICATION WEB APP ============================== */

function doGet(e) {
  var id = String((e && e.parameter && e.parameter.id) || '').trim().slice(0, 60);
  rememberVerifyUrl_();
  var t = HtmlService.createTemplateFromFile('Verify');
  t.query = id;
  t.result = id ? lookupCertificate_(id) : null;
  t.pageUrl = ScriptApp.getService().getUrl() || '';
  t.logo = verifyLogoDataUri_();
  return t.evaluate()
    .setTitle('Verify a Tsavone certificate')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** The first time the published page is opened, save its link into Settings > verify_url (if empty). */
function rememberVerifyUrl_() {
  try {
    var url = ScriptApp.getService().getUrl() || '';
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) return;
    var ss = registerSpreadsheet_();
    if (!String(readSettings_(ss).verify_url || '').trim()) setSetting_(ss, 'verify_url', url);
  } catch (e) { /* never block the page */ }
}

function registerSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Only ever returns public, non-sensitive fields. */
function lookupCertificate_(id) {
  var ss = registerSpreadsheet_();
  var sh = ss.getSheetByName(REGISTER_SHEET);
  if (!sh || sh.getLastRow() < 2) return { state: 'NOT_FOUND', id: id };
  var cols = headerMap_(sh);
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var wanted = id.toUpperCase();
  var tz = ss.getSpreadsheetTimeZone();
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    if (String(v[cols.certificate_id]).trim().toUpperCase() !== wanted) continue;
    var status = String(v[cols.status]).trim();
    var realId = String(v[cols.certificate_id]).trim();
    if (status === 'Revoked') return { state: 'REVOKED', id: realId };
    if (status !== 'Issued') return { state: 'NOT_FOUND', id: id };
    return {
      state: 'VALID',
      id: realId,
      name: String(v[cols.participant_name]).trim(),
      course: String(v[cols.course_title] || v[cols.course_code]).trim(),
      dates: formatDateRange_(v[cols.start_date], v[cols.end_date], function (d, p) { return Utilities.formatDate(d, tz, p); }),
      type: String(v[cols.certificate_type]).trim() === 'Competence' ? 'Certificate of Competence' : 'Certificate of Completion'
    };
  }
  return { state: 'NOT_FOUND', id: id };
}

function verifyLogoDataUri_() {
  try {
    var settings = readSettings_(registerSpreadsheet_());
    var logo = loadLogo_(settings);
    if (!logo.blob) return '';
    return 'data:' + logo.blob.getContentType() + ';base64,' + Utilities.base64Encode(logo.blob.getBytes());
  } catch (e) {
    return '';
  }
}

/* ======================================================================
 * QR CODES AND PNG IMAGES (no outside services)
 * ====================================================================== */


var QR_ECL_M_ = 1; // index into the tables below: 0 = L, 1 = M, 2 = Q, 3 = H
var QR_ECL_FORMAT_BITS_ = [1, 0, 3, 2];

var QR_ECC_CODEWORDS_PER_BLOCK_ = [
  [-1, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28, 28, 28, 30, 30, 26, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
  [-1, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28],
  [-1, 13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30, 28, 30, 30, 30, 30, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
  [-1, 17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28, 30, 24, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
];
var QR_NUM_EC_BLOCKS_ = [
  [-1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 24, 25],
  [-1, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49],
  [-1, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43, 45, 48, 51, 53, 56, 59, 62, 65, 68],
  [-1, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48, 51, 54, 57, 60, 63, 66, 70, 74, 77, 81]
];

/** Returns a square array of booleans (true = dark module) encoding the text. */
function qrMatrix(text) {
  var data = utf8Bytes_(String(text));
  var ecl = QR_ECL_M_;
  var version, capacityBits, ccBits;
  for (version = 1; version <= 40; version++) {
    capacityBits = qrNumDataCodewords_(version, ecl) * 8;
    ccBits = version < 10 ? 8 : 16;
    if (4 + ccBits + data.length * 8 <= capacityBits) break;
  }
  if (version > 40) throw new Error('Text is too long for a QR code.');

  // Bit stream: byte mode, length, data, terminator, padding
  var bits = [];
  var push = function (val, len) { for (var i = len - 1; i >= 0; i--) bits.push((val >>> i) & 1); };
  push(4, 4);
  push(data.length, ccBits);
  data.forEach(function (b) { push(b, 8); });
  push(0, Math.min(4, capacityBits - bits.length));
  push(0, (8 - bits.length % 8) % 8);
  for (var pad = 0xEC; bits.length < capacityBits; pad ^= 0xEC ^ 0x11) push(pad, 8);
  var codewords = [];
  for (var i = 0; i < bits.length; i += 8) {
    var b = 0;
    for (var j = 0; j < 8; j++) b = (b << 1) | bits[i + j];
    codewords.push(b);
  }

  var qr = new QrBuilder_(version, ecl);
  qr.drawFunctionPatterns();
  qr.drawCodewords(qr.addEccAndInterleave(codewords));

  // Pick the mask with the lowest penalty
  var bestMask = 0, bestPenalty = Infinity;
  for (var m = 0; m < 8; m++) {
    qr.applyMask(m);
    qr.drawFormatBits(m);
    var p = qr.penalty();
    if (p < bestPenalty) { bestPenalty = p; bestMask = m; }
    qr.applyMask(m); // undo (XOR)
  }
  qr.applyMask(bestMask);
  qr.drawFormatBits(bestMask);
  return qr.modules;
}

function QrBuilder_(version, ecl) {
  this.version = version;
  this.ecl = ecl;
  this.size = version * 4 + 17;
  this.modules = [];
  this.isFunction = [];
  for (var y = 0; y < this.size; y++) {
    this.modules.push(new Array(this.size).fill(false));
    this.isFunction.push(new Array(this.size).fill(false));
  }
}

QrBuilder_.prototype.setFunction = function (x, y, dark) {
  this.modules[y][x] = dark;
  this.isFunction[y][x] = true;
};

QrBuilder_.prototype.drawFunctionPatterns = function () {
  var size = this.size, i, j;
  for (i = 0; i < size; i++) {
    this.setFunction(6, i, i % 2 === 0);
    this.setFunction(i, 6, i % 2 === 0);
  }
  this.drawFinder(3, 3);
  this.drawFinder(size - 4, 3);
  this.drawFinder(3, size - 4);
  var pos = qrAlignmentPositions_(this.version);
  var n = pos.length;
  for (i = 0; i < n; i++) {
    for (j = 0; j < n; j++) {
      if ((i === 0 && j === 0) || (i === 0 && j === n - 1) || (i === n - 1 && j === 0)) continue;
      this.drawAlignment(pos[i], pos[j]);
    }
  }
  this.drawFormatBits(0);
  this.drawVersion();
};

QrBuilder_.prototype.drawFinder = function (x, y) {
  for (var dy = -4; dy <= 4; dy++) {
    for (var dx = -4; dx <= 4; dx++) {
      var dist = Math.max(Math.abs(dx), Math.abs(dy));
      var xx = x + dx, yy = y + dy;
      if (xx >= 0 && xx < this.size && yy >= 0 && yy < this.size) this.setFunction(xx, yy, dist !== 2 && dist !== 4);
    }
  }
};

QrBuilder_.prototype.drawAlignment = function (x, y) {
  for (var dy = -2; dy <= 2; dy++) {
    for (var dx = -2; dx <= 2; dx++) this.setFunction(x + dx, y + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);
  }
};

QrBuilder_.prototype.drawFormatBits = function (mask) {
  var data = (QR_ECL_FORMAT_BITS_[this.ecl] << 3) | mask;
  var rem = data;
  for (var i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
  var bits = ((data << 10) | rem) ^ 0x5412;
  var bit = function (k) { return ((bits >>> k) & 1) !== 0; };
  var size = this.size;
  for (i = 0; i <= 5; i++) this.setFunction(8, i, bit(i));
  this.setFunction(8, 7, bit(6));
  this.setFunction(8, 8, bit(7));
  this.setFunction(7, 8, bit(8));
  for (i = 9; i < 15; i++) this.setFunction(14 - i, 8, bit(i));
  for (i = 0; i < 8; i++) this.setFunction(size - 1 - i, 8, bit(i));
  for (i = 8; i < 15; i++) this.setFunction(8, size - 15 + i, bit(i));
  this.setFunction(8, size - 8, true);
};

QrBuilder_.prototype.drawVersion = function () {
  if (this.version < 7) return;
  var rem = this.version;
  for (var i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >>> 11) * 0x1F25);
  var bits = (this.version << 12) | rem;
  for (i = 0; i < 18; i++) {
    var dark = ((bits >>> i) & 1) !== 0;
    var a = this.size - 11 + i % 3, b = Math.floor(i / 3);
    this.setFunction(a, b, dark);
    this.setFunction(b, a, dark);
  }
};

QrBuilder_.prototype.addEccAndInterleave = function (data) {
  var ver = this.version, ecl = this.ecl;
  var numBlocks = QR_NUM_EC_BLOCKS_[ecl][ver];
  var blockEccLen = QR_ECC_CODEWORDS_PER_BLOCK_[ecl][ver];
  var rawCodewords = Math.floor(qrNumRawDataModules_(ver) / 8);
  var numShortBlocks = numBlocks - rawCodewords % numBlocks;
  var shortBlockLen = Math.floor(rawCodewords / numBlocks);
  var divisor = rsDivisor_(blockEccLen);
  var blocks = [];
  for (var i = 0, k = 0; i < numBlocks; i++) {
    var len = shortBlockLen - blockEccLen + (i < numShortBlocks ? 0 : 1);
    var dat = data.slice(k, k + len);
    k += len;
    var ecc = rsRemainder_(dat, divisor);
    if (i < numShortBlocks) dat.push(0);
    blocks.push(dat.concat(ecc));
  }
  var result = [];
  for (i = 0; i < blocks[0].length; i++) {
    for (var j = 0; j < blocks.length; j++) {
      if (i !== shortBlockLen - blockEccLen || j >= numShortBlocks) result.push(blocks[j][i]);
    }
  }
  return result;
};

QrBuilder_.prototype.drawCodewords = function (data) {
  var i = 0, size = this.size;
  for (var right = size - 1; right >= 1; right -= 2) {
    if (right === 6) right = 5;
    for (var vert = 0; vert < size; vert++) {
      for (var j = 0; j < 2; j++) {
        var x = right - j;
        var upward = ((right + 1) & 2) === 0;
        var y = upward ? size - 1 - vert : vert;
        if (!this.isFunction[y][x] && i < data.length * 8) {
          this.modules[y][x] = ((data[i >>> 3] >>> (7 - (i & 7))) & 1) !== 0;
          i++;
        }
      }
    }
  }
};

QrBuilder_.prototype.applyMask = function (mask) {
  for (var y = 0; y < this.size; y++) {
    for (var x = 0; x < this.size; x++) {
      var invert;
      switch (mask) {
        case 0: invert = (x + y) % 2 === 0; break;
        case 1: invert = y % 2 === 0; break;
        case 2: invert = x % 3 === 0; break;
        case 3: invert = (x + y) % 3 === 0; break;
        case 4: invert = (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0; break;
        case 5: invert = (x * y) % 2 + (x * y) % 3 === 0; break;
        case 6: invert = ((x * y) % 2 + (x * y) % 3) % 2 === 0; break;
        default: invert = ((x + y) % 2 + (x * y) % 3) % 2 === 0;
      }
      if (!this.isFunction[y][x] && invert) this.modules[y][x] = !this.modules[y][x];
    }
  }
};

/** Standard QR penalty score: lower is easier to scan. */
QrBuilder_.prototype.penalty = function () {
  var size = this.size, m = this.modules, result = 0, x, y, run, dark = 0;
  var lines = [];
  for (y = 0; y < size; y++) {
    var row = [], col = [];
    for (x = 0; x < size; x++) { row.push(m[y][x]); col.push(m[x][y]); }
    lines.push(row, col);
  }
  lines.forEach(function (line) {
    // Rule 1: runs of five or more of the same colour
    run = 1;
    for (var i = 1; i <= line.length; i++) {
      if (i < line.length && line[i] === line[i - 1]) run++;
      else { if (run >= 5) result += run - 2; run = 1; }
    }
    // Rule 3: finder-like patterns 1:1:3:1:1 with four light modules on one side
    for (i = 0; i + 6 < line.length; i++) {
      if (line[i] && !line[i + 1] && line[i + 2] && line[i + 3] && line[i + 4] && !line[i + 5] && line[i + 6]) {
        var before = true, after = true;
        for (var k = 1; k <= 4; k++) {
          if (i - k >= 0 && line[i - k]) before = false;
          if (i + 6 + k < line.length && line[i + 6 + k]) after = false;
        }
        if (before || after) result += 40;
      }
    }
  });
  // Rule 2: 2x2 blocks of one colour
  for (y = 0; y < size - 1; y++) {
    for (x = 0; x < size - 1; x++) {
      var c = m[y][x];
      if (c === m[y][x + 1] && c === m[y + 1][x] && c === m[y + 1][x + 1]) result += 3;
    }
  }
  // Rule 4: balance of dark and light
  for (y = 0; y < size; y++) for (x = 0; x < size; x++) if (m[y][x]) dark++;
  var total = size * size;
  result += Math.max(0, Math.ceil(Math.abs(dark * 20 - total * 10) / total) - 1) * 10;
  return result;
};

function qrAlignmentPositions_(ver) {
  if (ver === 1) return [];
  var num = Math.floor(ver / 7) + 2;
  var step = ver === 32 ? 26 : Math.ceil((ver * 4 + 4) / (num * 2 - 2)) * 2;
  var result = [6];
  for (var pos = ver * 4 + 17 - 7; result.length < num; pos -= step) result.splice(1, 0, pos);
  return result;
}

function qrNumRawDataModules_(ver) {
  var result = (16 * ver + 128) * ver + 64;
  if (ver >= 2) {
    var numAlign = Math.floor(ver / 7) + 2;
    result -= (25 * numAlign - 10) * numAlign - 55;
    if (ver >= 7) result -= 36;
  }
  return result;
}

function qrNumDataCodewords_(ver, ecl) {
  return Math.floor(qrNumRawDataModules_(ver) / 8) - QR_ECC_CODEWORDS_PER_BLOCK_[ecl][ver] * QR_NUM_EC_BLOCKS_[ecl][ver];
}

function rsMultiply_(x, y) {
  var z = 0;
  for (var i = 7; i >= 0; i--) {
    z = (z << 1) ^ ((z >>> 7) * 0x11D);
    z ^= ((y >>> i) & 1) * x;
  }
  return z;
}

function rsDivisor_(degree) {
  var result = [];
  for (var i = 0; i < degree - 1; i++) result.push(0);
  result.push(1);
  var root = 1;
  for (i = 0; i < degree; i++) {
    for (var j = 0; j < result.length; j++) {
      result[j] = rsMultiply_(result[j], root);
      if (j + 1 < result.length) result[j] ^= result[j + 1];
    }
    root = rsMultiply_(root, 0x02);
  }
  return result;
}

function rsRemainder_(data, divisor) {
  var result = divisor.map(function () { return 0; });
  data.forEach(function (b) {
    var factor = b ^ result.shift();
    result.push(0);
    divisor.forEach(function (coef, i) { result[i] ^= rsMultiply_(coef, factor); });
  });
  return result;
}

function utf8Bytes_(str) {
  var out = [];
  for (var i = 0; i < str.length; i++) {
    var c = str.charCodeAt(i);
    if (c >= 0xD800 && c <= 0xDBFF && i + 1 < str.length) {
      c = 0x10000 + ((c - 0xD800) << 10) + (str.charCodeAt(++i) - 0xDC00);
    }
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xC0 | (c >> 6), 0x80 | (c & 63));
    else if (c < 0x10000) out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    else out.push(0xF0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return out;
}

/* ---------- PNG maker (greyscale, no compression needed) ---------- */

/**
 * Turns a QR matrix into PNG bytes (numbers 0–255).
 * scale = pixels per module; a 4-module white border is added, as the QR standard requires.
 */
function qrPngBytes(matrix, scale) {
  scale = scale || 8;
  var quiet = 4;
  var n = matrix.length;
  var px = (n + quiet * 2) * scale;
  var raw = [];
  for (var y = 0; y < px; y++) {
    raw.push(0); // filter type: none
    var my = Math.floor(y / scale) - quiet;
    for (var x = 0; x < px; x++) {
      var mx = Math.floor(x / scale) - quiet;
      var dark = my >= 0 && my < n && mx >= 0 && mx < n && matrix[my][mx];
      raw.push(dark ? 0 : 255);
    }
  }
  var ihdr = [].concat(u32_(px), u32_(px), [8, 0, 0, 0, 0]); // 8-bit greyscale
  return [137, 80, 78, 71, 13, 10, 26, 10]
    .concat(pngChunk_('IHDR', ihdr))
    .concat(pngChunk_('IDAT', zlibStore_(raw)))
    .concat(pngChunk_('IEND', []));
}

/** Returns an Apps Script PNG blob of a QR code for the text. */
function qrPngBlob(text, name) {
  var bytes = qrPngBytes(qrMatrix(text), 8).map(function (b) { return b > 127 ? b - 256 : b; });
  return Utilities.newBlob(bytes, 'image/png', name || 'qr.png');
}

function u32_(v) { return [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255]; }

function pngChunk_(type, data) {
  var typeBytes = type.split('').map(function (ch) { return ch.charCodeAt(0); });
  var crcInput = typeBytes.concat(data);
  return [].concat(u32_(data.length), typeBytes, data, u32_(crc32_(crcInput)));
}

var CRC_TABLE_ = null;
function crc32_(bytes) {
  if (!CRC_TABLE_) {
    CRC_TABLE_ = [];
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      CRC_TABLE_.push(c >>> 0);
    }
  }
  var crc = 0xFFFFFFFF;
  for (var i = 0; i < bytes.length; i++) crc = CRC_TABLE_[(crc ^ bytes[i]) & 255] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

/** zlib wrapper using uncompressed ("stored") blocks. */
function zlibStore_(data) {
  var out = [0x78, 0x01];
  for (var i = 0; i < data.length || i === 0; i += 65535) {
    var chunk = data.slice(i, i + 65535);
    var last = i + 65535 >= data.length ? 1 : 0;
    var len = chunk.length;
    out.push(last, len & 255, (len >>> 8) & 255, ~len & 255, (~len >>> 8) & 255);
    for (var j = 0; j < chunk.length; j++) out.push(chunk[j]);
    if (last) break;
  }
  var a = 1, b = 0;
  for (i = 0; i < data.length; i++) { a = (a + data[i]) % 65521; b = (b + a) % 65521; }
  return out.concat(u32_(((b << 16) | a) >>> 0));
}

/* ======================================================================
 * BUILT-IN LOGO: setup() saves this to Drive as "Tsavone logo for certificates.png"
 * ====================================================================== */

var TSAVONE_LOGO_PNG_BASE64 = [
  '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'
].join('');
