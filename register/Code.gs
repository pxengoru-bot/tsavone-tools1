/**
 * Tsavone Consulting – Training Register
 *
 * Paste this file into Extensions > Apps Script of the Google Sheet
 * "Tsavone Training Register", together with Qr.gs and Verify.html.
 * See INSTALL.md for step-by-step instructions.
 *
 * Uses only built-in Google services: SpreadsheetApp, DriveApp, DocumentApp,
 * GmailApp (for drafts), HtmlService and Utilities.
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
  ['logo_drive_file_id', '', 'Drive file ID of the logo (PNG or JPG). See INSTALL.md step 4'],
  ['signatory_title', 'Managing Director', 'Title printed under the signature line (no personal name)'],
  ['sender_email_name', 'Tsavone Consulting Training', 'Sender name shown on certificate emails'],
  ['verify_url', '', 'Web app link of the verification page. See INSTALL.md step 5'],
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
  onOpen();
  ss.setActiveSheet(sh);
  ss.toast('Setup complete. Next: fill in the Settings tab (see INSTALL.md).', 'Tsavone', 8);
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

  var msg = made + ' certificate(s) saved to Google Drive > ' + root.getName() + '.';
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

/* ============================== VERIFICATION WEB APP ============================== */

function doGet(e) {
  var id = String((e && e.parameter && e.parameter.id) || '').trim().slice(0, 60);
  var t = HtmlService.createTemplateFromFile('Verify');
  t.query = id;
  t.result = id ? lookupCertificate_(id) : null;
  t.pageUrl = ScriptApp.getService().getUrl() || '';
  t.logo = verifyLogoDataUri_();
  return t.evaluate()
    .setTitle('Verify a Tsavone certificate')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
