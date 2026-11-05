/**
 * Tsavone Consulting – Tender Tracker
 *
 * Paste this file into Extensions > Apps Script of the private Google Sheet
 * "Tsavone Tender Tracker", together with Sidebar.html. See INSTALL.md.
 *
 * Uses only built-in Google services: SpreadsheetApp, DriveApp, DocumentApp, MailApp and
 * HtmlService (plus ScriptApp for the daily reminder schedule and Utilities for dates).
 * Nothing is made public: folders and documents stay private to your account.
 */

/* ============================== SETTINGS YOU MAY CHANGE ============================== */

var REMINDER_EMAIL = 'p.engoru@tsavone.com';   // who receives the daily reminder
var REMINDER_HOUR = 7;                          // 07:00 (Google runs it between 07:00 and 08:00)
var TIME_ZONE = 'Africa/Kampala';
var BIDS_FOLDER = 'Tsavone Bids';               // Google Drive folder (in My Drive) for bid folders
var REMINDER_DAYS = 7;                          // remind about tenders due within this many days
var DOC_WARNING_DAYS = 30;                      // warn about documents expiring within this many days

/* ============================== CONSTANTS ============================== */

var TENDERS = 'Tenders';
var DOCUMENTS = 'Documents';
var DASHBOARD = 'Dashboard';

var COLUMNS = [
  'ref_id', 'organisation', 'title', 'notice_type', 'category', 'source_link', 'date_found',
  'deadline', 'submission_method', 'status', 'lead_person', 'documents_required', 'documents_ready',
  'bid_value_ugx', 'go_no_go', 'result', 'notes', 'submitted_date'
];

var LISTS = {
  notice_type: ['Prequalification', 'RFQ', 'RFP', 'EOI', 'Tender', 'Framework', 'Call for consultants'],
  category: ['Training', 'Consultancy', 'Research and evaluation', 'Goods supply', 'Services', 'Works'],
  submission_method: ['Email', 'Hard copy', 'Online portal', 'Other'],
  status: ['Watching', 'Preparing', 'Submitted', 'Shortlisted', 'Won', 'Lost', 'Not pursued', 'Cancelled'],
  go_no_go: ['Go', 'No go', 'Undecided'],
  result: ['Pending', 'Won', 'Lost', 'No response']
};

// Statuses that no longer need deadline warnings (as specified)
var CLOSED_FOR_COLOURS = ['Submitted', 'Won', 'Lost', 'Not pursued', 'Cancelled'];
// Statuses that no longer need a reminder (Shortlisted is also past submission)
var CLOSED_FOR_REMINDER = ['Submitted', 'Shortlisted', 'Won', 'Lost', 'Not pursued', 'Cancelled'];
var PIPELINE = ['Preparing', 'Submitted', 'Shortlisted'];

var DOC_COLUMNS = ['document', 'status', 'issue_date', 'expiry_date', 'drive_link', 'notes'];
var DOC_STATUSES = ['Have', 'Missing', 'Expired'];
var STANDING_DOCUMENTS = [
  'Certificate of Incorporation',
  'Memorandum and Articles of Association',
  'Company Form 18 (registered office) and Form 20 (directors)',
  'URSB annual return confirmation',
  'URA TIN certificate',
  'URA tax clearance certificate',
  'Trading licence',
  'NSSF registration and compliance certificate',
  'VAT registration certificate (if registered)',
  'Bank reference letter / bank details letter',
  'Audited financial statements (last 1-3 years)',
  'Company profile',
  'Capability statement',
  'CVs of key experts',
  'Past contracts, LPOs and completion certificates',
  'Recommendation letters',
  'PPDA supplier registration (if applicable)',
  'Signed safeguarding, anti-fraud and data protection policies',
  'Power of attorney for bid signing',
  'Bid declaration and anti-bribery declaration templates'
];

// Words people use in tender notices -> the standing document they mean
var DOC_ALIASES = [
  [/incorporat/, 'Certificate of Incorporation'],
  [/memorandum|articles of assoc|\bmemart/, 'Memorandum and Articles of Association'],
  [/form\s*18|form\s*20|registered office|particulars of directors/, 'Company Form 18 (registered office) and Form 20 (directors)'],
  [/annual return/, 'URSB annual return confirmation'],
  [/\btin\b|tax identification/, 'URA TIN certificate'],
  [/tax clearance|tcc\b/, 'URA tax clearance certificate'],
  [/trading licen[cs]e/, 'Trading licence'],
  [/nssf|social security/, 'NSSF registration and compliance certificate'],
  [/\bvat\b/, 'VAT registration certificate (if registered)'],
  [/bank/, 'Bank reference letter / bank details letter'],
  [/audited|financial statement|accounts/, 'Audited financial statements (last 1-3 years)'],
  [/company profile|corporate profile|firm profile/, 'Company profile'],
  [/capability|capacity statement/, 'Capability statement'],
  [/\bcvs?\b|curricul|key expert|key staff|personnel/, 'CVs of key experts'],
  [/past contract|\blpos?\b|completion cert|similar assignment|experience|previous contract/, 'Past contracts, LPOs and completion certificates'],
  [/recommendation|reference letter|testimonial/, 'Recommendation letters'],
  [/ppda|supplier regist|provider regist/, 'PPDA supplier registration (if applicable)'],
  [/safeguard|anti.?fraud|data protection|psea|polic(y|ies)/, 'Signed safeguarding, anti-fraud and data protection policies'],
  [/power of attorney|authori[sz]ation to sign|bid signing/, 'Power of attorney for bid signing'],
  [/declaration|anti.?bribery|integrity pledge|code of ethical/, 'Bid declaration and anti-bribery declaration templates']
];

var STANDARD_STEPS = [
  ['Clarification deadline', 'Note the last date for questions; send any questions in writing and keep the replies.'],
  ['Go / no-go decision', 'Confirm eligibility, capacity, timeline and value; record the decision in the tracker.'],
  ['Pricing', 'Build the financial proposal or price schedule; check taxes, currency and validity period.'],
  ['Technical proposal', 'Methodology, work plan, team and CVs, experience; follow the evaluation criteria headings.'],
  ['Signatures and stamps', 'Sign and stamp every form that needs it; check the power of attorney covers the signatory.'],
  ['Compliance check', 'Tick every required document against the notice; check page limits, format, number of copies and bid security.'],
  ['Submission and proof of submission', 'Submit by the method required; keep the email receipt, portal confirmation or stamped acknowledgement.']
];

var NAVY = '#193591';
var LIGHT_BLUE = '#EFFCFF';
var TEST_NOTE = 'TEST ROW - delete after testing';

/* ============================== MENU AND TRIGGERS ============================== */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tsavone Bids')
    .addItem('Add new tender', 'showAddTenderSidebar')
    .addItem('Create bid folder and checklist (selected tender)', 'menuCreateBidFolder')
    .addItem('Mark as submitted (selected tender)', 'menuMarkSubmitted')
    .addItem('Refresh dashboard', 'refreshDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Daily reminder')
      .addItem('Turn daily reminder on', 'installReminder')
      .addItem('Turn daily reminder off', 'removeReminder')
      .addItem('Send reminder now (test)', 'menuSendReminderNow'))
    .addSubMenu(ui.createMenu('Testing')
      .addItem('Add sample tenders', 'menuAddSampleTenders')
      .addItem('Remove sample tenders', 'menuRemoveSampleTenders'))
    .addItem('Run setup again', 'setup')
    .addToUi();
}

/** Simple trigger: gives new rows a ref_id (and sensible defaults) as soon as you type an organisation or title. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== TENDERS || e.range.getLastRow() < 2) return;
    fillMissingIds_(sh, Math.max(2, e.range.getRow()), e.range.getLastRow());
  } catch (err) { /* never interrupt typing */ }
}

/* ============================== SETUP ============================== */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSpreadsheetTimeZone() !== TIME_ZONE) ss.setSpreadsheetTimeZone(TIME_ZONE);

  var sh = ss.getSheetByName(TENDERS) || ss.getSheets()[0];
  sh.setName(TENDERS);
  ensureHeaders_(sh, COLUMNS);
  styleHeader_(sh);
  var c = headerMap_(sh);
  var rows = Math.max(sh.getMaxRows() - 1, 1);

  Object.keys(LISTS).forEach(function (col) {
    sh.getRange(2, c[col] + 1, rows, 1).setDataValidation(listRule_(LISTS[col]));
  });
  ['date_found', 'deadline', 'submitted_date'].forEach(function (col) {
    sh.getRange(2, c[col] + 1, rows, 1).setNumberFormat('d mmm yyyy');
  });
  sh.getRange(2, c.bid_value_ugx + 1, rows, 1).setNumberFormat('"UGX" #,##0');
  sh.getRange(2, c.ref_id + 1, rows, 1).setNumberFormat('@');
  var widths = { ref_id: 120, organisation: 200, title: 280, documents_required: 280, notes: 280, source_link: 160, deadline: 110 };
  Object.keys(widths).forEach(function (k) { sh.setColumnWidth(c[k] + 1, widths[k]); });
  sh.getRange(2, c.documents_required + 1, rows, 1).setWrap(true);
  tenderColourRules_(sh, c);

  setupDocuments_(ss);
  fillMissingIds_(sh, 2, Math.max(sh.getLastRow(), 2));
  refreshDashboard(true);
  onOpen();
  ss.setActiveSheet(sh);
  ss.toast('Setup complete. Next: turn on the daily reminder (Tsavone Bids > Daily reminder).', 'Tsavone Bids', 10);
}

function ensureHeaders_(sh, cols) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  if (existing.join('') === '') {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    return;
  }
  var lower = existing.map(function (h) { return h.toLowerCase(); });
  var missing = cols.filter(function (col) { return lower.indexOf(col) === -1; });
  if (missing.length) sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
}

function styleHeader_(sh) {
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, sh.getLastColumn())
    .setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 34);
}

function listRule_(list) {
  return SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
}

/** Deadline colours: passed grey, 3 days red, 7 days orange, 14 days yellow (open tenders only). */
function tenderColourRules_(sh, c) {
  var d = '$' + colLetter_(c.deadline + 1) + '2';
  var s = '$' + colLetter_(c.status + 1) + '2';
  var open = 'NOT(OR(' + CLOSED_FOR_COLOURS.map(function (x) { return s + '="' + x + '"'; }).join(',') + '))';
  var range = sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), sh.getLastColumn());
  var rule = function (cond, colour) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(' + d + '),' + open + ',' + cond + ')')
      .setBackground(colour).setRanges([range]).build();
  };
  sh.setConditionalFormatRules(tenderColourFormulas_(d).map(function (r) { return rule(r[0], r[1]); }));
}

/** The four deadline conditions, in priority order. */
function tenderColourFormulas_(d) {
  return [
    [d + '<TODAY()', '#d9d9d9'],
    [d + '-TODAY()<=3', '#f8c9c4'],
    [d + '-TODAY()<=7', '#ffd8b0'],
    [d + '-TODAY()<=14', '#fff2a8']
  ];
}

function setupDocuments_(ss) {
  var sh = ss.getSheetByName(DOCUMENTS) || ss.insertSheet(DOCUMENTS);
  ensureHeaders_(sh, DOC_COLUMNS);
  styleHeader_(sh);
  var c = headerMap_(sh);
  var have = sh.getLastRow() > 1 ? sh.getRange(2, c.document + 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim().toLowerCase(); }) : [];
  var add = STANDING_DOCUMENTS.filter(function (d) { return have.indexOf(d.toLowerCase()) === -1; });
  if (add.length) {
    var width = sh.getLastColumn();
    var values = add.map(function (d) { var row = new Array(width).fill(''); row[c.document] = d; return row; });
    sh.getRange(sh.getLastRow() + 1, 1, values.length, width).setValues(values);
  }
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, c.status + 1, rows, 1).setDataValidation(listRule_(DOC_STATUSES));
  sh.getRange(2, c.issue_date + 1, rows, 1).setNumberFormat('d mmm yyyy');
  sh.getRange(2, c.expiry_date + 1, rows, 1).setNumberFormat('d mmm yyyy');
  sh.setColumnWidth(c.document + 1, 380);
  sh.setColumnWidth(c.drive_link + 1, 260);
  sh.setColumnWidth(c.notes + 1, 260);

  var e = '$' + colLetter_(c.expiry_date + 1) + '2';
  var st = '$' + colLetter_(c.status + 1) + '2';
  var range = sh.getRange(2, 1, rows, sh.getLastColumn());
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR(' + st + '="Expired",AND(ISNUMBER(' + e + '),' + e + '<TODAY()))')
      .setBackground('#f8c9c4').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(' + e + '),' + e + '-TODAY()<=' + DOC_WARNING_DAYS + ')')
      .setBackground('#ffd8b0').setRanges([range]).build()
  ]);
}

/* ============================== HELPERS ============================== */

function headerMap_(sh) {
  var head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  var map = {};
  head.forEach(function (h, i) { var k = String(h).trim().toLowerCase(); if (k && map[k] === undefined) map[k] = i; });
  return map;
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function tendersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TENDERS);
  if (!sh) throw new Error('No "Tenders" tab. Run Tsavone Bids > Run setup again.');
  var c = headerMap_(sh);
  var missing = COLUMNS.filter(function (k) { return c[k] === undefined; });
  if (missing.length) throw new Error('Missing columns: ' + missing.join(', ') + '. Run Tsavone Bids > Run setup again.');
  return { ss: ss, sh: sh, c: c };
}

function readRows_(sh, c) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues().map(function (v, i) {
    var o = { rowNumber: i + 2 };
    Object.keys(c).forEach(function (k) { o[k] = v[c[k]]; });
    return o;
  }).filter(function (o) { return String(o.organisation || '').trim() || String(o.title || '').trim() || String(o.ref_id || '').trim(); });
}

function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }
function blank_(v) { return v === null || v === undefined || String(v).trim() === ''; }

/** "yyyy-MM-dd" of a moment, as a Kampala calendar date. */
function dayKey_(d) { return Utilities.formatDate(d, TIME_ZONE, 'yyyy-MM-dd'); }

/** Midnight in Kampala (UTC+3, no daylight saving) for a calendar date. */
function kampalaDate_(y, m, d) { return new Date(Date.UTC(y, m - 1, d) - 3 * 3600 * 1000); }

function todayKampala_() {
  var k = dayKey_(new Date()).split('-').map(Number);
  return kampalaDate_(k[0], k[1], k[2]);
}

/** Whole days from today (Kampala) to the date: 0 = today, negative = passed. */
function daysUntil_(date, today) {
  var a = dayKey_(date).split('-').map(Number), b = dayKey_(today || new Date()).split('-').map(Number);
  return Math.round((Date.UTC(a[0], a[1] - 1, a[2]) - Date.UTC(b[0], b[1] - 1, b[2])) / 86400000);
}

function fmtDate_(d) { return isDate_(d) ? Utilities.formatDate(d, TIME_ZONE, 'd MMM yyyy') : ''; }

/**
 * Reads a date typed by a person: 2026-11-05, 5/11/2026 (day first, as in Uganda), 5 Nov 2026 or 5 November 2026.
 * Returns a Kampala-midnight Date, or null.
 */
function parseDate_(text) {
  var s = String(text || '').trim();
  if (!s) return null;
  var m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) return validDate_(+m[1], +m[2], +m[3]);
  if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/))) return validDate_(+m[3], +m[2], +m[1]);
  if ((m = s.match(/^(\d{1,2})\s+([A-Za-z]{3,})\.?,?\s+(\d{4})$/))) {
    var mon = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].indexOf(m[2].slice(0, 3).toLowerCase());
    if (mon >= 0) return validDate_(+m[3], mon + 1, +m[1]);
  }
  return null;
}

function validDate_(y, mo, d) {
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  var probe = new Date(Date.UTC(y, mo - 1, d));
  if (probe.getUTCMonth() !== mo - 1) return null;
  return kampalaDate_(y, mo, d);
}

/** Next unique ID such as TND-2026-0007. used = object of IDs already taken. */
function nextRefId_(year, used) {
  var prefix = 'TND-' + year + '-';
  var max = 0;
  Object.keys(used).forEach(function (id) {
    if (id.indexOf(prefix) === 0) { var n = parseInt(id.slice(prefix.length), 10); if (n > max) max = n; }
  });
  var id = prefix + ('000' + (max + 1)).slice(-Math.max(4, String(max + 1).length));
  used[id] = true;
  return id;
}

/** Gives rows between first and last a ref_id (and default date_found, status, go/no-go, result) if missing. */
function fillMissingIds_(sh, first, last) {
  var c = headerMap_(sh);
  if (c.ref_id === undefined) return 0;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  last = Math.min(last, lastRow);
  if (last < first) return 0;
  var all = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var used = {};
  all.forEach(function (r) { var id = String(r[c.ref_id]).trim(); if (id) used[id] = true; });
  var made = 0;
  for (var row = first; row <= last; row++) {
    var v = all[row - 2];
    if (!String(v[c.ref_id]).trim() && (String(v[c.organisation]).trim() || String(v[c.title]).trim())) {
      // Write only the cells being filled, so anything else in the row (including formulas) is untouched
      var found = isDate_(v[c.date_found]) ? v[c.date_found] : todayKampala_();
      var id = nextRefId_(Utilities.formatDate(found, TIME_ZONE, 'yyyy'), used);
      sh.getRange(row, c.ref_id + 1).setValue(id);
      if (blank_(v[c.date_found])) sh.getRange(row, c.date_found + 1).setValue(found);
      if (blank_(v[c.status])) sh.getRange(row, c.status + 1).setValue('Watching');
      if (blank_(v[c.go_no_go])) sh.getRange(row, c.go_no_go + 1).setValue('Undecided');
      if (blank_(v[c.result])) sh.getRange(row, c.result + 1).setValue('Pending');
      made++;
    }
  }
  return made;
}

/** The tender row the person has selected (active cell). */
function selectedTender_(t) {
  var active = t.ss.getActiveSheet();
  if (active.getName() !== TENDERS) throw new Error('Go to the Tenders tab and click any cell in the tender\'s row first.');
  var row = active.getActiveRange().getRow();
  var r = readRows_(t.sh, t.c).filter(function (x) { return x.rowNumber === row; })[0];
  if (!r) throw new Error('Click a cell in a tender\'s row first (not the heading row or an empty row).');
  if (blank_(r.ref_id)) { fillMissingIds_(t.sh, row, row); r = readRows_(t.sh, t.c).filter(function (x) { return x.rowNumber === row; })[0]; }
  return r;
}

function setCell_(t, row, col, value) { t.sh.getRange(row, t.c[col] + 1).setValue(value); }

function alert_(title, msg) { var ui = SpreadsheetApp.getUi(); ui.alert(title, msg, ui.ButtonSet.OK); }

function safeName_(s) { return String(s || '').replace(/[\\\/:*?"<>|#%]+/g, ' ').replace(/\s+/g, ' ').trim(); }

/** First words of a title, for folder names. */
function shortTitle_(title) {
  var words = safeName_(title).split(' ').filter(String);
  var out = '';
  for (var i = 0; i < words.length; i++) {
    var next = out ? out + ' ' + words[i] : words[i];
    if (next.length > 45 || i >= 7) break;
    out = next;
  }
  return out || 'Tender';
}

/** Splits documents_required on commas, semicolons or new lines. */
function splitRequired_(text) {
  return String(text || '').split(/\r?\n|;|,(?![^(]*\))/).map(function (s) { return s.replace(/^[\s\-•*\d.)]+/, '').trim(); }).filter(String);
}

/** Finds the standing document a requirement refers to. docs = [{document, status, drive_link, expiry_date}] */
function matchDocument_(required, docs) {
  var text = String(required).toLowerCase();
  for (var i = 0; i < DOC_ALIASES.length; i++) {
    if (DOC_ALIASES[i][0].test(text)) {
      var name = DOC_ALIASES[i][1].toLowerCase();
      var hit = docs.filter(function (d) { return String(d.document).toLowerCase() === name; })[0];
      if (hit) return hit;
    }
  }
  // Fall back to shared words (for documents you added yourself)
  var stop = { the: 1, and: 1, of: 1, for: 1, a: 1, an: 1, to: 1, copy: 1, copies: 1, valid: 1, certified: 1, original: 1, if: 1, applicable: 1, letter: 1, certificate: 1, company: 1 };
  var words = function (s) { return String(s).toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(function (w) { return w.length > 2 && !stop[w]; }); };
  var want = words(text);
  var best = null, bestScore = 0;
  docs.forEach(function (d) {
    var have = words(d.document);
    var common = want.filter(function (w) { return have.indexOf(w) !== -1; }).length;
    var score = common / Math.max(1, Math.min(want.length, have.length));
    if (common >= 1 && score > bestScore) { bestScore = score; best = d; }
  });
  return bestScore >= 0.5 ? best : null;
}

function readDocuments_(ss) {
  var sh = ss.getSheetByName(DOCUMENTS);
  if (!sh || sh.getLastRow() < 2) return [];
  var c = headerMap_(sh);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().map(function (v) {
    var o = {};
    Object.keys(c).forEach(function (k) { o[k] = v[c[k]]; });
    return o;
  }).filter(function (d) { return String(d.document || '').trim(); });
}

/** Effective document status: Expired if the expiry date has passed. */
function docState_(d, today) {
  var st = String(d.status || '').trim() || 'Not checked';
  if (isDate_(d.expiry_date) && daysUntil_(d.expiry_date, today) < 0) return 'Expired';
  return st;
}

/* ============================== ADD NEW TENDER (SIDEBAR) ============================== */

function showAddTenderSidebar() {
  var t = HtmlService.createTemplateFromFile('Sidebar');
  t.lists = LISTS;
  t.today = dayKey_(new Date());
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('Add new tender'));
}

/** Called from the sidebar. Returns the new ref_id. */
function addTender(form) {
  var t = tendersSheet_();
  form = form || {};
  var org = String(form.organisation || '').trim(), title = String(form.title || '').trim();
  if (!org || !title) throw new Error('Please fill in at least the organisation and the title.');
  var deadline = form.deadline ? parseDate_(form.deadline) : null;
  if (form.deadline && !deadline) throw new Error('The deadline date was not understood.');
  var found = parseDate_(form.date_found) || todayKampala_();
  var value = String(form.bid_value_ugx || '').replace(/[^\d.]/g, '');
  var used = {};
  readRows_(t.sh, t.c).forEach(function (r) { if (String(r.ref_id).trim()) used[String(r.ref_id).trim()] = true; });
  var id = nextRefId_(Utilities.formatDate(found, TIME_ZONE, 'yyyy'), used);
  var row = new Array(t.sh.getLastColumn()).fill('');
  var set = function (k, v) { if (t.c[k] !== undefined) row[t.c[k]] = v; };
  set('ref_id', id); set('organisation', org); set('title', title);
  set('notice_type', pick_(form.notice_type, LISTS.notice_type));
  set('category', pick_(form.category, LISTS.category));
  set('source_link', String(form.source_link || '').trim());
  set('date_found', found);
  set('deadline', deadline || '');
  set('submission_method', pick_(form.submission_method, LISTS.submission_method));
  set('status', pick_(form.status, LISTS.status) || 'Watching');
  set('lead_person', String(form.lead_person || '').trim());
  set('documents_required', String(form.documents_required || '').trim());
  set('bid_value_ugx', value ? Number(value) : '');
  set('go_no_go', pick_(form.go_no_go, LISTS.go_no_go) || 'Undecided');
  set('result', 'Pending');
  set('notes', String(form.notes || '').trim());
  var target = t.sh.getLastRow() + 1;
  t.sh.getRange(target, 1, 1, row.length).setValues([row]);
  return id;
}

function pick_(value, list) { return list.indexOf(value) !== -1 ? value : ''; }

/* ============================== BID FOLDER AND CHECKLIST ============================== */

function menuCreateBidFolder() {
  var ui = SpreadsheetApp.getUi();
  try {
    var t = tendersSheet_();
    var r = selectedTender_(t);
    var res = createBidFolder_(t, r);
    if (!res) return;
    ui.alert('Bid folder ready',
      'Folder: Tsavone Bids / ' + res.folderName + '\n\nChecklist: ' + res.matched + ' of ' + res.total + ' required document(s) found in your Documents tab' +
      (res.ready ? ' (' + res.ready + ' marked "Have")' : '') + '.\n\nThe folder link has been added to the notes column.',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Create bid folder', e.message, ui.ButtonSet.OK);
  }
}

function createBidFolder_(t, r) {
  var ui = SpreadsheetApp.getUi();
  var folderName = safeName_(r.ref_id + ' - ' + r.organisation + ' - ' + shortTitle_(r.title));
  var root = childFolder_(DriveApp.getRootFolder(), BIDS_FOLDER);
  var folder = childFolder_(root, folderName);
  var docName = 'Bid checklist - ' + r.ref_id;
  var old = folder.getFilesByName(docName);
  if (old.hasNext()) {
    var choice = ui.alert('Checklist already exists', 'This tender already has a checklist. Replace it with a fresh one?\n(The old one goes to the Drive Bin.)', ui.ButtonSet.YES_NO);
    if (choice !== ui.Button.YES) return null;
    while (old.hasNext()) old.next().setTrashed(true);
  }

  var today = todayKampala_();
  var docs = readDocuments_(t.ss);
  var required = splitRequired_(r.documents_required);
  var lines = required.map(function (req) {
    var d = matchDocument_(req, docs);
    return { required: req, doc: d, state: d ? docState_(d, today) : 'Not in Documents tab', link: d ? String(d.drive_link || '').trim() : '' };
  });

  var doc = DocumentApp.create(docName);
  var body = doc.getBody();
  body.setMarginTop(42).setMarginBottom(42).setMarginLeft(50).setMarginRight(50);
  var title = body.getChild(0).asParagraph();
  title.setText('Bid checklist: ' + r.ref_id).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.editAsText().setForegroundColor(NAVY);
  body.appendParagraph(String(r.title)).setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var details = [
    ['Organisation', String(r.organisation || '')],
    ['Notice type', String(r.notice_type || '')],
    ['Category', String(r.category || '')],
    ['Deadline', isDate_(r.deadline) ? fmtDate_(r.deadline) + daysText_(r.deadline, today) : 'Not set'],
    ['Submission method', String(r.submission_method || '')],
    ['Lead person', String(r.lead_person || '')],
    ['Source / notice', String(r.source_link || '')],
    ['Bid folder', folder.getUrl()]
  ];
  var dt = body.appendTable(details);
  styleTable_(dt, false);
  linkCells_(dt, 1, [6, 7]);

  body.appendParagraph('Required documents').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (lines.length) {
    var rows = [['Done', 'Required by the notice', 'Our document', 'Status', 'Drive link']];
    lines.forEach(function (l) {
      rows.push(['☐', l.required, l.doc ? String(l.doc.document) : '–', l.state, l.link || (l.doc ? 'Add link in Documents tab' : 'Prepare for this bid')]);
    });
    var rt = body.appendTable(rows);
    styleTable_(rt, true);
    for (var i = 1; i < rows.length; i++) {
      var st = rows[i][3];
      var colour = st === 'Have' ? '#e6f4ea' : st === 'Expired' || st === 'Missing' ? '#fde2df' : '#fff4d6';
      rt.getCell(i, 3).setBackgroundColor(colour);
      if (lines[i - 1].link) linkCell_(rt.getCell(i, 4), lines[i - 1].link);
    }
  } else {
    body.appendParagraph('No documents listed yet. Copy the list from the notice into the tracker\'s documents_required column (one per line or separated by commas), then create the checklist again.');
  }

  body.appendParagraph('Bid steps').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var steps = [['Done', 'Step', 'What to check', 'Who', 'By when']];
  STANDARD_STEPS.forEach(function (s) { steps.push(['☐', s[0], s[1], '', '']); });
  styleTable_(body.appendTable(steps), true);

  body.appendParagraph('Created ' + Utilities.formatDate(new Date(), TIME_ZONE, 'd MMM yyyy, HH:mm') + ' from the Tsavone Tender Tracker. This document is private to your Google account.')
    .editAsText().setFontSize(8).setForegroundColor('#5a6385');
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(folder);

  // Record the folder link and document readiness in the tracker
  var notes = String(r.notes || '');
  if (notes.indexOf(folder.getUrl()) === -1) {
    setCell_(t, r.rowNumber, 'notes', (notes ? notes + ' | ' : '') + 'Bid folder: ' + folder.getUrl());
  }
  var matched = lines.filter(function (l) { return l.doc; }).length;
  var ready = lines.filter(function (l) { return l.state === 'Have'; }).length;
  if (blank_(r.documents_ready) && lines.length) setCell_(t, r.rowNumber, 'documents_ready', ready + ' of ' + lines.length + ' ready');
  return { folderName: folderName, total: lines.length, matched: matched, ready: ready, folderUrl: folder.getUrl(), docUrl: doc.getUrl() };
}

function daysText_(date, today) {
  var n = daysUntil_(date, today);
  return n < 0 ? ' (passed)' : n === 0 ? ' (today)' : n === 1 ? ' (tomorrow)' : ' (in ' + n + ' days)';
}

function childFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function styleTable_(table, header) {
  table.setBorderColor('#dbe3f0');
  if (!header) {
    for (var r = 0; r < table.getNumRows(); r++) table.getCell(r, 0).setBackgroundColor(LIGHT_BLUE).editAsText().setBold(true);
    return;
  }
  var head = table.getRow(0);
  for (var i = 0; i < head.getNumCells(); i++) {
    head.getCell(i).setBackgroundColor(NAVY).editAsText().setForegroundColor('#ffffff').setBold(true);
  }
}

function linkCells_(table, col, rows) {
  rows.forEach(function (r) {
    var cell = table.getCell(r, col);
    var url = cell.getText();
    if (/^https?:\/\//.test(url)) linkCell_(cell, url);
  });
}

function linkCell_(cell, url) {
  if (!/^https?:\/\//.test(url)) return;
  var text = cell.editAsText();
  if (text.getText().length) text.setLinkUrl(url);
}

/* ============================== MARK AS SUBMITTED ============================== */

function menuMarkSubmitted() {
  var ui = SpreadsheetApp.getUi();
  try {
    var t = tendersSheet_();
    var r = selectedTender_(t);
    var resp = ui.prompt('Mark as submitted',
      r.ref_id + ': ' + r.title + '\n\nSubmission date? Leave empty for today, or type a date such as ' + Utilities.formatDate(new Date(), TIME_ZONE, 'd MMM yyyy') + '.',
      ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    var text = resp.getResponseText().trim();
    var date = text ? parseDate_(text) : todayKampala_();
    if (!date) { ui.alert('Mark as submitted', 'That date was not understood. Try again with a date such as 5 Nov 2026.', ui.ButtonSet.OK); return; }
    setCell_(t, r.rowNumber, 'status', 'Submitted');
    setCell_(t, r.rowNumber, 'submitted_date', date);
    if (blank_(r.result)) setCell_(t, r.rowNumber, 'result', 'Pending');
    if (blank_(r.go_no_go) || String(r.go_no_go) === 'Undecided') setCell_(t, r.rowNumber, 'go_no_go', 'Go');
    ui.alert('Mark as submitted', r.ref_id + ' is now marked Submitted on ' + fmtDate_(date) + '.\n\nRemember to save the proof of submission in its bid folder.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Mark as submitted', e.message, ui.ButtonSet.OK);
  }
}

/* ============================== DASHBOARD ============================== */

/** Rebuilds the Dashboard tab. The numbers are live formulas, so they stay up to date between refreshes. */
function refreshDashboard(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var t = tendersSheet_();
  var dsh = ss.getSheetByName(DOCUMENTS);
  var dc = dsh ? headerMap_(dsh) : null;
  var sh = ss.getSheetByName(DASHBOARD) || ss.insertSheet(DASHBOARD);
  sh.clear();
  sh.setConditionalFormatRules([]);
  var f = dashboardFormulas_(t.c, dc);

  // Column A-B: key figures, status and category breakdowns
  var put = function (row, col, value) { sh.getRange(row, col).setValue(value); };
  var putF = function (row, col, formula) { sh.getRange(row, col).setFormula(formula); };
  put(1, 1, 'Tsavone bid and tender dashboard');
  put(2, 1, 'Refreshed ' + Utilities.formatDate(new Date(), TIME_ZONE, 'd MMM yyyy, HH:mm') + ' (figures update automatically)');

  var row = 4;
  put(row, 1, 'Key figures'); put(row, 2, ''); heading_(sh, row, 1, 2); row++;
  f.kpis.forEach(function (k) { put(row, 1, k[0]); putF(row, 2, k[1]); if (k[2]) sh.getRange(row, 2).setNumberFormat(k[2]); row++; });

  row++;
  put(row, 1, 'Tenders by status'); put(row, 2, 'Number'); heading_(sh, row, 1, 2); row++;
  f.byStatus.forEach(function (k) { put(row, 1, k[0]); putF(row, 2, k[1]); row++; });

  row++;
  put(row, 1, 'By category'); put(row, 2, 'Number'); put(row, 3, 'Value (UGX)'); heading_(sh, row, 1, 3); row++;
  f.byCategory.forEach(function (k) { put(row, 1, k[0]); putF(row, 2, k[1]); putF(row, 3, k[2]); sh.getRange(row, 3).setNumberFormat('#,##0'); row++; });

  // Columns E-H: due in 7 days; J-M: due in 30 days; O-R: documents needing attention
  var block = function (col, title, heads, formula, dateCols) {
    put(1, col, title); heading_(sh, 1, col, heads.length);
    heads.forEach(function (h, i) { put(2, col + i, h); });
    sh.getRange(2, col, 1, heads.length).setFontWeight('bold').setBackground(LIGHT_BLUE);
    putF(3, col, formula);
    (dateCols || []).forEach(function (i) { sh.getRange(3, col + i, 200, 1).setNumberFormat('d mmm yyyy'); });
  };
  block(5, 'Due in the next 7 days (not submitted)', ['Ref', 'Organisation', 'Title', 'Deadline'], f.due7, [3]);
  block(10, 'Due in the next 30 days (not submitted)', ['Ref', 'Organisation', 'Title', 'Deadline'], f.due30, [3]);
  if (f.docs) block(15, 'Documents missing, expired or expiring within ' + DOC_WARNING_DAYS + ' days', ['Document', 'Status', 'Expiry date'], f.docs, [2]);

  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor(NAVY);
  sh.getRange(2, 1).setFontColor('#5a6385');
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 140); sh.setColumnWidth(4, 30);
  [5, 10].forEach(function (c0) { sh.setColumnWidth(c0, 110); sh.setColumnWidth(c0 + 1, 180); sh.setColumnWidth(c0 + 2, 260); sh.setColumnWidth(c0 + 3, 100); sh.setColumnWidth(c0 + 4, 30); });
  sh.setColumnWidth(15, 340); sh.setColumnWidth(16, 110); sh.setColumnWidth(17, 110);
  sh.setFrozenRows(2);
  if (silent !== true) ss.setActiveSheet(sh);
}

function heading_(sh, row, col, width) {
  sh.getRange(row, col, 1, width).setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold');
}

/** All dashboard formulas, built from the actual column positions. */
function dashboardFormulas_(c, dc) {
  var col = function (k) { var L = colLetter_(c[k] + 1); return "'" + TENDERS + "'!" + L + '2:' + L; };
  var S = col('status'), R = col('result'), V = col('bid_value_ugx'), D = col('deadline'), SD = col('submitted_date');
  var REF = col('ref_id'), ORG = col('organisation'), TI = col('title'), CAT = col('category');
  var list = function (arr) { return '{' + arr.map(function (x) { return '"' + x + '"'; }).join(',') + '}'; };
  var won = '((' + S + '="Won")+(' + R + '="Won")>0)';
  var lost = '((' + S + '="Lost")+(' + R + '="Lost")>0)';
  var openForDue = 'ISNA(MATCH(' + S + ',' + list(CLOSED_FOR_REMINDER) + ',0))';
  var due = function (days) {
    return '=IFERROR(SORT(FILTER({' + REF + ',' + ORG + ',' + TI + ',' + D + '},ISNUMBER(' + D + '),' + D + '>=TODAY(),' + D + '<=TODAY()+' + days + ',' + openForDue + '),4,TRUE),"None")';
  };
  var out = {
    kpis: [
      ['Tenders tracked', '=COUNTA(' + REF + ')'],
      ['Open (Watching or Preparing)', '=COUNTIF(' + S + ',"Watching")+COUNTIF(' + S + ',"Preparing")'],
      ['Due in the next 7 days', '=IFERROR(ROWS(FILTER(' + D + ',ISNUMBER(' + D + '),' + D + '>=TODAY(),' + D + '<=TODAY()+7,' + openForDue + ')),0)'],
      ['Submitted this year', '=ARRAYFORMULA(SUMPRODUCT(ISNUMBER(MATCH(' + S + ',' + list(['Submitted', 'Shortlisted', 'Won', 'Lost']) + ',0))*(YEAR(IF(ISNUMBER(' + SD + '),' + SD + ',IF(ISNUMBER(' + D + '),' + D + ',0)))=YEAR(TODAY()))))'],
      ['Won', '=SUMPRODUCT(--' + won + ')'],
      ['Lost', '=SUMPRODUCT(--' + lost + ')'],
      ['Win rate (Won ÷ (Won + Lost))', '=IFERROR(SUMPRODUCT(--' + won + ')/(SUMPRODUCT(--' + won + ')+SUMPRODUCT(--' + lost + ')),"–")', '0%'],
      ['Total value won (UGX)', '=ARRAYFORMULA(SUMPRODUCT(--' + won + ',IFERROR(--' + V + ',0)))', '#,##0'],
      ['Value in pipeline: Preparing, Submitted, Shortlisted (UGX)', '=' + PIPELINE.map(function (p) { return 'SUMIFS(' + V + ',' + S + ',"' + p + '")'; }).join('+'), '#,##0']
    ],
    byStatus: LISTS.status.map(function (s) { return [s, '=COUNTIF(' + S + ',"' + s + '")']; }),
    byCategory: LISTS.category.map(function (k) { return [k, '=COUNTIF(' + CAT + ',"' + k + '")', '=SUMIFS(' + V + ',' + CAT + ',"' + k + '")']; }),
    due7: due(7),
    due30: due(30),
    docs: null
  };
  if (dc && dc.document !== undefined) {
    var dcol = function (k) { var L = colLetter_(dc[k] + 1); return "'" + DOCUMENTS + "'!" + L + '2:' + L; };
    var DN = dcol('document'), DS = dcol('status'), DE = dcol('expiry_date');
    out.docs = '=IFERROR(FILTER({' + DN + ',' + DS + ',' + DE + '},LEN(' + DN + '),(' + DS + '="Missing")+(' + DS + '="Expired")+ISNUMBER(' + DE + ')*(' + DE + '<=TODAY()+' + DOC_WARNING_DAYS + ')),"None")';
  }
  return out;
}

/* ============================== DAILY REMINDER ============================== */

/** Turns on the 07:00 Kampala reminder (replaces any earlier one). */
function installReminder() {
  removeReminder(true);
  ScriptApp.newTrigger('sendDailyReminder').timeBased().atHour(REMINDER_HOUR).everyDays(1).inTimezone(TIME_ZONE).create();
  try { alert_('Daily reminder', 'The daily reminder is ON. Each morning between 07:00 and 08:00 (Kampala time) an email goes to ' + REMINDER_EMAIL + ', but only when something needs attention.'); } catch (e) { /* run from the editor */ }
}

/** Turns the reminder off. */
function removeReminder(silent) {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'sendDailyReminder') { ScriptApp.deleteTrigger(tr); n++; }
  });
  if (silent !== true) {
    try { alert_('Daily reminder', n ? 'The daily reminder is OFF.' : 'The daily reminder was already off.'); } catch (e) { /* run from the editor */ }
  }
}

function menuSendReminderNow() {
  var res = sendDailyReminder();
  alert_('Send reminder now', res.sent
    ? 'Reminder sent to ' + REMINDER_EMAIL + ' (' + res.tenders + ' tender(s), ' + res.docs + ' document(s)).'
    : 'Nothing needs attention today, so no email was sent.');
}

/** Builds the reminder. Returns {tenders:[...], docs:[...]} of items needing attention. */
function reminderItems_(tenderRows, docs, today) {
  var tenders = tenderRows.filter(function (r) {
    if (!isDate_(r.deadline)) return false;
    if (CLOSED_FOR_REMINDER.indexOf(String(r.status).trim()) !== -1) return false;
    var n = daysUntil_(r.deadline, today);
    return n >= 0 && n <= REMINDER_DAYS;
  }).sort(function (a, b) { return a.deadline - b.deadline; });
  var docItems = docs.filter(function (d) {
    if (String(d.status).trim() === 'Expired') return true;
    return isDate_(d.expiry_date) && daysUntil_(d.expiry_date, today) <= DOC_WARNING_DAYS;
  });
  return { tenders: tenders, docs: docItems };
}

/** Runs every morning. Sends one short email, only if something needs attention. */
function sendDailyReminder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var t = tendersSheet_();
  var today = new Date();
  var items = reminderItems_(readRows_(t.sh, t.c), readDocuments_(ss), today);
  if (!items.tenders.length && !items.docs.length) return { sent: false, tenders: 0, docs: 0 };

  var url = ss.getUrl();
  var lines = ['Good morning,', '', 'The Tsavone Tender Tracker has items that need attention today.', ''];
  var html = ['<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#2c3557">', '<p>Good morning,</p><p>The Tsavone Tender Tracker has items that need attention today.</p>'];
  var esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };

  if (items.tenders.length) {
    lines.push('TENDERS DUE WITHIN ' + REMINDER_DAYS + ' DAYS (not yet submitted)');
    html.push('<p style="margin:16px 0 6px;font-weight:bold;color:#193591">Tenders due within ' + REMINDER_DAYS + ' days (not yet submitted)</p><ul style="margin:0;padding-left:20px">');
    items.tenders.forEach(function (r) {
      var when = fmtDate_(r.deadline) + daysText_(r.deadline, today);
      var who = String(r.lead_person || '').trim();
      lines.push('- ' + when + ': ' + r.ref_id + ', ' + r.organisation + ', ' + r.title + ' [' + (r.status || 'no status') + (who ? ', lead: ' + who : '') + ']');
      html.push('<li><b>' + esc(when) + '</b>: ' + esc(r.ref_id) + ', ' + esc(r.organisation) + ', ' + esc(r.title) + ' <span style="color:#5a6385">[' + esc(r.status || 'no status') + (who ? ', lead: ' + esc(who) : '') + ']</span></li>');
    });
    lines.push('');
    html.push('</ul>');
  }
  if (items.docs.length) {
    lines.push('COMPANY DOCUMENTS EXPIRED OR EXPIRING WITHIN ' + DOC_WARNING_DAYS + ' DAYS');
    html.push('<p style="margin:16px 0 6px;font-weight:bold;color:#193591">Company documents expired or expiring within ' + DOC_WARNING_DAYS + ' days</p><ul style="margin:0;padding-left:20px">');
    items.docs.forEach(function (d) {
      var when = isDate_(d.expiry_date) ? (daysUntil_(d.expiry_date, today) < 0 ? 'expired ' : 'expires ') + fmtDate_(d.expiry_date) : 'marked Expired';
      lines.push('- ' + d.document + ': ' + when);
      html.push('<li>' + esc(d.document) + ': <b>' + esc(when) + '</b></li>');
    });
    lines.push('');
    html.push('</ul>');
  }
  lines.push('Open the tracker: ' + url, '', 'This is an automatic message from your private Tsavone Tender Tracker.');
  html.push('<p style="margin-top:18px"><a href="' + esc(url) + '" style="color:#193591;font-weight:bold">Open the Tender Tracker</a></p>');
  html.push('<p style="color:#5a6385;font-size:12px">This is an automatic message from your private Tsavone Tender Tracker.</p></div>');

  var subject = 'Tsavone bids: ' + [
    items.tenders.length ? items.tenders.length + ' tender(s) due within ' + REMINDER_DAYS + ' days' : '',
    items.docs.length ? items.docs.length + ' document(s) expiring' : ''
  ].filter(String).join(', ');
  MailApp.sendEmail({ to: REMINDER_EMAIL, subject: subject, body: lines.join('\n'), htmlBody: html.join('') });
  return { sent: true, tenders: items.tenders.length, docs: items.docs.length };
}

/* ============================== TESTING ============================== */

function menuAddSampleTenders() {
  var t = tendersSheet_();
  var rows = readRows_(t.sh, t.c);
  if (rows.some(function (r) { return String(r.notes || '').indexOf(TEST_NOTE) === 0; })) {
    alert_('Add sample tenders', 'Sample tenders are already in the tracker. Use Tsavone Bids > Testing > Remove sample tenders first to start again.');
    return;
  }
  var today = todayKampala_();
  var plus = function (n) { return new Date(today.getTime() + n * 86400000); };
  var samples = [
    { organisation: 'Example District Local Government', title: 'Training of extension workers in climate-smart agriculture', notice_type: 'RFQ', category: 'Training', submission_method: 'Hard copy', deadline: plus(2), status: 'Preparing', lead_person: 'Training lead', bid_value_ugx: 45000000, go_no_go: 'Go',
      documents_required: 'Certificate of incorporation, Valid tax clearance certificate, Trading licence, CVs of trainers, Signed bid form' },
    { organisation: 'Sample Development Partner', title: 'End-line evaluation of a youth employment programme', notice_type: 'RFP', category: 'Research and evaluation', submission_method: 'Email', deadline: plus(6), status: 'Watching', lead_person: 'MEL lead', bid_value_ugx: 120000000, go_no_go: 'Undecided',
      documents_required: 'Company profile\nAudited accounts for the last 2 years\nCVs of key experts\nThree similar assignments with completion certificates\nPower of attorney\nAnti-bribery declaration' },
    { organisation: 'Test Agricultural Exporters Association', title: 'EUDR traceability and due diligence support', notice_type: 'EOI', category: 'Consultancy', submission_method: 'Online portal', deadline: plus(12), status: 'Watching', lead_person: 'Agrifood lead', bid_value_ugx: 18500000, go_no_go: 'Undecided',
      documents_required: 'Capability statement, NSSF compliance certificate, PPDA registration, Bank details letter' }
  ];
  var width = t.sh.getLastColumn();
  var used = {};
  rows.forEach(function (r) { if (String(r.ref_id).trim()) used[String(r.ref_id).trim()] = true; });
  var values = samples.map(function (s) {
    var row = new Array(width).fill('');
    var set = function (k, v) { if (t.c[k] !== undefined) row[t.c[k]] = v; };
    Object.keys(s).forEach(function (k) { set(k, s[k]); });
    set('ref_id', nextRefId_(Utilities.formatDate(today, TIME_ZONE, 'yyyy'), used));
    set('date_found', today); set('source_link', 'https://www.example.com/tender-notice'); set('result', 'Pending'); set('notes', TEST_NOTE);
    return row;
  });
  var first = t.sh.getLastRow() + 1;
  t.sh.getRange(first, 1, values.length, width).setValues(values);
  t.sh.setRowHeights(first, values.length, 42);
  alert_('Add sample tenders', '3 sample tenders added (rows ' + first + '–' + (first + 2) + '), due in 2, 6 and 12 days, so they show red, orange and yellow.\n\nTry: click a sample row, then Tsavone Bids > Create bid folder and checklist. Then Refresh dashboard and Send reminder now.\n\nWhen finished: Tsavone Bids > Testing > Remove sample tenders.');
}

function menuRemoveSampleTenders() {
  var ui = SpreadsheetApp.getUi();
  var t = tendersSheet_();
  var tests = readRows_(t.sh, t.c).filter(function (r) { return String(r.notes || '').indexOf(TEST_NOTE) === 0; });
  if (!tests.length) { ui.alert('Remove sample tenders', 'No sample tenders found.', ui.ButtonSet.OK); return; }
  if (ui.alert('Remove sample tenders', 'Delete the ' + tests.length + ' sample tender row(s) and their bid folders?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var folders = 0;
  tests.forEach(function (r) {
    var m = String(r.notes || '').match(/folders\/([A-Za-z0-9_-]+)/);
    if (m) { try { DriveApp.getFolderById(m[1]).setTrashed(true); folders++; } catch (e) { /* already gone */ } }
  });
  tests.map(function (r) { return r.rowNumber; }).sort(function (a, b) { return b - a; }).forEach(function (n) { t.sh.deleteRow(n); });
  ui.alert('Remove sample tenders', 'Removed ' + tests.length + ' sample tender(s) and ' + folders + ' bid folder(s) (moved to the Drive Bin).', ui.ButtonSet.OK);
}
