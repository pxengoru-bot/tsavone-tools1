// End-to-end run of every Tsavone Bids menu item against strict stand-ins for Google services.
// The stand-ins only accept real Apps Script method names, so a typo or wrong call fails the test.
// Run: node tenders/tests/flow.test.js
const fs = require('fs');
const vm = require('vm');
const path = require('path');

function strict(name, obj) {
  return new Proxy(obj, {
    get(t, prop) {
      if (typeof prop === 'symbol' || ['then', 'toJSON', 'inspect', 'constructor'].includes(prop)) return t[prop];
      if (!(prop in t)) throw new Error(`${name}.${String(prop)} is not a real Apps Script method (or not mocked)`);
      return t[prop];
    }
  });
}
function chain(name, methods, extra = {}) {
  const o = Object.assign({}, extra);
  let P;
  methods.forEach((m) => { o[m] = (...a) => { (o._calls = o._calls || []).push([m, a]); return P; }; });
  P = strict(name, o);
  return P;
}

/* ---------- Utilities ---------- */
function formatDate(d, tz, pattern) {
  const p = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
    .formatToParts(d).forEach((x) => { p[x.type] = x.value; });
  const mm = String(new Date(d.toLocaleString('en-US', { timeZone: tz })).getMonth() + 1).padStart(2, '0');
  return pattern.replace('yyyy', p.year).replace('MMMM', p.month).replace('MMM', p.month.slice(0, 3)).replace('MM', mm)
    .replace('dd', p.day.padStart(2, '0')).replace(/\bd\b/, p.day).replace('HH', p.hour).replace('mm', p.minute);
}
const Utilities = strict('Utilities', { formatDate });

/* ---------- Spreadsheet ---------- */
function makeSheet(ss, name) {
  const s = { _name: name, _grid: [], _cf: [], _formulas: {}, _active: [2, 1] };
  const lastRow = () => { let n = s._grid.length; while (n > 0 && (s._grid[n - 1] || []).every((v) => v === '' || v === undefined)) n--; return n; };
  const width = () => Math.max(0, ...s._grid.map((r) => { let w = (r || []).length; while (w > 0 && (r[w - 1] === '' || r[w - 1] === undefined)) w--; return w; }));
  const cellRow = (r, c) => { while (s._grid.length < r) s._grid.push([]); const row = s._grid[r - 1]; while (row.length < c) row.push(''); return row; };
  function range(r, c, nr = 1, nc = 1) {
    if (!(r >= 1 && c >= 1 && nr >= 1 && nc >= 1)) throw new Error(`bad range ${r},${c},${nr},${nc}`);
    const R = {
      getValues: () => { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (s._grid[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; },
      setValues: (vals) => { if (vals.length !== nr || vals.some((x) => x.length !== nc)) throw new Error(`setValues size ${vals.length}x${vals[0] && vals[0].length} vs ${nr}x${nc}`); vals.forEach((row, i) => row.forEach((v, j) => { cellRow(r + i, c + j)[c - 1 + j] = v; })); return P; },
      setValue: (v) => { cellRow(r, c)[c - 1] = v; return P; },
      setFormula: (f) => { if (!String(f).startsWith('=')) throw new Error('formula must start with ='); s._formulas[r + ',' + c] = f; cellRow(r, c)[c - 1] = '#FORMULA'; return P; },
      getRow: () => r, getLastRow: () => r + nr - 1, getSheet: () => SP
    };
    ['setBackground', 'setFontColor', 'setFontWeight', 'setVerticalAlignment', 'setWrap', 'setDataValidation', 'setNumberFormat', 'setFontSize'].forEach((m) => { R[m] = () => P; });
    const P = strict('Range', R);
    return P;
  }
  Object.assign(s, {
    getName: () => s._name, setName: (n) => { s._name = n; return SP; },
    getLastRow: lastRow, getLastColumn: width, getMaxRows: () => 1000,
    getRange: range, getActiveRange: () => range(s._active[0], s._active[1]),
    setFrozenRows: () => SP, setRowHeight: () => SP, setRowHeights: () => SP, setColumnWidth: () => SP,
    setConditionalFormatRules: (rules) => { s._cf = rules; }, clear: () => { s._grid = []; s._formulas = {}; return SP; },
    deleteRow: (n) => { s._grid.splice(n - 1, 1); return SP; }
  });
  const SP = strict('Sheet', s);
  return SP;
}
const book = { _sheets: [], _tz: 'America/New_York', _toasts: [] };
book._active = null;
Object.assign(book, {
  getSheets: () => book._sheets,
  getSheetByName: (n) => book._sheets.find((x) => x.getName() === n) || null,
  insertSheet: (n) => { const sh = makeSheet(BOOK, n); book._sheets.push(sh); return sh; },
  getSpreadsheetTimeZone: () => book._tz, setSpreadsheetTimeZone: (tz) => { book._tz = tz; },
  getActiveSheet: () => book._active, setActiveSheet: (sh) => { book._active = sh; return sh; },
  toast: (m) => { book._toasts.push(m); }, getUrl: () => 'https://docs.google.com/spreadsheets/d/TRACKER/edit'
});
const BOOK = strict('Spreadsheet', book);
book._sheets.push(makeSheet(BOOK, 'Sheet1'));
book._active = book._sheets[0];

const alerts = [], prompts = [];
let promptAnswer = '', alertAnswer = 'YES', sidebar = null;
const Ui = strict('Ui', {
  createMenu: () => chain('Menu', ['addItem', 'addSeparator', 'addSubMenu', 'addToUi']),
  alert: (...a) => { alerts.push(a.length === 1 ? a[0] : a[0] + ': ' + a[1]); return a[2] === 'YES_NO' || a[2] === 'YES_NO_CANCEL' ? alertAnswer : 'OK'; },
  prompt: (...a) => { prompts.push(a[1]); return strict('PromptResponse', { getSelectedButton: () => 'OK', getResponseText: () => promptAnswer }); },
  showSidebar: (out) => { sidebar = out; },
  ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' },
  Button: { OK: 'OK', CANCEL: 'CANCEL', CLOSE: 'CLOSE', YES: 'YES', NO: 'NO' }
});
const SpreadsheetApp = strict('SpreadsheetApp', {
  getActiveSpreadsheet: () => BOOK,
  getUi: () => Ui,
  newDataValidation: () => chain('DataValidationBuilder', ['requireValueInList', 'setAllowInvalid', 'build']),
  newConditionalFormatRule: () => {
    const rule = { formula: null, bg: null };
    const P = strict('ConditionalFormatRuleBuilder', {
      whenFormulaSatisfied: (f) => { rule.formula = f; return P; },
      setBackground: (c) => { rule.bg = c; return P; },
      setRanges: () => P,
      build: () => rule
    });
    return P;
  }
});

/* ---------- Drive ---------- */
let idn = 0;
const allFiles = [], allFolders = [];
function makeFile(name) {
  const f = { _name: name, _trashed: false, _id: 'f' + (++idn), _parent: null };
  Object.assign(f, {
    getId: () => f._id, getName: () => f._name, setTrashed: (v) => { f._trashed = v; return P; },
    moveTo: (folder) => { f._parent = folder; folder._files.push(P); return P; }
  });
  const P = strict('File', f);
  allFiles.push(P);
  return P;
}
function makeFolder(name) {
  const fo = { _name: name, _folders: [], _files: [], _trashed: false, _id: 'd' + (++idn) };
  const iter = (list) => { let i = 0; return strict('Iterator', { hasNext: () => i < list.length, next: () => list[i++] }); };
  Object.assign(fo, {
    getName: () => fo._name, getId: () => fo._id, getUrl: () => 'https://drive.google.com/drive/folders/' + fo._id,
    getFoldersByName: (n) => iter(fo._folders.filter((x) => x.getName() === n && !x._trashed)),
    createFolder: (n) => { const c = makeFolder(n); fo._folders.push(c); return c; },
    getFilesByName: (n) => iter(fo._files.filter((x) => x.getName() === n && !x._trashed)),
    setTrashed: (v) => { fo._trashed = v; return P; }
  });
  const P = strict('Folder', fo);
  allFolders.push(P);
  return P;
}
const root = makeFolder('My Drive');
const DriveApp = strict('DriveApp', {
  getRootFolder: () => root,
  getFileById: (id) => allFiles.find((f) => f.getId() === id) || (() => { throw new Error('no file ' + id); })(),
  getFolderById: (id) => allFolders.find((f) => f.getId() === id) || (() => { throw new Error('no folder ' + id); })()
});

/* ---------- Docs ---------- */
const docs = [];
function makeText(holder) {
  const T = {};
  ['setForegroundColor', 'setFontSize', 'setBold'].forEach((m) => { T[m] = (v) => { holder.style[m] = v; return P; }; });
  T.setLinkUrl = (u) => { holder.link = u; return P; };
  T.getText = () => holder.text;
  const P = strict('Text', T);
  return P;
}
function makePara(text) {
  const p = { kind: 'p', text: text || '', style: {} };
  Object.assign(p, {
    setText: (s) => { p.text = s; return P; }, setHeading: (h) => { p.heading = h; return P; },
    editAsText: () => makeText(p), asParagraph: () => P
  });
  const P = strict('Paragraph', p);
  return P;
}
function makeTable(cells) {
  const t = { kind: 'table', rows: cells.map((r) => r.map((v) => ({ text: String(v), style: {} }))) };
  const cell = (r, c) => {
    const C = t.rows[r][c];
    const CP = strict('TableCell', { setBackgroundColor: (v) => { C.style.bg = v; return CP; }, editAsText: () => makeText(C), getText: () => C.text });
    return CP;
  };
  Object.assign(t, {
    setBorderColor: () => P, getNumRows: () => t.rows.length, getCell: cell,
    getRow: (r) => strict('TableRow', { getNumCells: () => t.rows[r].length, getCell: (c) => cell(r, c) })
  });
  const P = strict('Table', t);
  return P;
}
const DocumentApp = strict('DocumentApp', {
  create: (name) => {
    const body = { children: [makePara('')] };
    Object.assign(body, {
      setMarginTop: () => B, setMarginBottom: () => B, setMarginLeft: () => B, setMarginRight: () => B,
      getChild: (i) => body.children[i],
      appendParagraph: (s) => { const p = makePara(s); body.children.push(p); return p; },
      appendTable: (c) => { const t = makeTable(c); body.children.push(t); return t; }
    });
    const B = strict('Body', body);
    const file = makeFile(name);
    const doc = { name, body, file };
    docs.push(doc);
    return strict('Document', { getBody: () => B, getId: () => file.getId(), getUrl: () => 'https://docs.google.com/document/d/' + file.getId(), saveAndClose: () => {} });
  },
  ParagraphHeading: { HEADING1: 'H1', HEADING2: 'H2' }
});

/* ---------- Mail, triggers, html ---------- */
const mails = [];
const MailApp = strict('MailApp', { sendEmail: (o) => { if (!o.to || !o.subject || !o.body) throw new Error('sendEmail needs to, subject, body'); mails.push(o); } });
let triggers = [];
const ScriptApp = strict('ScriptApp', {
  newTrigger: (fn) => {
    const spec = { fn };
    const P = strict('TriggerBuilder', {
      timeBased: () => P, atHour: (h) => { spec.hour = h; return P; }, everyDays: (n) => { spec.days = n; return P; },
      inTimezone: (tz) => { spec.tz = tz; return P; },
      create: () => { const tr = strict('Trigger', { getHandlerFunction: () => fn, _spec: spec }); triggers.push(tr); return tr; }
    });
    return P;
  },
  getProjectTriggers: () => triggers.slice(),
  deleteTrigger: (tr) => { triggers = triggers.filter((x) => x !== tr); }
});
let template = null;
const HtmlService = strict('HtmlService', {
  createTemplateFromFile: (n) => {
    if (n !== 'Sidebar') throw new Error('unknown file ' + n);
    template = {};
    template.evaluate = () => chain('HtmlOutput', ['setTitle']);
    return template;
  }
});

/* ---------- Load ---------- */
const ctx = { SpreadsheetApp, DriveApp, DocumentApp, MailApp, ScriptApp, HtmlService, Utilities, console };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx, { filename: 'Code.gs' });

let pass = 0, fail = 0;
function ok(name, cond, extra) { if (cond) pass++; else { fail++; console.log('FAIL', name, extra !== undefined ? JSON.stringify(extra).slice(0, 400) : ''); } }
function run(name, fn) { try { return fn(); } catch (e) { fail++; console.log('FAIL (error)', name, e.message, (e.stack || '').split('\n').slice(1, 3).join(' ')); } }
const tenders = () => BOOK.getSheetByName('Tenders');
const col = (n) => ctx.COLUMNS.indexOf(n);
const cell = (r, n) => (tenders()._grid[r - 1] || [])[col(n)];
const kampalaKey = (d) => formatDate(d, 'Africa/Kampala', 'yyyy-MM-dd');

/* 1. setup */
run('setup', () => ctx.setup());
ok('tab renamed Tenders', !!tenders());
ok('headers', tenders()._grid[0].join(',') === ctx.COLUMNS.join(','), tenders()._grid[0]);
ok('time zone Kampala', book._tz === 'Africa/Kampala');
ok('4 deadline colour rules', tenders()._cf.length === 4 && tenders()._cf[0].formula.startsWith('=AND(ISNUMBER($H2),NOT(OR($J2="Submitted"'), tenders()._cf.map((r) => r.formula));
const docSheet = BOOK.getSheetByName('Documents');
ok('documents prefilled', docSheet && docSheet.getLastRow() === 21, docSheet && docSheet.getLastRow());
ok('documents colour rules', docSheet._cf.length === 2 && docSheet._cf[0].formula.includes('$D2<TODAY()'), docSheet._cf.map((r) => r.formula));
const dash = BOOK.getSheetByName('Dashboard');
ok('dashboard formulas written', dash && Object.keys(dash._formulas).length >= 30, dash && Object.keys(dash._formulas).length);
run('setup twice', () => ctx.setup());
ok('documents not duplicated', docSheet.getLastRow() === 21);

/* 2. Add new tender (sidebar) */
run('open sidebar', () => ctx.showAddTenderSidebar());
ok('sidebar shown with lists', sidebar && template.lists.status.length === 8 && /^\d{4}-\d{2}-\d{2}$/.test(template.today));
const id1 = run('addTender', () => ctx.addTender({ organisation: 'Ministry X', title: 'Training of district staff in M&E', notice_type: 'RFP', category: 'Training', deadline: '2026-11-20', submission_method: 'Email', status: 'Watching', lead_person: 'Training lead', bid_value_ugx: '45,000,000', go_no_go: 'Undecided', documents_required: 'Certificate of incorporation, Valid tax clearance certificate\nCVs of key experts\nSigned bid form', notes: '' }));
ok('ref id format', /^TND-\d{4}-0001$/.test(id1), id1);
ok('deadline stored as Kampala date', kampalaKey(cell(2, 'deadline')) === '2026-11-20', cell(2, 'deadline'));
ok('value number', cell(2, 'bid_value_ugx') === 45000000);
ok('result pending', cell(2, 'result') === 'Pending');
let err = null; try { ctx.addTender({ organisation: '', title: '' }); } catch (e) { err = e.message; }
ok('required fields enforced', /organisation and the title/.test(err), err);

/* 3. onEdit gives a typed row its ID */
tenders()._grid[2] = []; tenders()._grid[2][col('organisation')] = 'Typed Org'; tenders()._grid[2][col('title')] = 'Typed tender';
tenders()._grid[2][col('notes')] = '=HYPERLINK("x")';
run('onEdit', () => ctx.onEdit({ range: tenders().getRange(3, col('title') + 1) }));
ok('typed row got next ID', /^TND-\d{4}-0002$/.test(cell(3, 'ref_id')), cell(3, 'ref_id'));
ok('defaults filled', cell(3, 'status') === 'Watching' && cell(3, 'go_no_go') === 'Undecided' && cell(3, 'result') === 'Pending');
ok('other cells untouched', cell(3, 'notes') === '=HYPERLINK("x")');

/* 4. Documents: mark some as Have with links */
const dg = docSheet._grid;
dg.forEach((r, i) => { if (i === 0) return; if (/Incorporation|tax clearance/i.test(r[0])) { r[1] = 'Have'; r[4] = 'https://drive.google.com/file/d/DOC' + i; } if (/CVs/.test(r[0])) r[1] = 'Missing'; });

/* 5. Bid folder and checklist for row 2 */
book._active = tenders(); tenders()._active = [2, 3];
run('create bid folder', () => ctx.menuCreateBidFolder());
const bids = root._folders.find((f) => f.getName() === 'Tsavone Bids');
const bidFolder = bids && bids._folders[0];
ok('folder created in Tsavone Bids', bidFolder && /^TND-\d{4}-0001 - Ministry X - Training of district staff in M&E$/.test(bidFolder.getName()), bidFolder && bidFolder.getName());
ok('checklist moved into folder', bidFolder && bidFolder._files.length === 1 && bidFolder._files[0].getName().startsWith('Bid checklist - TND-'));
const doc = docs[0];
const reqTable = doc && doc.body.children.find((c) => c.kind === 'table' && c.rows[0][1] && c.rows[0][1].text === 'Required by the notice');
ok('required docs table rows', reqTable && reqTable.rows.length === 5, reqTable && reqTable.rows.map((r) => r.map((c) => c.text)));
ok('matched statuses', reqTable && reqTable.rows.slice(1).map((r) => r[3].text).join('|') === 'Have|Have|Missing|Not in Documents tab', reqTable && reqTable.rows.slice(1).map((r) => r[3].text));
ok('drive link attached', reqTable && reqTable.rows[1][4].link && reqTable.rows[1][4].link.startsWith('https://drive.google.com/file/d/DOC'));
const stepsTable = doc && doc.body.children.find((c) => c.kind === 'table' && c.rows[0][1] && c.rows[0][1].text === 'Step');
ok('standard steps', stepsTable && stepsTable.rows.length === 8 && stepsTable.rows.map((r) => r[1].text).includes('Submission and proof of submission'));
ok('notes has folder link', String(cell(2, 'notes')).includes('Bid folder: https://drive.google.com/drive/folders/'), cell(2, 'notes'));
ok('documents_ready filled', cell(2, 'documents_ready') === '2 of 4 ready', cell(2, 'documents_ready'));
ok('folder message', /2 of 4|3 of 4/.test(alerts[alerts.length - 1]), alerts[alerts.length - 1]);
alertAnswer = 'YES';
run('create bid folder again', () => ctx.menuCreateBidFolder());
ok('old checklist replaced', bidFolder._files.filter((f) => !f._trashed).length === 1 && bidFolder._files.length === 2);
ok('notes link not duplicated', (String(cell(2, 'notes')).match(/Bid folder:/g) || []).length === 1);
tenders()._active = [1, 1];
run('header row selected', () => ctx.menuCreateBidFolder());
ok('friendly error on header', /Click a cell in a tender/.test(alerts[alerts.length - 1]), alerts[alerts.length - 1]);

/* 6. Mark as submitted */
tenders()._active = [2, 1]; promptAnswer = '5/11/2026';
run('mark submitted', () => ctx.menuMarkSubmitted());
ok('status submitted', cell(2, 'status') === 'Submitted');
ok('submitted date', kampalaKey(cell(2, 'submitted_date')) === '2026-11-05', cell(2, 'submitted_date'));
ok('go set', cell(2, 'go_no_go') === 'Go');
tenders()._active = [3, 1]; promptAnswer = 'someday';
run('mark submitted bad date', () => ctx.menuMarkSubmitted());
ok('bad date refused', cell(3, 'status') === 'Watching' && /not understood/.test(alerts[alerts.length - 1]));

/* 7. Dashboard */
run('refresh dashboard', () => ctx.refreshDashboard());
ok('dashboard active', book._active === BOOK.getSheetByName('Dashboard'));
const fx = Object.values(BOOK.getSheetByName('Dashboard')._formulas);
ok('due lists and docs formulas', fx.filter((f) => f.startsWith('=IFERROR(SORT(FILTER(')).length === 2 && fx.some((f) => f.includes("'Documents'!")));

/* 8. Reminder on / off */
run('install', () => ctx.installReminder());
run('install again', () => ctx.installReminder());
ok('one trigger at 7 Kampala', triggers.length === 1 && triggers[0]._spec.hour === 7 && triggers[0]._spec.tz === 'Africa/Kampala' && triggers[0]._spec.days === 1, triggers.map((t) => t._spec));
run('remove', () => ctx.removeReminder());
ok('trigger removed', triggers.length === 0);

/* 9. Reminder email: only when needed */
run('reminder with nothing due', () => ctx.menuSendReminderNow());
ok('no email when nothing due', mails.length === 0 && /no email was sent/.test(alerts[alerts.length - 1]), alerts[alerts.length - 1]);
run('add samples', () => ctx.menuAddSampleTenders());
ok('3 samples added', ctx.readRows_(tenders(), ctx.headerMap_(tenders())).filter((r) => String(r.notes).startsWith('TEST ROW')).length === 3);
ok('sample ids continue', /^TND-\d{4}-0003$/.test(cell(4, 'ref_id')), cell(4, 'ref_id'));
dg.forEach((r, i) => { if (/Trading licence/.test(r[0])) { r[1] = 'Have'; r[3] = new Date(Date.now() + 10 * 86400000); } });
run('reminder with items', () => ctx.menuSendReminderNow());
ok('one email to owner', mails.length === 1 && mails[0].to === 'p.engoru@tsavone.com', mails.map((m) => m.to));
ok('email lists 2 tenders due within 7 days and the licence', mails[0] && /2 tender\(s\) due within 7 days, 1 document\(s\) expiring/.test(mails[0].subject), mails[0] && mails[0].subject);
ok('email has sheet link', mails[0] && mails[0].body.includes('https://docs.google.com/spreadsheets/d/TRACKER/edit') && mails[0].htmlBody.includes('Open the Tender Tracker'));
ok('12-day sample not in email', mails[0] && !mails[0].body.includes('EUDR'));

/* 10. Sample folder + removal */
book._active = tenders(); tenders()._active = [4, 1];
run('folder for sample', () => ctx.menuCreateBidFolder());
const sampleFolder = bids._folders.find((f) => f.getName().includes('Example District'));
run('remove samples', () => ctx.menuRemoveSampleTenders());
ok('samples removed', ctx.readRows_(tenders(), ctx.headerMap_(tenders())).every((r) => !String(r.notes).startsWith('TEST ROW')) && tenders().getLastRow() === 3, tenders().getLastRow());
ok('sample folder trashed', sampleFolder && sampleFolder._trashed === true);
ok('real tenders kept', cell(2, 'organisation') === 'Ministry X' && cell(3, 'organisation') === 'Typed Org');

console.log('Alerts:\n  ' + alerts.map((a) => a.replace(/\n+/g, ' / ')).join('\n  '));
console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
