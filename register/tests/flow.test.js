// End-to-end run of every menu item against strict stand-ins for Google services.
// The stand-ins only accept real Apps Script method names, so a typo or wrong call fails the test.
// It also writes certificate-preview.html (a rough HTML rendering of the certificate layout).
// Run: NODE_PATH=<folder with node_modules/jsqr> node register/tests/flow.test.js
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const TZ = 'Africa/Kampala';
const calls = [];

function strict(name, obj) {
  return new Proxy(obj, {
    get(t, prop) {
      if (typeof prop === 'symbol' || prop === 'then' || prop === 'toJSON' || prop === 'inspect' || prop === 'constructor') return t[prop];
      if (!(prop in t)) throw new Error(`${name}.${String(prop)} is not a real Apps Script method (or not mocked)`);
      return t[prop];
    }
  });
}
function chain(name, methods, self) {
  const o = self || {};
  methods.forEach((m) => { o[m] = (...a) => { calls.push(`${name}.${m}`); return proxy; }; });
  const proxy = strict(name, o);
  return proxy;
}

/* ---------------- Utilities ---------------- */
function formatDate(d, tz, pattern) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: 'long', day: 'numeric' }).formatToParts(d).forEach((p) => { parts[p.type] = p.value; });
  const mm = String(new Date(d.toLocaleString('en-US', { timeZone: tz })).getMonth() + 1).padStart(2, '0');
  return pattern.replace('yyyy', parts.year).replace('MMMM', parts.month).replace('MMM', parts.month.slice(0, 3)).replace('MM', mm).replace('dd', parts.day.padStart(2, '0')).replace(/\bd\b/, parts.day);
}
function makeBlob(bytes, type, name) {
  const b = { _bytes: bytes, _type: type, _name: name };
  b.setName = (n) => { b._name = n; return proxyB; };
  b.getName = () => b._name;
  b.getContentType = () => b._type;
  b.getBytes = () => b._bytes;
  const proxyB = strict('Blob', b);
  return proxyB;
}
const Utilities = strict('Utilities', {
  formatDate,
  newBlob: (bytes, type, name) => {
    if (!Array.isArray(bytes) || bytes.some((x) => x < -128 || x > 127 || !Number.isInteger(x))) throw new Error('newBlob needs a byte array of -128..127');
    return makeBlob(bytes, type, name);
  },
  base64Encode: (bytes) => Buffer.from(bytes.map((b) => b & 255)).toString('base64'),
  base64Decode: (str) => [...Buffer.from(str, 'base64')].map((b) => (b > 127 ? b - 256 : b))
});

/* ---------------- Spreadsheet ---------------- */
function makeSheet(name, grid) {
  const s = { _name: name, _grid: grid || [], _frozen: 0, _active: null };
  const width = () => Math.max(0, ...s._grid.map((r) => { let w = r.length; while (w > 0 && (r[w - 1] === '' || r[w - 1] === undefined)) w--; return w; }));
  const lastRow = () => { let n = s._grid.length; while (n > 0 && s._grid[n - 1].every((v) => v === '' || v === undefined)) n--; return n; };
  const cell = (r, c) => { while (s._grid.length < r) s._grid.push([]); const row = s._grid[r - 1]; while (row.length < c) row.push(''); return row; };
  function range(r, c, nr = 1, nc = 1) {
    const R = {
      getValues: () => { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (s._grid[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; },
      setValues: (vals) => { if (vals.length !== nr || vals.some((row) => row.length !== nc)) throw new Error(`setValues size mismatch ${vals.length}x${vals[0] && vals[0].length} vs ${nr}x${nc}`); vals.forEach((row, i) => row.forEach((v, j) => { cell(r + i, c + j)[c - 1 + j] = v; })); return P; },
      setValue: (v) => { cell(r, c)[c - 1] = v; return P; },
      getRow: () => r, getNumRows: () => nr, getNumColumns: () => nc
    };
    ['setBackground', 'setFontColor', 'setFontWeight', 'setVerticalAlignment', 'setWrap', 'setDataValidation', 'setNumberFormat', 'setFontSize', 'setBorder', 'setHorizontalAlignment', 'merge']
      .forEach((m) => { R[m] = () => P; });
    const P = strict('Range', R);
    return P;
  }
  Object.assign(s, {
    getName: () => s._name, setName: (n) => { s._name = n; return P; },
    getLastRow: lastRow, getLastColumn: width, getMaxRows: () => 1000, getMaxColumns: () => 26,
    getRange: range, getDataRange: () => range(1, 1, Math.max(lastRow(), 1), Math.max(width(), 1)),
    setFrozenRows: () => P, setRowHeight: () => P, setRowHeights: () => P, setColumnWidth: () => P,
    deleteRow: (n) => { s._grid.splice(n - 1, 1); return P; },
    getActiveRangeList: () => (s._active ? strict('RangeList', { getRanges: () => s._active.map(([a, b, cols]) => range(a, 1, b - a + 1, cols || 26)) }) : null)
  });
  const P = strict('Sheet', s);
  return P;
}
function makeSpreadsheet(id, sheets) {
  const ss = { _sheets: sheets, _active: sheets[0], _toasts: [] };
  Object.assign(ss, {
    getId: () => id,
    getSheets: () => ss._sheets,
    getSheetByName: (n) => ss._sheets.find((x) => x.getName() === n) || null,
    insertSheet: (n) => { const sh = makeSheet(n); ss._sheets.push(sh); return sh; },
    deleteSheet: (sh) => { ss._sheets = ss._sheets.filter((x) => x !== sh); },
    getSpreadsheetTimeZone: () => ss._tz || 'America/New_York',
    setSpreadsheetTimeZone: (tz) => { ss._tz = tz; },
    getActiveSheet: () => ss._active,
    setActiveSheet: (sh) => { ss._active = sh; return sh; },
    toast: (m) => { ss._toasts.push(m); }
  });
  return strict('Spreadsheet', ss);
}

const register = makeSheet('Sheet1');
const book = makeSpreadsheet('register-id', [register]);
const courseBook = makeSpreadsheet('course-id', [makeSheet('Courses', [
  ['code', 'track_no', 'track_name', 'title', 'days'],
  ['T1.1', 1, 'Evidence', 'Practical Monitoring, Evaluation and Learning for Programmes', 5],
  ['T8.1', 8, 'Governance', 'Safeguarding and Protection from Sexual Exploitation, Abuse and Harassment', 1]
])]);
// The course sheet uses "title" not "days" header order; loadCourses_ finds columns by name
courseBook.getSheets()[0]._grid = courseBook.getSheets()[0]._grid.map((r) => [r[0], r[1], r[2], r[3], r[4]]);

const alerts = [], prompts = [];
let promptAnswer = '1', alertAnswer = 'YES';
const Ui = strict('Ui', {
  createMenu: (t) => chain('Menu', ['addItem', 'addSeparator', 'addSubMenu', 'addToUi']),
  alert: (...a) => { alerts.push(a.length === 1 ? a[0] : a[0] + ': ' + a[1]); return a[2] === 'OK_CANCEL' ? 'OK' : a[2] === 'OK' ? 'OK' : alertAnswer; },
  prompt: (...a) => { prompts.push(a[1]); return strict('PromptResponse', { getSelectedButton: () => 'OK', getResponseText: () => promptAnswer }); },
  ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' },
  Button: { OK: 'OK', CANCEL: 'CANCEL', CLOSE: 'CLOSE', YES: 'YES', NO: 'NO' }
});
const validationBuilder = () => chain('DataValidationBuilder', ['requireValueInList', 'setAllowInvalid', 'build']);
const SpreadsheetApp = strict('SpreadsheetApp', {
  getActiveSpreadsheet: () => book,
  openById: (id) => (id === 'register-id' ? book : id === '10V8_jXPTFwRl2y940v_AtmMDPU8RGQBI6yH5qDy6N4I' ? courseBook : (() => { throw new Error('no sheet ' + id); })()),
  getUi: () => Ui,
  newDataValidation: validationBuilder,
  flush: () => {},
  BorderStyle: { SOLID: 'SOLID' }
});

/* ---------------- Drive ---------------- */
const files = [];
function makeFile(name, blob, parent) {
  const f = { _name: name, _blob: blob, _trashed: false, _parent: parent, _id: 'file' + files.length };
  Object.assign(f, {
    getName: () => f._name, getId: () => f._id, getBlob: () => f._blob,
    getAs: (t) => { if (t !== 'application/pdf') throw new Error('getAs type'); return makeBlob([37, 80, 68, 70], 'application/pdf', f._name + '.pdf'); },
    setTrashed: (v) => { f._trashed = v; return P; }, isTrashed: () => f._trashed
  });
  const P = strict('File', f);
  files.push(P);
  return P;
}
function makeFolder(name) {
  const fo = { _name: name, _folders: [], _files: [] };
  const iter = (list) => { let i = 0; return strict('Iterator', { hasNext: () => i < list.length, next: () => list[i++] }); };
  Object.assign(fo, {
    getName: () => fo._name,
    getFoldersByName: (n) => iter(fo._folders.filter((x) => x.getName() === n)),
    createFolder: (n) => { const c = makeFolder(n); fo._folders.push(c); return c; },
    getFiles: () => iter(fo._files.filter((x) => !x.isTrashed())),
    getFilesByName: (n) => iter(fo._files.filter((x) => !x.isTrashed() && x.getName() === n)),
    createFile: (blob) => { const f = makeFile(blob.getName(), blob, P); fo._files.push(f); return f; }
  });
  const P = strict('Folder', fo);
  return P;
}
const rootFolder = makeFolder('My Drive');
const docFiles = {};
const DriveApp = strict('DriveApp', {
  getRootFolder: () => rootFolder,
  getFileById: (id) => docFiles[id] || files.find((f) => f.getId() === id) || (() => { throw new Error('file not found ' + id); })()
});

/* ---------------- Docs (records a layout tree) ---------------- */
const docs = [];
function makeText(owner) {
  const t = {};
  ['setFontSize', 'setForegroundColor', 'setBold', 'setFontFamily'].forEach((m) => {
    t[m] = (v) => { owner.style[m] = v; return P; };
  });
  const P = strict('Text', t);
  return P;
}
function makeParagraph(text) {
  const p = { kind: 'p', text: text || '', style: {}, images: [], children: [] };
  Object.assign(p, {
    setAlignment: (a) => { p.style.align = a; return P; },
    setSpacingBefore: (v) => { p.style.before = v; return P; },
    setSpacingAfter: (v) => { p.style.after = v; return P; },
    setLineSpacing: (v) => { p.style.line = v; return P; },
    setAttributes: (a) => { Object.assign(p.style, a); return P; },
    setText: (s) => { p.text = s; return P; },
    getText: () => p.text,
    getNumChildren: () => p.images.length + (p.text ? 1 : 0),
    appendText: (s) => { p.text += s; return makeText(p); },
    editAsText: () => { if (!p.text) throw new Error('editAsText on empty paragraph would fail'); return makeText(p); },
    appendInlineImage: (blob) => {
      if (!['image/png', 'image/jpeg', 'image/gif'].includes(blob.getContentType())) throw new Error('bad image type');
      const img = { w: 400, h: 357, blob };
      if (blob.getName() === 'qr.png') { img.w = 456; img.h = 456; }
      p.images.push(img);
      return strict('InlineImage', { getWidth: () => img.w, getHeight: () => img.h, setWidth: (v) => { img.w = v; return I; }, setHeight: (v) => { img.h = v; return I; } });
      var I;
    },
    getType: () => 'PARAGRAPH',
    asParagraph: () => P
  });
  const P = strict('Paragraph', p);
  // keep the InlineImage chain working
  const orig = p.appendInlineImage;
  p.appendInlineImage = (blob) => { const r = orig(blob); const img = p.images[p.images.length - 1]; const I = strict('InlineImage', { getWidth: () => img.w, getHeight: () => img.h, setWidth: (v) => { img.w = v; return I; }, setHeight: (v) => { img.h = v; return I; } }); return I; };
  return P;
}
function makeContainer(kind) {
  const c = { kind, children: [], style: {} };
  const P0 = {
    appendParagraph: (t) => { const p = makeParagraph(t); c.children.push(p); return p; },
    appendTable: (cells) => {
      const t = makeTable(cells);
      c.children.push(t);
      c.children.push(makeParagraph('')); // Docs always keeps a paragraph after a table
      return t;
    },
    getChild: (i) => { if (!c.children[i]) throw new Error('no child ' + i); return c.children[i]; },
    getNumChildren: () => c.children.length
  };
  return { c, P0 };
}
function makeTable(cells) {
  const t = { kind: 'table', rows: [], style: { widths: {} } };
  cells.forEach((row) => {
    const r = { cells: [], style: {} };
    row.forEach(() => r.cells.push(makeCell()));
    t.rows.push(r);
  });
  Object.assign(t, {
    setBorderWidth: (v) => { t.style.border = v; return P; },
    setBorderColor: (v) => { t.style.borderColor = v; return P; },
    setColumnWidth: (i, w) => { if (typeof w !== 'number' || w <= 0) throw new Error('bad width'); t.style.widths[i] = w; return P; },
    getCell: (r, c) => t.rows[r].cells[c],
    getType: () => 'TABLE',
    getRow: (r) => strict('TableRow', { setMinimumHeight: (h) => { t.rows[r].style.minH = h; return null; } })
  });
  const P = strict('Table', t);
  return P;
}
function makeCell() {
  const { c, P0 } = makeContainer('cell');
  c.children.push(makeParagraph(''));
  Object.assign(c, P0, {
    setBackgroundColor: (v) => { c.style.bg = v; return P; },
    setPaddingTop: (v) => { c.style.pt = v; return P; }, setPaddingBottom: (v) => { c.style.pb = v; return P; },
    setPaddingLeft: (v) => { c.style.pl = v; return P; }, setPaddingRight: (v) => { c.style.pr = v; return P; },
    setVerticalAlignment: (v) => { c.style.valign = v; return P; }
  });
  const P = strict('TableCell', c);
  return P;
}
const DocumentApp = strict('DocumentApp', {
  create: (name) => {
    const { c: body, P0 } = makeContainer('body');
    body.children.push(makeParagraph(''));
    Object.assign(body, P0, {
      setPageWidth: (v) => { body.style.w = v; return B; }, setPageHeight: (v) => { body.style.h = v; return B; },
      setMarginTop: (v) => { body.style.mt = v; return B; }, setMarginBottom: (v) => { body.style.mb = v; return B; },
      setMarginLeft: (v) => { body.style.ml = v; return B; }, setMarginRight: (v) => { body.style.mr = v; return B; }
    });
    const B = strict('Body', body);
    const id = 'doc' + docs.length;
    const doc = strict('Document', { getBody: () => B, getId: () => id, saveAndClose: () => {} });
    docs.push({ name, body });
    docFiles[id] = makeFile(name, null, null);
    return doc;
  },
  HorizontalAlignment: { CENTER: 'center', LEFT: 'left' },
  VerticalAlignment: { BOTTOM: 'bottom', CENTER: 'middle' },
  ElementType: { PARAGRAPH: 'PARAGRAPH', TABLE: 'TABLE' },
  Attribute: { FONT_SIZE: 'FONT_SIZE' }
});
// Tables are not paragraphs: give them getType for shrinkEmptyParagraphs_
const origMakeTable = makeTable;

/* ---------------- Gmail, Properties, Script, Html ---------------- */
const drafts = [];
const GmailApp = strict('GmailApp', {
  createDraft: (to, subject, body, opts) => { drafts.push({ to, subject, body, opts, deleted: false }); return {}; },
  getDrafts: () => drafts.filter((d) => !d.deleted).map((d) => strict('GmailDraft', {
    getMessage: () => strict('GmailMessage', { getTo: () => d.to, getSubject: () => d.subject }),
    deleteDraft: () => { d.deleted = true; }
  }))
});
const Session = strict('Session', {
  getActiveUser: () => strict('User', { getEmail: () => 'px.engoru@gmail.com' }),
  getEffectiveUser: () => strict('User', { getEmail: () => 'px.engoru@gmail.com' })
});
const props = {};
const PropertiesService = strict('PropertiesService', { getScriptProperties: () => strict('Properties', { setProperty: (k, v) => { props[k] = v; }, getProperty: (k) => props[k] || null }) });
const ScriptApp = strict('ScriptApp', { getService: () => strict('Service', { getUrl: () => 'https://script.google.com/macros/s/AKfycbxEXAMPLE1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ/exec' }) });
let lastTemplate = null;
const HtmlService = strict('HtmlService', {
  createTemplateFromFile: (n) => {
    if (n !== 'Verify') throw new Error('unknown template ' + n);
    lastTemplate = {};
    lastTemplate.evaluate = () => chain('HtmlOutput', ['setTitle', 'addMetaTag']);
    return lastTemplate;
  }
});

/* ---------------- Load the project ---------------- */
const ctx = { Utilities, SpreadsheetApp, DriveApp, DocumentApp, GmailApp, PropertiesService, ScriptApp, HtmlService, Session, console };
vm.createContext(ctx);
['Code.gs'].forEach((f) => vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), ctx, { filename: f }));

// Give table objects getType (Docs Table element)
const _mt = makeTable;

let pass = 0, fail = 0;
function ok(name, cond, extra) { if (cond) pass++; else { fail++; console.log('FAIL', name, extra !== undefined ? JSON.stringify(extra) : ''); } }
function run(name, fn) { try { fn(); } catch (e) { fail++; console.log('FAIL (error)', name, e.message, e.stack.split('\n').slice(1, 3).join(' ')); } }

/* ---------------- 1. setup ---------------- */
run('setup', () => ctx.setup());
const reg = book.getSheetByName('Register');
ok('register renamed', !!reg);
ok('headers written', reg.getRange(1, 1, 1, 19).getValues()[0].join(',') === ctx.COLUMNS.join(','));
const settingsSheet = book.getSheetByName('Settings');
ok('settings created', settingsSheet && settingsSheet.getLastRow() === 1 + ctx.DEFAULT_SETTINGS.length);
ok('spreadsheet id stored', props.SPREADSHEET_ID === 'register-id');
ok('time zone set to Kampala', book.getSpreadsheetTimeZone() === 'Africa/Kampala');
const sgrid = settingsSheet._grid;
const setting = (k) => sgrid.find((r) => r[0] === k)[1];
const logoId = setting('logo_drive_file_id');
const logoFile = files.find((f) => f.getId() === logoId);
ok('setup saved the built-in logo to Drive', !!logoFile && logoFile.getName() === 'Tsavone logo for certificates.png' && rootFolder._files.includes(logoFile), logoId);
ok('logo bytes are the PNG', logoFile && Buffer.from(logoFile.getBlob().getBytes().map((b) => b & 255)).equals(fs.readFileSync(path.join(__dirname, '..', 'tsavone-logo.png'))));
run('setup twice keeps values', () => ctx.setup());
ok('settings not duplicated', settingsSheet.getLastRow() === 1 + ctx.DEFAULT_SETTINGS.length);
ok('logo not duplicated', setting('logo_drive_file_id') === logoId && rootFolder._files.filter((f) => f.getName() === 'Tsavone logo for certificates.png').length === 1);
// Opening the published page once fills in verify_url
ok('verify_url starts empty', setting('verify_url') === '');
run('doGet first visit', () => ctx.doGet({ parameter: {} }));
ok('verify_url filled automatically', /\/exec$/.test(setting('verify_url')), setting('verify_url'));

/* ---------------- 2. sample rows from the test file ---------------- */
const sample = fs.readFileSync(path.join(__dirname, '..', 'sample-rows.tsv'), 'utf8').replace(/\n+$/, '').split('\n').map((l) => l.split('\t'));
ok('sample has 3 rows of 19 columns', sample.length === 3 && sample.every((r) => r.length === 19), sample.map((r) => r.length));
false && sample.forEach((r, i) => {
  const row = r.map((v, j) => {
    if (['start_date', 'end_date', 'issued_date'].includes(ctx.COLUMNS[j]) && v) return new Date(v + 'T00:00:00+03:00');
    if (['attendance_percent', 'pre_test_score', 'post_test_score'].includes(ctx.COLUMNS[j]) && v) return Number(v);
    return v;
  });
  reg._grid[i + 1] = row;
});

/* ---------------- 3. menu items ---------------- */
run('add test participants', () => ctx.menuAddTestParticipants());
ok('3 test rows added at row 2', reg.getRange(2, 2, 3, 1).getValues().map((r) => r[0]).join('|') === 'Test Participant One|Test Participant Two|Test Participant Three', reg._grid.slice(1, 4));
ok('test emails use own address', reg.getRange(2, 4).getValues()[0][0] === 'px.engoru+test1@gmail.com', reg.getRange(2, 4).getValues());
ok('test dates written as ISO text for the sheet to parse', reg.getRange(2, 8).getValues()[0][0] === '2027-03-10');
// Sheets turns ISO text into dates; do the same here
for (let r = 1; r <= 3; r++) for (const c of [7, 8]) reg._grid[r][c] = new Date(reg._grid[r][c] + 'T00:00:00+03:00');
run('add test participants twice', () => ctx.menuAddTestParticipants());
ok('no duplicate test rows', reg.getLastRow() === 4, reg.getLastRow());
run('onOpen', () => ctx.onOpen());
run('fill titles', () => ctx.menuFillCourseTitles());
ok('titles filled', reg.getRange(2, 7, 3, 1).getValues().every((r) => r[0] === 'Practical Monitoring, Evaluation and Learning for Programmes'), alerts[alerts.length - 1]);
run('calculate', () => ctx.menuCalculateResults());
const res = reg.getRange(2, 15, 3, 2).getValues();
ok('results', JSON.stringify(res) === JSON.stringify([['Pass', 'Competence'], ['Pass', 'Completion'], ['Did not complete', '']]), res);
ok('improvement note', String(reg.getRange(2, 19).getValues()[0][0]).includes('Improvement: +37 points'), reg.getRange(2, 19).getValues());
run('ids', () => ctx.menuCreateCertificateIds());
const ids = reg.getRange(2, 1, 3, 1).getValues().map((r) => r[0]);
ok('ids created for passing rows only', JSON.stringify(ids) === JSON.stringify(['TSV-2027-T1.1-0001', 'TSV-2027-T1.1-0002', '']), ids);

// Generate: no selection -> all passing rows without issued date
reg._active = [[2, 2, 1]]; book.setActiveSheet(reg); alertAnswer = 'YES';
run('generate', () => ctx.menuGenerateCertificates());
ok('single cell selection does not limit generation', /Make certificates for the 2 passing/.test(alerts[alerts.length - 2]), alerts.slice(-2));
ok('generate message', /2 certificate\(s\) saved/.test(alerts[alerts.length - 1]) && !/Skipped/.test(alerts[alerts.length - 1]), alerts[alerts.length - 1]);
const certRoot = rootFolder._folders.find((f) => f.getName() === 'Tsavone Certificates');
const courseFolder = certRoot && certRoot._folders.find((f) => f.getName() === 'T1.1');
const pdfs = courseFolder ? courseFolder._files.filter((f) => !f.isTrashed()).map((f) => f.getName()) : [];
ok('pdfs saved in folder/course', JSON.stringify(pdfs) === JSON.stringify(['TSV-2027-T1.1-0001 - Test Participant One.pdf', 'TSV-2027-T1.1-0002 - Test Participant Two.pdf']), pdfs);
ok('temporary docs trashed', Object.values(docFiles).every((f) => f.isTrashed()));
const st = reg.getRange(2, 17, 3, 2).getValues();
ok('issued and status set', st[0][1] === 'Issued' && st[1][1] === 'Issued' && st[2][1] === 'Draft' && Object.prototype.toString.call(st[0][0]) === '[object Date]', st);

// Regenerate selected row 2 -> replaces file, keeps issued date
const firstIssued = st[0][0];
reg._active = [[2, 2]]; alertAnswer = 'YES';
run('regenerate selected', () => ctx.menuGenerateCertificates());
const live = courseFolder._files.filter((f) => !f.isTrashed()).map((f) => f.getName());
ok('regenerate replaces, no duplicates', live.length === 2, live);
ok('issued date kept', reg.getRange(2, 17).getValues()[0][0] === firstIssued);

// Email drafts for selected rows 2-4 (row 4 failed -> skipped)
reg._active = [[2, 4]];
run('email', () => ctx.menuEmailCertificates());
ok('2 drafts', drafts.length === 2, drafts.length);
ok('draft has pdf + name', drafts[0] && drafts[0].opts.attachments.length === 1 && drafts[0].opts.name === 'Tsavone Consulting Training');
ok('draft link', drafts[0] && drafts[0].body.includes('/exec?id=TSV-2027-T1.1-0001'));
ok('email skipped failing row', /4 \(status is not Issued\)/.test(alerts[alerts.length - 1]), alerts[alerts.length - 1]);

// Course summary
promptAnswer = '1';
run('summary', () => ctx.menuCourseSummary());
const sum = book.getSheets().find((s) => s.getName().startsWith('Summary T1.1'));
ok('summary tab', !!sum, book.getSheets().map((s) => s.getName()));
const sumText = sum ? sum._grid.map((r) => r.join(' | ')).join('\n') : '';
ok('summary numbers', /Participants \| 3/.test(sumText) && /Pass rate \| 67%/.test(sumText) && /Average improvement \(points\) \| \+23\.3/.test(sumText), sumText);
run('summary again replaces tab', () => ctx.menuCourseSummary());
ok('one summary tab', book.getSheets().filter((s) => s.getName().startsWith('Summary T1.1')).length === 1);

// Verification page
run('doGet valid', () => ctx.doGet({ parameter: { id: 'TSV-2027-T1.1-0001' } }));
ok('template result valid', lastTemplate.result && lastTemplate.result.state === 'VALID' && lastTemplate.result.name === 'Test Participant One', lastTemplate.result);
ok('template logo', String(lastTemplate.logo).startsWith('data:image/png;base64,'));
run('doGet empty', () => ctx.doGet({ parameter: {} }));
ok('no id -> no result', lastTemplate.result === null);

// Revoke and check
reg._grid[1][17] = 'Revoked';
run('doGet revoked', () => ctx.doGet({ parameter: { id: 'TSV-2027-T1.1-0001' } }));
ok('revoked', lastTemplate.result.state === 'REVOKED');

/* ---------------- Remove test participants ---------------- */
reg._grid[1][17] = 'Issued';
alertAnswer = 'YES';
run('remove test participants', () => ctx.menuRemoveTestParticipants());
ok('test rows removed', reg.getLastRow() === 1, reg.getLastRow());
ok('test pdfs trashed', courseFolder._files.every((f) => f.isTrashed()), courseFolder._files.map((f) => [f.getName(), f.isTrashed()]));
ok('test drafts deleted', drafts.every((d) => d.deleted));
ok('test summary tab deleted', !book.getSheets().some((s) => s.getName().startsWith('Summary T1.1')));
ok('remove message', /Removed 3 test row\(s\), 2 certificate PDF\(s\).*2 Gmail draft\(s\) and 1 summary tab/.test(alerts[alerts.length - 1]), alerts[alerts.length - 1]);

/* ---------------- 4. Certificate layout preview ---------------- */
const cert = docs[0].body;
function css(o) { return Object.entries(o).map(([k, v]) => `${k}:${v}`).join(';'); }
function renderP(p) {
  const s = p.style;
  const size = s.setFontSize || s.FONT_SIZE || 11;
  const st = { margin: 0, 'padding-top': (s.before || 0) + 'pt', 'padding-bottom': (s.after || 0) + 'pt', 'text-align': s.align || 'left', 'font-size': size + 'pt', 'line-height': (s.line || 1.15) * 1.18, color: s.setForegroundColor || '#000', 'font-weight': s.setBold ? 700 : 400, 'font-family': `'${s.setFontFamily || 'Arial'}', Arial, sans-serif` };
  const imgs = p.images.map((im) => `<img src="data:image/png;base64,${Buffer.from(im.blob.getBytes().map((b) => b & 255)).toString('base64')}" style="width:${im.w}pt;height:${im.h}pt;vertical-align:bottom">`).join('');
  return `<p style="${css(st)}">${imgs}${p.text.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</p>`;
}
function renderChildren(list) {
  return list.map((ch) => (ch.kind === 'table' ? renderTable(ch) : renderP(ch))).join('');
}
function renderTable(t) {
  const rows = t.rows.map((r) => '<tr>' + r.cells.map((c, i) => {
    const s = c.style;
    const st = { background: s.bg || 'transparent', padding: `${s.pt || 5}pt ${s.pr || 5}pt ${s.pb || 5}pt ${s.pl || 5}pt`, 'vertical-align': s.valign || 'top', width: (t.style.widths[i] || 'auto') + (t.style.widths[i] ? 'pt' : ''), height: r.style.minH ? r.style.minH + 'pt' : 'auto', 'box-sizing': 'border-box' };
    return `<td style="${css(st)}">${renderChildren(c.children)}</td>`;
  }).join('') + '</tr>').join('');
  return `<table style="border-collapse:collapse;table-layout:fixed">${rows}</table>`;
}
const html = `<!doctype html><html><head><meta charset="utf-8"><link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&family=Source+Sans+3:wght@400;600&display=swap" rel="stylesheet">
<style>body{margin:0;background:#888}.page{width:${cert.style.w}pt;height:${cert.style.h}pt;background:#fff;padding:${cert.style.mt}pt ${cert.style.mr}pt ${cert.style.mb}pt ${cert.style.ml}pt;box-sizing:border-box;overflow:hidden;position:relative}.page.over{outline:4px solid red}</style></head>
<body><div class="page" id="page"><div id="content">${renderChildren(cert.children)}</div></div></body></html>`;
fs.writeFileSync(path.join(require('os').tmpdir(), 'certificate-preview.html'), html);
ok('page is A4 landscape', Math.round(cert.style.w) === 842 && Math.round(cert.style.h) === 595);

console.log('Alerts shown:\n  ' + alerts.join('\n  ').replace(/\n(?! )/g, ' / '));
console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
