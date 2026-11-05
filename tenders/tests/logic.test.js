// Tests the parts of tenders/Code.gs that do not need Google.
// Run: node tenders/tests/logic.test.js
const fs = require('fs');
const vm = require('vm');
const path = require('path');

function formatDate(d, tz, pattern) {
  const p = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
    .formatToParts(d).forEach((x) => { p[x.type] = x.value; });
  const mm = String(new Date(d.toLocaleString('en-US', { timeZone: tz })).getMonth() + 1).padStart(2, '0');
  return pattern.replace('yyyy', p.year).replace('MMMM', p.month).replace('MMM', p.month.slice(0, 3)).replace('MM', mm)
    .replace('dd', p.day.padStart(2, '0')).replace(/\bd\b/, p.day).replace('HH', p.hour).replace('mm', p.minute);
}
const ctx = { Utilities: { formatDate } };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx);

let pass = 0, fail = 0;
function eq(name, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (ok) pass++; else { fail++; console.log('FAIL', name, '\n  got ', JSON.stringify(got), '\n  want', JSON.stringify(want)); }
}
const key = (d) => (d ? formatDate(d, 'Africa/Kampala', 'yyyy-MM-dd') : null);

// Dates typed by people, read as Kampala calendar dates
eq('iso', key(ctx.parseDate_('2026-11-20')), '2026-11-20');
eq('day first', key(ctx.parseDate_('5/11/2026')), '2026-11-05');
eq('day first dots', key(ctx.parseDate_('05.11.2026')), '2026-11-05');
eq('text month', key(ctx.parseDate_('5 Nov 2026')), '2026-11-05');
eq('long month', key(ctx.parseDate_('20 November 2026')), '2026-11-20');
eq('bad date', ctx.parseDate_('31/02/2026'), null);
eq('nonsense', ctx.parseDate_('next week'), null);
eq('kampala midnight', ctx.kampalaDate_(2026, 11, 5).toISOString(), '2026-11-04T21:00:00.000Z');
eq('days until', ctx.daysUntil_(ctx.kampalaDate_(2026, 11, 12), ctx.kampalaDate_(2026, 11, 5)), 7);
eq('days until passed', ctx.daysUntil_(ctx.kampalaDate_(2026, 11, 1), ctx.kampalaDate_(2026, 11, 5)), -4);
eq('23:30 Kampala on the 5th: deadline on the 6th is tomorrow', ctx.daysUntil_(ctx.kampalaDate_(2026, 11, 6), new Date('2026-11-05T20:30:00Z')), 1);
eq('01:30 Kampala on the 6th (still the 5th in UTC): deadline is today', ctx.daysUntil_(ctx.kampalaDate_(2026, 11, 6), new Date('2026-11-05T22:30:00Z')), 0);

// Reference numbers
const used = { 'TND-2026-0001': true, 'TND-2026-0009': true, 'TND-2025-0040': true };
eq('next ref', ctx.nextRefId_('2026', used), 'TND-2026-0010');
eq('next ref again', ctx.nextRefId_('2026', used), 'TND-2026-0011');
eq('new year', ctx.nextRefId_('2027', used), 'TND-2027-0001');
eq('column letters', [1, 17, 18, 26, 27, 52].map(ctx.colLetter_), ['A', 'Q', 'R', 'Z', 'AA', 'AZ']);

// Splitting the required documents list
eq('split commas', ctx.splitRequired_('Certificate of incorporation, Tax clearance, CVs'), ['Certificate of incorporation', 'Tax clearance', 'CVs']);
eq('split lines and bullets', ctx.splitRequired_('- Company profile\n2. Audited accounts\n• Bank letter'), ['Company profile', 'Audited accounts', 'Bank letter']);
eq('keep commas in brackets', ctx.splitRequired_('Audited accounts (2024, 2025), Trading licence'), ['Audited accounts (2024, 2025)', 'Trading licence']);
eq('short title', ctx.shortTitle_('Provision of consultancy services for the end-line evaluation of a youth employment programme'), 'Provision of consultancy services for the');

// Matching notice wording to standing documents
const docs = ctx.STANDING_DOCUMENTS.map((d) => ({ document: d, status: 'Have', drive_link: '' })).concat([{ document: 'ISO 9001 certificate', status: 'Missing' }]);
const m = (s) => { const d = ctx.matchDocument_(s, docs); return d ? d.document : null; };
eq('incorporation', m('Certified copy of certificate of incorporation'), 'Certificate of Incorporation');
eq('tax clearance', m('Valid tax clearance certificate'), 'URA tax clearance certificate');
eq('tin', m('TIN'), 'URA TIN certificate');
eq('cvs', m('CVs of trainers'), 'CVs of key experts');
eq('audited', m('Audited accounts for the last 2 years'), 'Audited financial statements (last 1-3 years)');
eq('experience', m('Three similar assignments with completion certificates'), 'Past contracts, LPOs and completion certificates');
eq('poa', m('Power of attorney'), 'Power of attorney for bid signing');
eq('bribery', m('Anti-bribery declaration'), 'Bid declaration and anti-bribery declaration templates');
eq('nssf', m('NSSF compliance certificate'), 'NSSF registration and compliance certificate');
eq('ppda', m('PPDA registration'), 'PPDA supplier registration (if applicable)');
eq('bank', m('Bank details letter'), 'Bank reference letter / bank details letter');
eq('own document by words', m('ISO 9001 quality certificate'), 'ISO 9001 certificate');
eq('unknown stays unmatched', m('Signed bid form'), null);
eq('memart', m('MEMARTS'), 'Memorandum and Articles of Association');

// Reminder selection
const today = ctx.kampalaDate_(2026, 11, 5);
const plus = (n) => new Date(today.getTime() + n * 86400000);
const tenders = [
  { ref_id: 'A', deadline: plus(2), status: 'Preparing' },
  { ref_id: 'B', deadline: plus(7), status: 'Watching' },
  { ref_id: 'C', deadline: plus(8), status: 'Watching' },
  { ref_id: 'D', deadline: plus(1), status: 'Submitted' },
  { ref_id: 'E', deadline: plus(-1), status: 'Preparing' },
  { ref_id: 'F', deadline: plus(0), status: '' },
  { ref_id: 'G', deadline: '', status: 'Watching' },
  { ref_id: 'H', deadline: plus(3), status: 'Shortlisted' }
];
const docRows = [
  { document: 'Tax clearance', status: 'Have', expiry_date: plus(20) },
  { document: 'Trading licence', status: 'Have', expiry_date: plus(45) },
  { document: 'NSSF', status: 'Have', expiry_date: plus(-3) },
  { document: 'Profile', status: 'Expired', expiry_date: '' },
  { document: 'CVs', status: 'Missing', expiry_date: '' }
];
const items = ctx.reminderItems_(tenders, docRows, today);
eq('reminder tenders', items.tenders.map((t) => t.ref_id), ['F', 'A', 'B']);
eq('reminder docs', items.docs.map((d) => d.document), ['Tax clearance', 'NSSF', 'Profile']);
eq('doc state expired by date', ctx.docState_({ status: 'Have', expiry_date: plus(-1) }, today), 'Expired');
eq('doc state unchecked', ctx.docState_({ status: '' }, today), 'Not checked');

// Formulas use the real column letters and the right sheet names
const cols = {}; ctx.COLUMNS.forEach((c, i) => { cols[c] = i; });
const dcols = {}; ctx.DOC_COLUMNS.forEach((c, i) => { dcols[c] = i; });
const f = ctx.dashboardFormulas_(cols, dcols);
const all = [].concat(f.kpis.map((k) => k[1]), f.byStatus.map((k) => k[1]), f.byCategory.flatMap((k) => [k[1], k[2]]), [f.due7, f.due30, f.docs]);
eq('every formula starts with =', all.every((x) => x.startsWith('=')), true);
eq('balanced brackets', all.filter((x) => (x.match(/\(/g) || []).length !== (x.match(/\)/g) || []).length || (x.match(/\{/g) || []).length !== (x.match(/\}/g) || []).length), []);
eq('balanced quotes', all.filter((x) => (x.match(/"/g) || []).length % 2), []);
eq('status column J', f.byStatus[0][1], "=COUNTIF('Tenders'!J2:J,\"Watching\")");
eq('due 7 formula', f.due7, "=IFERROR(SORT(FILTER({'Tenders'!A2:A,'Tenders'!B2:B,'Tenders'!C2:C,'Tenders'!H2:H},ISNUMBER('Tenders'!H2:H),'Tenders'!H2:H>=TODAY(),'Tenders'!H2:H<=TODAY()+7,ISNA(MATCH('Tenders'!J2:J,{\"Submitted\",\"Shortlisted\",\"Won\",\"Lost\",\"Not pursued\",\"Cancelled\"},0))),4,TRUE),\"None\")");
eq('docs formula', f.docs, "=IFERROR(FILTER({'Documents'!A2:A,'Documents'!B2:B,'Documents'!D2:D},LEN('Documents'!A2:A),('Documents'!B2:B=\"Missing\")+('Documents'!B2:B=\"Expired\")+ISNUMBER('Documents'!D2:D)*('Documents'!D2:D<=TODAY()+30)),\"None\")");
eq('submitted uses R (submitted_date)', f.kpis[3][1].includes("'Tenders'!R2:R"), true);
eq('win rate', f.kpis[6][1], "=IFERROR(SUMPRODUCT(--(('Tenders'!J2:J=\"Won\")+('Tenders'!P2:P=\"Won\")>0))/(SUMPRODUCT(--(('Tenders'!J2:J=\"Won\")+('Tenders'!P2:P=\"Won\")>0))+SUMPRODUCT(--(('Tenders'!J2:J=\"Lost\")+('Tenders'!P2:P=\"Lost\")>0))),\"–\")");
eq('colour rules', ctx.tenderColourFormulas_('$H2').map((r) => r[0]), ['$H2<TODAY()', '$H2-TODAY()<=3', '$H2-TODAY()<=7', '$H2-TODAY()<=14']);

console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
