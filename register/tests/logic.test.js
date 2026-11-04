// Tests the parts of Code.gs that do not need Google: results, IDs, dates, summaries,
// email text and the verification lookup (with a fake spreadsheet).
// Run: node register/tests/logic.test.js
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const TZ = 'Africa/Kampala';
function formatDate(d, tz, pattern) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: 'long', day: 'numeric' })
    .formatToParts(d).forEach((p) => { parts[p.type] = p.value; });
  const mm = String(new Date(d.toLocaleString('en-US', { timeZone: tz })).getMonth() + 1).padStart(2, '0');
  return pattern
    .replace('yyyy', parts.year).replace('MMMM', parts.month).replace('MMM', parts.month.slice(0, 3))
    .replace('MM', mm).replace('dd', parts.day.padStart(2, '0')).replace(/\bd\b/, parts.day);
}

// Fake spreadsheet for lookupCertificate_
const header = ['certificate_id', 'participant_name', 'organisation', 'email', 'phone', 'course_code', 'course_title', 'start_date', 'end_date', 'location', 'format', 'attendance_percent', 'pre_test_score', 'post_test_score', 'result', 'certificate_type', 'issued_date', 'status', 'notes'];
const D = (s) => new Date(s + 'T09:00:00+03:00');
const data = [
  ['TSV-2027-T1.1-0001', 'Amina Nakato', 'Secret Org', 'amina@example.com', '+256700000001', 'T1.1', 'Practical Monitoring, Evaluation and Learning for Programmes', D('2027-03-10'), D('2027-03-14'), 'Kampala', 'Open enrolment', 100, 45, 82, 'Pass', 'Competence', D('2027-03-15'), 'Issued', 'private note'],
  ['TSV-2027-T1.1-0002', 'Brian Okello', 'Org B', 'brian@example.com', '', 'T1.1', 'Practical Monitoring, Evaluation and Learning for Programmes', D('2027-03-10'), D('2027-03-14'), 'Kampala', 'Open enrolment', 90, 50, 68, 'Pass', 'Completion', D('2027-03-15'), 'Revoked', ''],
  ['TSV-2027-T1.1-0003', 'Carol Achen', 'Org C', '', '', 'T1.1', 'Practical Monitoring, Evaluation and Learning for Programmes', D('2027-03-10'), D('2027-03-14'), 'Kampala', 'Open enrolment', 95, 40, 70, 'Pass', 'Completion', '', 'Draft', '']
];
const fakeSheet = {
  getLastRow: () => data.length + 1,
  getLastColumn: () => header.length,
  getRange: (r, c, nr, nc) => ({ getValues: () => (r === 1 ? [header] : data.slice(r - 2, r - 2 + nr)) })
};
const fakeSS = { getSheetByName: (n) => (n === 'Register' ? fakeSheet : null), getSpreadsheetTimeZone: () => TZ };

const ctx = {
  Utilities: { formatDate },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  SpreadsheetApp: { getActiveSpreadsheet: () => fakeSS, openById: () => fakeSS }
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx);

let pass = 0, fail = 0;
function eq(name, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (ok) pass++; else { fail++; console.log('FAIL', name, '\n  got ', JSON.stringify(got), '\n  want', JSON.stringify(want)); }
}

const S = { pass_attendance_percent: 80, pass_post_test_score: 60, competence_post_test_score: 75, hours_per_day: 7 };
const fmt = (d, p) => formatDate(d, TZ, p);

// Results
eq('pass competence', ctx.computeResult_(100, 82, S), { result: 'Pass', certificate_type: 'Competence' });
eq('pass completion', ctx.computeResult_(90, 68, S), { result: 'Pass', certificate_type: 'Completion' });
eq('boundary pass', ctx.computeResult_(80, 60, S), { result: 'Pass', certificate_type: 'Completion' });
eq('boundary competence', ctx.computeResult_(80, 75, S), { result: 'Pass', certificate_type: 'Competence' });
eq('low attendance', ctx.computeResult_(60, 90, S), { result: 'Did not complete', certificate_type: '' });
eq('low score', ctx.computeResult_(95, 55, S), { result: 'Did not complete', certificate_type: '' });
eq('missing data', ctx.computeResult_(null, 70, S), null);
eq('percent as fraction', ctx.toNumber_(0.85, true), 85);
eq('percent text', ctx.toNumber_('85%', true), 85);
eq('blank', ctx.toNumber_('', true), null);
eq('score not scaled', ctx.toNumber_(1, false), 1);

// Improvement note keeps other notes and replaces an old improvement
eq('improvement new', ctx.improvementNote_('', 45, 82), 'Improvement: +37 points (pre 45, post 82)');
eq('improvement replace', ctx.improvementNote_('Needs follow-up | Improvement: +5 points (pre 1, post 6)', 50, 68), 'Needs follow-up | Improvement: +18 points (pre 50, post 68)');
eq('improvement negative', ctx.improvementNote_('x', 70, 65), 'x | Improvement: -5 points (pre 70, post 65)');
eq('improvement no pre', ctx.improvementNote_('keep me', null, 65), 'keep me');

// Certificate IDs
const used = { 'TSV-2026-T1.1-0001': true, 'TSV-2026-T1.1-0007': true, 'TSV-2026-T2.1-0003': true };
eq('next id', ctx.nextCertificateId_('2026', 't1.1 ', used), 'TSV-2026-T1.1-0008');
eq('next id again', ctx.nextCertificateId_('2026', 'T1.1', used), 'TSV-2026-T1.1-0009');
eq('new course', ctx.nextCertificateId_('2026', 'T8.5', used), 'TSV-2026-T8.5-0001');
eq('new year', ctx.nextCertificateId_('2027', 'T1.1', used), 'TSV-2027-T1.1-0001');
eq('beyond 9999', ctx.nextCertificateId_('2026', 'T9.9', { 'TSV-2026-T9.9-9999': true }), 'TSV-2026-T9.9-10000');

// Dates
eq('same month', ctx.formatDateRange_(D('2027-03-10'), D('2027-03-14'), fmt), '10–14 March 2027');
eq('two months', ctx.formatDateRange_(D('2027-02-28'), D('2027-03-03'), fmt), '28 February – 3 March 2027');
eq('two years', ctx.formatDateRange_(D('2026-12-30'), D('2027-01-02'), fmt), '30 December 2026 – 2 January 2027');
eq('one day', ctx.formatDateRange_(D('2027-03-10'), D('2027-03-10'), fmt), '10 March 2027');
eq('no end', ctx.formatDateRange_(D('2027-03-10'), '', fmt), '10 March 2027');
eq('no dates', ctx.formatDateRange_('', '', fmt), '');
eq('days from list', ctx.courseDays_({ days: 5 }, D('2027-03-10'), D('2027-03-12')), 5);
eq('days from dates', ctx.courseDays_(null, D('2027-03-10'), D('2027-03-14')), 5);

// Summary statistics
const rows = data.map((v) => { const o = {}; header.forEach((h, i) => { o[h] = v[i]; }); return o; });
rows.push({ course_code: 'T1.1', participant_name: 'Dan', organisation: 'Org B', attendance_percent: 60, pre_test_score: 30, post_test_score: 50, result: 'Did not complete', status: 'Draft', sex: 'Male' });
rows[0].sex = 'Female'; rows[1].sex = 'Male'; rows[2].sex = 'Female';
const cols = {}; header.concat(['sex']).forEach((h, i) => { cols[h] = i; });
const st = ctx.courseStats_(rows, cols);
eq('participants', st.participants, 4);
eq('pass/fail', [st.pass, st.fail], [3, 1]);
eq('pass rate', st.passRate, 0.75);
eq('avg attendance', st.avgAttendance, 86.3);
eq('avg pre/post', [st.avgPre, st.avgPost], [41.3, 67.5]);
eq('avg improvement', st.avgImprovement, 26.3);
eq('issued/competence/revoked', [st.issued, st.competence, st.revoked], [1, 1, 1]);
eq('by sex', st.bySex, { Female: 2, Male: 2 });
eq('by org', st.byOrg, { 'Secret Org': 1, 'Org B': 2, 'Org C': 1 });
const runs = ctx.courseRuns_(rows.concat([{ course_code: 'T2.1', start_date: D('2027-05-01') }]), (d) => (d instanceof Date ? fmt(d, 'yyyy-MM-dd') : ''));
eq('runs newest first', runs.map((r) => r.code + ' ' + r.start), ['T2.1 2027-05-01', 'T1.1 2027-03-10', 'T1.1 ']);

// Email text
const em = ctx.certificateEmail_({ name: 'Amina Nakato', title: 'Practical MEL', dates: '10–14 March 2027', location: 'Kampala', type: 'Competence', id: 'TSV-2027-T1.1-0001', verifyUrl: 'https://example.com/exec?id=TSV-2027-T1.1-0001' });
eq('email subject', em.subject, 'Your Tsavone certificate: Practical MEL');
eq('email greeting', em.text.split('\n')[0], 'Dear Amina,');
eq('email has link', em.text.includes('https://example.com/exec?id=TSV-2027-T1.1-0001'), true);
eq('email html escapes', ctx.certificateEmail_({ name: '<b>X', title: 'T', dates: '', location: '', type: 'Completion', id: 'I', verifyUrl: '' }).html.includes('&lt;b&gt;X'), true);
eq('email no link when unset', ctx.certificateEmail_({ name: 'A', title: 'T', dates: '', location: '', type: 'Completion', id: 'I', verifyUrl: '?id=I' }).text.includes('confirm it'), false);

// Verification lookup: only public fields, correct states
const v = ctx.lookupCertificate_('tsv-2027-t1.1-0001');
eq('valid state', v.state, 'VALID');
eq('valid fields only', Object.keys(v).sort(), ['course', 'dates', 'id', 'name', 'state', 'type']);
eq('valid values', [v.name, v.dates, v.type], ['Amina Nakato', '10–14 March 2027', 'Certificate of Competence']);
const leaked = JSON.stringify(v);
eq('no private data', ['Secret Org', 'amina@', '+2567', 'private note', '45', '82'].filter((s) => leaked.includes(s)), []);
eq('revoked', ctx.lookupCertificate_('TSV-2027-T1.1-0002'), { state: 'REVOKED', id: 'TSV-2027-T1.1-0002' });
eq('draft hidden', ctx.lookupCertificate_('TSV-2027-T1.1-0003').state, 'NOT_FOUND');
eq('unknown', ctx.lookupCertificate_('TSV-1999-X-0001').state, 'NOT_FOUND');

console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
