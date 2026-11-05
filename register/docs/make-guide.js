// Builds the Tsavone Training Register operations guide (.docx).
// Run: node register/docs/make-guide.js register/docs/Tsavone-Training-Register-Operations-Guide.docx register/tsavone-logo.png
// Needs the docx npm package. Convert to PDF with LibreOffice.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, LevelFormat, BorderStyle, Footer, Header, PageNumber, PageBreak,
  TableOfContents, ExternalHyperlink, VerticalAlign, LineRuleType
} = require('docx');

const OUT = process.argv[2];
const LOGO = process.argv[3];
const NAVY = '193591', ORANGE = 'F17B19', ORANGE_DARK = 'B04E00', SKY = 'EFFCFF', INK = '15204A', MUTED = '5A6385', LINE = 'DBE3F0';
const FONT = 'Arial';
const W = 9638; // A4 width minus 2 cm margins each side, in DXA

const SHEET = 'https://docs.google.com/spreadsheets/d/1B3MdLQuxSe1q2pwYIj6iUM-YVUosEEoKqYTfySsEsCw/edit';
const VERIFY = 'https://script.google.com/macros/s/AKfycbwV6gNlu3byQcLeGEd--6XVK4cYH-ceS-qnuvizbm0I7jiH2XnipLkEGkoMutA5QVo/exec';
const CODE_RAW = 'https://raw.githubusercontent.com/pxengoru-bot/tsavone-tools1/main/register/Code.gs';
const VERIFY_RAW = 'https://raw.githubusercontent.com/pxengoru-bot/tsavone-tools1/main/register/Verify.html';

/* ---------- helpers ---------- */
const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
function runs(spec) {
  // spec: string with **bold** segments
  const out = [];
  String(spec).split(/(\*\*[^*]+\*\*)/).forEach((part) => {
    if (!part) return;
    if (part.startsWith('**')) out.push(t(part.slice(2, -2), { bold: true }));
    else out.push(t(part));
  });
  return out;
}
const p = (spec, o = {}) => new Paragraph({ children: Array.isArray(spec) ? spec : runs(spec), spacing: { after: 120 }, ...o });
const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(text)] });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(text)] });
const bullet = (spec, level = 0) => new Paragraph({ numbering: { reference: 'bullets', level }, children: runs(spec), spacing: { after: 60 } });
let numRef = 0;
function steps(list) {
  const ref = 'steps' + (numRef++);
  numberingConfigs.push({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 440, hanging: 360 } }, run: { bold: true, color: NAVY } } }] });
  return list.map((s) => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(s), spacing: { after: 80 } }));
}
const link = (text, url) => new ExternalHyperlink({ link: url, children: [t(text, { color: NAVY, underline: {} })] });

const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, rows, opts = {}) {
  const head = opts.head !== false;
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: head && ri === 0,
      cantSplit: true,
      children: r.map((c, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA },
        borders,
        shading: head && ri === 0 ? { fill: NAVY, type: ShadingType.CLEAR, color: 'auto' } : (ri % 2 === 0 ? { fill: 'F6F9FD', type: ShadingType.CLEAR, color: 'auto' } : undefined),
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.TOP,
        children: (Array.isArray(c) ? c : [c]).map((cc) => cc instanceof Paragraph ? cc : new Paragraph({
          children: head && ri === 0 ? [t(String(cc), { bold: true, color: 'FFFFFF' })] : runs(cc),
          spacing: { after: 40 }
        }))
      }))
    }))
  });
}
function callout(spec, fill = SKY, edge = ORANGE) {
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR, color: 'auto' },
      borders: { top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: BorderStyle.SINGLE, size: 24, color: edge } },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: (Array.isArray(spec) ? spec : [spec]).map((s) => new Paragraph({ children: runs(s), spacing: { after: 60 } }))
    })] })]
  });
}
const gap = () => new Paragraph({ children: [], spacing: { after: 120 } });

const numberingConfigs = [
  { reference: 'bullets', levels: [
    { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 440, hanging: 300 } } } },
    { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 880, hanging: 300 } } } }
  ] }
];

/* ---------- content ---------- */
const body = [];

// Cover
body.push(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1400, after: 300 }, children: [new ImageRun({ type: 'png', data: fs.readFileSync(LOGO), transformation: { width: 170, height: 152 }, altText: { title: 'Tsavone Consulting', description: 'Tsavone Consulting logo', name: 'logo' } })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [t('TRAINING AND CAPACITY DEVELOPMENT', { bold: true, color: ORANGE_DARK, size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 120, line: 240, lineRule: LineRuleType.AUTO }, children: [t('Training Register', { bold: true, color: NAVY, size: 64 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [t('Operations Guide', { bold: true, color: INK, size: 40 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [t('How to record participants, issue certificates, send them and report on a course', { color: MUTED, size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1600, after: 60 }, children: [t('Version 1.0  ·  November 2026', { color: MUTED, size: 20 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [t('Tsavone Consulting  |  www.tsavone.com  |  info@tsavone.com', { color: MUTED, size: 20 })] }),
  new Paragraph({ children: [new PageBreak()] })
);

// Contents
body.push(
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t('Contents')] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('1. What the Training Register does', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('2. Where and how you use it', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('3. The course cycle, step by step', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 440 }, children: [t('A. Before the course: enter participants', { size: 22, bold: false, color: MUTED })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 440 }, children: [t('B. After the course: scores and results', { size: 22, bold: false, color: MUTED })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 440 }, children: [t('C. Make the certificates', { size: 22, bold: false, color: MUTED })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 440 }, children: [t('D. Email the certificates', { size: 22, bold: false, color: MUTED })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 440 }, children: [t('E. Report on the course', { size: 22, bold: false, color: MUTED })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('4. Corrections, reissues and revoking', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('5. The verification page', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('6. Settings', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('7. Do and do not', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('8. Trying things out safely', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('9. If something goes wrong', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ spacing: { after: 80 }, indent: { left: 0 }, children: [t('10. Updating the program (only when asked)', { size: 24, bold: true, color: INK })] }),
  new Paragraph({ children: [new PageBreak()] })
);

// 1
body.push(
  h1('1. What the Training Register does'),
  p('The Training Register is a Google Sheet with a **Tsavone** menu. From that menu you can:'),
  bullet('fill in course titles from course codes'),
  bullet('work out who passed and which certificate each person earns'),
  bullet('give every certificate a unique ID, such as **TSV-2027-T1.1-0001**'),
  bullet('make A4 certificates as PDF files, each with a QR code'),
  bullet('prepare certificate emails as Gmail drafts for you to check and send'),
  bullet('build a course summary for your completion report'),
  p('Anyone can check a certificate on the public **verification page**, by scanning its QR code or typing its ID.', { spacing: { before: 120, after: 120 } }),
  h2('Where everything lives'),
  table([2600, 7038], [
    ['Item', 'Where to find it'],
    ['The register (use this every time)', [new Paragraph({ children: [t('Google Sheet "Tsavone Training Register": '), link('open the sheet', SHEET)] })]],
    ['Certificate PDFs', 'Google Drive → My Drive → **Tsavone Certificates** → one folder per course code (for example T1.1)'],
    ['Certificate emails', 'Gmail → **Drafts** (they are never sent until you press Send)'],
    ['Verification page (public)', [new Paragraph({ children: [link(VERIFY, VERIFY)] })]],
    ['Course list (titles and days)', 'Your public Google Sheet "Tsavone Courses"'],
    ['Logo used on certificates', 'Google Drive → "Tsavone logo for certificates.png" (do not delete)'],
    ['Rules and settings', 'The **Settings** tab in the register']
  ]),
  gap(),
  callout('**Tip:** bookmark the register in your browser so it is always one click away.')
);

// 2
body.push(
  h1('2. Where and how you use it'),
  p('You do **not** need to install or upload anything again. The program is stored inside the register sheet itself. Each time you use it:'),
  ...steps([
    'Open the **Tsavone Training Register** sheet in a web browser (Chrome is best) on a **computer**.',
    'Make sure you are signed in to the Google account that owns the sheet.',
    'Wait a few seconds until the **Tsavone** menu appears at the top, after Help.',
    'Use the Tsavone menu items described in this guide.'
  ]),
  callout([
    '**Good to know**',
    'The Tsavone menu does not appear in the Google Sheets app on phones and tablets. Use a computer browser to run the menu.',
    'The verification page works on any device and needs no sign-in.',
    'You can open the sheet from Google Drive, from your bookmark or from the link above. They all open the same file.'
  ]),
  gap(),
  h2('Colleagues using the register'),
  p('To let a colleague use the register, share the sheet with them as an **Editor** (Share button, top right). The first time they use the Tsavone menu, Google asks them to approve the same permissions you approved. Certificates they generate are saved in **their** Google Drive, and drafts go to **their** Gmail. Decide who generates certificates so that all PDFs end up in one place.')
);

// 3
body.push(
  h1('3. The course cycle, step by step'),
  h2('A. Before the course: enter participants'),
  p('On the **Register** tab, add one row per participant, starting in the first empty row. Fill in only the white columns in the table below; the tool fills the rest.'),
  table([2500, 3700, 3438], [
    ['Column', 'You type', 'Example'],
    ['participant_name', 'Full name as it should appear on the certificate', 'Amina Nakato'],
    ['organisation', 'Employer or sponsor', 'Gulu District Local Government'],
    ['email', 'Address for the certificate email', 'amina@example.org'],
    ['phone', 'Optional', '+256 700 000 000'],
    ['course_code', 'Code from the catalogue', 'T2.1'],
    ['start_date, end_date', 'Course dates', '10 Mar 2027 and 14 Mar 2027'],
    ['location', 'Where it ran', 'Kampala, Uganda'],
    ['format', 'Choose from the dropdown', 'Open enrolment'],
    ['status', 'Choose **Draft**', 'Draft'],
    ['notes', 'Optional notes', 'Sponsored by project X']
  ]),
  gap(),
  p('Then click **Tsavone → Fill course titles from codes**. The course_title column fills in from your course list. Titles you typed yourself are kept.'),
  callout('**Leave empty:** certificate_id, course_title, result, certificate_type and issued_date. The tool fills them.'),
  gap(),

  h2('B. After the course: scores and results'),
  ...steps([
    'Enter **attendance_percent** (for example 95), **pre_test_score** and **post_test_score** for each person.',
    'Click **Tsavone → Calculate results**.',
    'Check the **result** column: **Pass** or **Did not complete**. The **certificate_type** column shows **Competence** or **Completion**, and **notes** shows the improvement, for example "Improvement: +23 points".'
  ]),
  table([3200, 6438], [
    ['Rule (change in Settings)', 'Meaning'],
    ['Attendance at least **80%** and post-test at least **60**', '**Pass**: the participant gets a certificate'],
    ['Pass with post-test at least **75**', '**Certificate of Competence**'],
    ['Pass with post-test below 75', '**Certificate of Completion**'],
    ['Anything else', '**Did not complete**: no certificate']
  ]),
  gap(),
  p('Rows that are already **Issued** or **Revoked** are never changed by Calculate results.'),

  h2('C. Make the certificates'),
  ...steps([
    'Click **Tsavone → Generate certificates**.',
    'It asks: "Make certificates for the N passing participant(s) who do not have one yet?" Click **OK**.',
    'Wait. Each certificate takes a few seconds. A message confirms how many were saved.',
    'Open **Google Drive → Tsavone Certificates → [course code]** and **check the first PDF** before sending anything.'
  ]),
  p('Each row that receives a certificate gets a **certificate_id** (for example TSV-2027-T2.1-0001), today\'s **issued_date** and status **Issued**.'),
  callout('**Large courses:** Google limits each run to 6 minutes, which is usually enough for several dozen certificates. If the message says it stopped early, run Generate certificates again; it continues where it stopped.'),
  gap(),

  h2('D. Email the certificates'),
  ...steps([
    'On the left edge of the sheet, click the **grey row number** of the first person to email. Hold **Shift** and click the grey row number of the last person. The rows turn blue.',
    'Click **Tsavone → Email certificates (create Gmail drafts)**.',
    'Open **Gmail → Drafts**. Each draft has the PDF attached, a short congratulations message, the certificate ID and the verification link.',
    'Check each draft and press **Send**.'
  ]),
  p('Rows that are not Issued, or have no valid email, are skipped and listed in the message.'),

  h2('E. Report on the course'),
  ...steps([
    'Click **Tsavone → Course summary**.',
    'A numbered list of courses appears (course code and start date). Type the number of the course and click **OK**.',
    'A new tab opens, for example "Summary T2.1 2027-03-10".'
  ]),
  p('The summary shows participants, passes, pass rate, average attendance, average pre-test, post-test and improvement, and certificates issued. It also has counts by organisation, and a ready-made sentence for your completion report. Copy the table and sentence into the report. Running the summary again replaces the tab with fresh numbers.'),
  callout('**Optional:** add a column headed **sex** (or **gender**) to the Register, and the summary also counts participants by sex.')
);

// 4
body.push(
  h1('4. Corrections, reissues and revoking'),
  table([3000, 6638], [
    ['Situation', 'What to do'],
    ['A name or detail on an issued certificate is wrong', 'Correct the row. Click its **grey row number**, then **Tsavone → Generate certificates** and answer **YES** ("make or remake certificates for just these rows"). The old PDF is replaced and the original issue date is kept. Then email a new draft.'],
    ['A certificate must be withdrawn', 'Set its **status** to **Revoked**. The verification page immediately shows REVOKED for that ID.'],
    ['A participant was entered by mistake', 'If no certificate was issued, delete the row. If one was issued, set the status to **Revoked** instead, so the ID is never reused.'],
    ['Scores were entered wrongly before issuing', 'Fix the scores and run **Calculate results** again.'],
    ['Scores were wrong after issuing', 'Set status back to **Draft**, fix the scores, run Calculate results, then remake the certificate as above.']
  ])
);

// 5
body.push(
  h1('5. The verification page'),
  p('Every certificate carries a QR code and its ID. Anyone, such as an employer, a donor or the participant, can check it:'),
  bullet('scan the QR code with a phone camera, or'),
  bullet('open the verification page (link in section 1) and type the certificate ID.'),
  table([2400, 7238], [
    ['Page shows', 'When'],
    ['**VALID** (green)', 'Status is Issued. Shows name, course, dates and certificate type.'],
    ['**REVOKED** (red)', 'Status is Revoked.'],
    ['**NOT FOUND** (orange)', 'The ID does not exist, or the row is still Draft.']
  ]),
  gap(),
  callout('**Privacy:** the page never shows email, phone, organisation, scores or notes. It does show the name to anyone who has the ID. Because IDs follow a sequence, someone could guess other IDs. For a sensitive course, keep rows in Draft status; they then show NOT FOUND.')
);

// 6
body.push(
  h1('6. Settings'),
  p('Open the **Settings** tab and change only column B (value).'),
  table([3000, 1800, 4838], [
    ['Setting', 'Current value', 'What it controls'],
    ['pass_attendance_percent', '80', 'Minimum attendance (%) to pass'],
    ['pass_post_test_score', '60', 'Minimum post-test score to pass'],
    ['competence_post_test_score', '75', 'Score for a Certificate of Competence'],
    ['certificate_folder_name', 'Tsavone Certificates', 'Drive folder for PDFs'],
    ['logo_drive_file_id', '(filled in)', 'Logo on certificates and the verification page'],
    ['signatory_title', 'Managing Director', 'Title under the signature line'],
    ['sender_email_name', 'Tsavone Consulting Training', 'Sender name on emails'],
    ['verify_url', '(your /exec link)', 'Where QR codes point. **Do not change.**'],
    ['course_sheet_id', '(your course sheet)', 'Source of course titles and days'],
    ['hours_per_day', '7', 'Course hours = days x this number']
  ])
);

// 7
body.push(
  h1('7. Do and do not'),
  table([4819, 4819], [
    ['Do', 'Do not'],
    [
      ['Check the first PDF of every course before sending.', 'Use the Tsavone menu from a computer browser.', 'Keep status values to Draft, Issued or Revoked.', 'Revoke instead of deleting an issued certificate.', 'Ask colleagues to use the same account, or agree who issues certificates.'].map((s) => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, children: runs(s), spacing: { after: 40 } })),
      ['Rename or reorder the column headings in row 1.', 'Delete the **Settings** tab or the logo file in Drive.', 'Create a **new** deployment of the verification page (QR codes would break).', 'Send out test certificates.', 'Change verify_url.'].map((s) => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, children: runs(s), spacing: { after: 40 } }))
    ]
  ])
);

// 8
body.push(
  h1('8. Trying things out safely'),
  ...steps([
    'Click **Tsavone → Testing → Add test participants**. Three test rows appear, using your own email address with +test1, +test2 and +test3.',
    'Try any menu item on them.',
    'When finished, click **Tsavone → Testing → Remove test participants** and confirm. The test rows, their PDFs, their drafts and their summary tab are removed. Real participants are never touched.'
  ])
);

// 9
body.push(
  h1('9. If something goes wrong'),
  table([3400, 6238], [
    ['Problem', 'What to do'],
    ['No Tsavone menu', 'Reload the page (F5) and wait 10 seconds. Make sure you are on a computer, signed in to the owner account.'],
    ['Google asks for permission', 'Click **Review permissions**, choose your account, then **Advanced**, **Go to Tsavone Training Register**, **Select all** and **Continue**. This happens once per new feature.'],
    ['"Codes not found in the course list"', 'Check the course code spelling (for example T1.1) against the Tsavone Courses sheet.'],
    ['"Missing attendance or post-test score"', 'Fill in those cells and run Calculate results again.'],
    ['A certificate was not made', 'The person must show **Pass** in the result column. Run Calculate results first.'],
    ['Wrong date on a certificate', 'Check the start and end dates in the row, then remake the certificate.'],
    ['Logo missing', 'Clear Settings → logo_drive_file_id and click **Tsavone → Run setup again**.'],
    ['Anything else', 'Take a screenshot of the message and the sheet, and send it to your support contact.']
  ])
);

// 10
body.push(
  h1('10. Updating the program (only when asked)'),
  p('You only need this if an improved version of the program is provided.'),
  ...steps([
    'In the sheet, click **Extensions → Apps Script**.',
    'Open the latest **Code.gs** (link below). Press **Ctrl+A**, then **Ctrl+C**. In Apps Script, click **Code.gs**, click in the code, then press **Ctrl+A**, **Ctrl+V** and **Ctrl+S**.',
    'Do the same for **Verify.html**, if it changed.',
    'Click **Deploy → Manage deployments**, then the pencil icon. Set **Version: New version** and click **Deploy**. This keeps the same verification link.'
  ]),
  new Paragraph({ spacing: { before: 120 }, children: [t('Latest files: '), link('Code.gs', CODE_RAW), t('  ·  '), link('Verify.html', VERIFY_RAW)] })
);

/* ---------- document ---------- */
const doc = new Document({
  creator: 'Tsavone Consulting',
  title: 'Training Register – Operations Guide',
  description: 'How to use the Tsavone Training Register',
  styles: {
    default: { document: { run: { font: FONT, size: 22, color: '2C3557' }, paragraph: { spacing: { line: 276, lineRule: LineRuleType.AUTO } } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 34, bold: true, color: NAVY },
        paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, keepLines: true, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 26, bold: true, color: INK },
        paragraph: { spacing: { before: 280, after: 100 }, keepNext: true, keepLines: true, outlineLevel: 1 } }
    ]
  },
  numbering: { config: numberingConfigs },
  features: { updateFields: true },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } }, titlePage: true },
    headers: {
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [t('Tsavone Consulting  ·  Training Register – Operations Guide', { size: 16, color: MUTED })] })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [t('Page ', { size: 16, color: MUTED }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: MUTED }), t('  ·  www.tsavone.com  ·  info@tsavone.com', { size: 16, color: MUTED })] })] }),
      first: new Footer({ children: [new Paragraph({ children: [] })] })
    },
    children: body
  }]
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log('written', OUT, buf.length); });
