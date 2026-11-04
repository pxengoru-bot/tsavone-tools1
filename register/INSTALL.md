# Tsavone Training Register: installation and user guide

This tool runs inside your Google Sheet **"Tsavone Training Register"**. It adds a **Tsavone** menu that:

- fills in course titles
- works out who passed
- gives each certificate a unique ID
- makes A4 certificate PDFs with a QR code
- prepares certificate emails as Gmail drafts
- builds a course summary for your completion report

It also gives you a public **verification page** where anyone can check that a certificate is genuine.

Everything runs on your own Google account. No outside services are used.

---

## What is in this folder

| File | What it is | Where it goes |
|---|---|---|
| `Code.gs` | The main program | Apps Script, file "Code" |
| `Qr.gs` | Makes the QR codes | Apps Script, file "Qr" |
| `Verify.html` | The verification web page | Apps Script, HTML file "Verify" |
| `tsavone-logo.png` | Logo for certificates (PNG, because Google Docs cannot use SVG logos) | Upload to Google Drive (step 4) |
| `sample-rows.tsv` | Three test participants | Paste into the sheet to try the tool (section 8) |
| `tests/` | Automatic checks for developers | Not needed in Apps Script |

To copy a file from GitHub: open it, click **Raw** (or the copy icon), select all (Ctrl+A), copy (Ctrl+C).

---

## Step 1: Paste the code into Apps Script (10 minutes)

1. Open your Google Sheet **Tsavone Training Register**.
2. Click **Extensions → Apps Script**. A new tab opens with a file called `Code.gs`.
3. Click the project name at the top ("Untitled project") and rename it **Tsavone Training Register**.
4. **Code.gs:** delete everything in the editor. Paste the full contents of `Code.gs` from this folder.
5. **Qr.gs:** click the **+** next to "Files" and choose **Script**. Name it `Qr` (Apps Script adds `.gs`). Delete what is in it, then paste the contents of `Qr.gs`.
6. **Verify.html:** click **+** again and choose **HTML**. Name it exactly `Verify`. Delete what is in it, then paste the contents of `Verify.html`.
7. Click the **Save** icon (or Ctrl+S).

You should now have three files: `Code.gs`, `Qr.gs` and `Verify.html`.

---

## Step 2: Run setup once and approve permissions

1. In the Apps Script toolbar, find the box that lists functions (next to **Debug**). Choose **setup**.
2. Click **Run**.
3. Google asks for permission. Click **Review permissions**, choose your Google account, and then:
   - You will see "Google hasn't verified this app". This is normal for your own scripts. Click **Advanced**, then **Go to Tsavone Training Register (unsafe)**.
   - Click **Allow**.
4. Wait until the log says "Execution completed".
5. Go back to your sheet and **reload the page**. A **Tsavone** menu now appears next to Help.

**What setup does:**
- Names the first tab **Register**.
- Writes the column headings, or adds any that are missing.
- Freezes and colours the header row (navy with white bold text).
- Adds dropdowns for format, status and certificate type, and sets date formats.
- Creates a **Settings** tab.

You can run it again at any time (Tsavone → Run setup again). It never deletes your data or overwrites your settings.

**Why the permissions are needed:**

| Permission | Used to |
|---|---|
| Sheets | Read and write the register |
| Drive | Read the logo and save certificate PDFs |
| Docs | Lay out each certificate before it becomes a PDF |
| Gmail | Create **drafts** only; nothing is sent automatically. (Google's simpler MailApp cannot make drafts, so the script uses GmailApp.) |

---

## Step 3: Check the Settings tab

Open the **Settings** tab. Change only the **value** column (column B).

| Setting | Starts as | Meaning |
|---|---|---|
| `pass_attendance_percent` | 80 | Minimum attendance (%) to pass |
| `pass_post_test_score` | 60 | Minimum post-test score to pass |
| `competence_post_test_score` | 75 | Post-test at or above this earns a **Certificate of Competence**; below it (but passing) earns a **Certificate of Completion** |
| `certificate_folder_name` | Tsavone Certificates | Folder in your Google Drive (My Drive) where PDFs are saved, one sub-folder per course code |
| `logo_drive_file_id` | (blank) | Fill in at step 4 |
| `signatory_title` | Managing Director | Printed under the signature line. No personal name is printed. |
| `sender_email_name` | Tsavone Consulting Training | Sender name on the certificate emails |
| `verify_url` | (blank) | Fill in at step 5 |
| `course_sheet_id` | 10V8_jXP… | Your public "Tsavone Courses" sheet, used to look up course titles and days |
| `hours_per_day` | 7 | Course hours on certificates = days × this number |

---

## Step 4: Upload the logo and copy its file ID

1. Download `tsavone-logo.png` from this folder. You can use your own logo instead, but it must be **PNG or JPG, not SVG**.
2. Upload it to Google Drive (New → File upload).
3. In Drive, right-click the uploaded logo and choose **Share → Copy link**. The link looks like:
   `https://drive.google.com/file/d/`**`1AbCdEfGhIjKlMnOpQrStUvWxYz`**`/view?usp=sharing`
4. Copy only the long code between `/d/` and `/view`. That is the **file ID**.
5. Paste it into **Settings → logo_drive_file_id**.

You do not need to make the logo public. The script uses it as you.

---

## Step 5: Publish the verification page and save its link

Every certificate has a QR code that opens this page, so do this **before** making real certificates.

1. In Apps Script, click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Fill in:
   - **Description:** Tsavone certificate verification
   - **Execute as:** **Me**
   - **Who has access:** **Anyone**
4. Click **Deploy**. Approve permissions again if asked.
5. Copy the **Web app URL**. It ends in `/exec`.
6. Paste it into **Settings → verify_url**.
7. Test it: open the link in a private browser window. You should see "Verify a Tsavone certificate" with a box. Type any ID and you should get **NOT FOUND**.

**Important: when you update the code later, keep the same link.**
1. Use **Deploy → Manage deployments**.
2. Click the pencil (edit) icon.
3. Set **Version: New version**.
4. Click **Deploy**.

Do **not** create a new deployment each time. That gives a new link, and QR codes on certificates already issued would stop working.

**What the page shows:**

| Certificate status | Page shows |
|---|---|
| **Issued** | **VALID**, with participant name, course title, dates and certificate type |
| **Revoked** | **REVOKED** |
| Anything else, including Draft or an unknown ID | **NOT FOUND** |

Email, phone, organisation, scores and notes are **never** shown.

---

## Step 6: Using the register for a real course

### Before the course
1. Add one row per participant on the **Register** tab. Fill in:
   - participant_name, organisation, email, phone
   - course_code (for example `T1.1`)
   - start_date, end_date (for example 10 Mar 2027)
   - location, format (from the dropdown)
   - status = **Draft**

   Leave certificate_id, course_title, result, certificate_type and issued_date **empty**. The tool fills them.
2. **Tsavone → Fill course titles from codes.** Titles come from your course list, and only empty title cells are filled.

### After the course
3. Enter **attendance_percent** (for example 95), **pre_test_score** and **post_test_score** for each person.
4. **Tsavone → Calculate results.** For each row:
   - **result** becomes **Pass** if attendance and post-test both meet the Settings, otherwise **Did not complete**.
   - **certificate_type** becomes **Competence** or **Completion** (left empty if the person did not pass).
   - **notes** gets a line such as "Improvement: +23 points (pre 45, post 68)". Your other notes are kept.
   - Rows already **Issued** or **Revoked** are never changed.
5. *(Optional)* **Tsavone → Create certificate IDs.**
   - Each passing participant without an ID gets one, such as `TSV-2027-T1.1-0001`.
   - The year comes from the course start date, and numbers count up separately for each year and course.
   - IDs are never reused while a row exists.
   - The next step does this automatically if you skip it.

### Certificates
6. **Tsavone → Generate certificates.**
   - You are asked what to do:
     - **NO** = make certificates for **all** passing participants who do not yet have an issued_date. This is the usual choice after a course.
     - **YES** = make or **remake** certificates for just the rows you selected (click the grey row numbers on the left; Ctrl-click for several). Any older PDF for those rows is replaced.
   - If you start from another tab (for example Settings), there is no question and all pending certificates are made.
   - Each PDF is saved in **Google Drive → Tsavone Certificates → [course code]**, named like `TSV-2027-T1.1-0001 - Amina Nakato.pdf`.
   - The row's issued_date is set to today (an existing issue date is kept when you remake a certificate), and its status becomes **Issued**.
   - Google stops a run after 6 minutes, which is usually enough for several dozen certificates. If you see "I stopped early", just run it again (choose NO) and it carries on.
   - **Open the first PDF and check it** before emailing anyone.
7. **Tsavone → Email certificates (create Gmail drafts).**
   - First select the rows to email, then choose the menu item.
   - A Gmail **draft** is created for each Issued row with a valid email. Each draft has the PDF attached, a short professional message, the certificate ID and the verification link.
   - Open **Gmail → Drafts**, check each one, and press **Send**. Nothing is sent without you.

### Reporting
8. **Tsavone → Course summary.** Pick the course from the numbered list (course code and start date). A new tab, such as **"Summary T1.1 2027-03-10"**, shows:
   - participants, passes, pass rate
   - average attendance, average pre-test, post-test and improvement
   - certificates issued
   - counts by organisation, and by sex or gender if you add a column called `sex` or `gender` to the Register

   It also gives a ready-made sentence for the completion report. Copy the table into your report. Running it again replaces the tab with fresh numbers.

### Corrections
- **Wrong name or details on a certificate:** correct the row, select it, run **Generate certificates → YES**, then email a new draft.
- **Revoke a certificate:** set its status to **Revoked**. The verification page shows REVOKED straight away.

---

## Step 7: Changing the design

The colours and fonts are near the top of `Code.gs`:
- `NAVY`, `ORANGE`, `LIGHT_BLUE`
- `FONT_HEAD = 'Montserrat'`, `FONT_BODY = 'Source Sans 3'`

If a font is not available in Google Docs, Docs uses a standard font instead. After any change, click Save. To update the verification page as well, follow the "keep the same link" steps in step 5.

---

## 8. Try it first with three test participants

1. Open `sample-rows.tsv` in this folder, click **Raw**, then select all and copy.
2. In the **Register** tab, click cell **A2** (the first cell under `certificate_id`) and paste (Ctrl+V). Three rows appear.
   - If your sheet already has participants, click column A of the first empty row instead.
3. Change the three email addresses to **your own address**. For example, `info+test1@tsavone.com`, `info+test2@tsavone.com` and `info+test3@tsavone.com` all arrive in the info@tsavone.com inbox.
4. Run the menu items in order and check what you get:

| Menu item | What you should see |
|---|---|
| Fill course titles from codes | All three rows show "Practical Monitoring, Evaluation and Learning for Programmes" |
| Calculate results | One: **Pass / Competence**. Two: **Pass / Completion**. Three: **Did not complete** (attendance 60%). Notes show the improvement, for example "+37 points" for One. |
| Create certificate IDs | One: `TSV-2027-T1.1-0001`. Two: `TSV-2027-T1.1-0002`. Three: no ID. |
| Generate certificates, and choose **NO** when asked | 2 PDFs in Drive → Tsavone Certificates → T1.1; both rows become **Issued** |
| Scan a certificate's QR code with your phone | The verification page shows **VALID** for that participant |
| Select rows 2–4, then Email certificates | 2 Gmail drafts; row 4 is skipped because it did not pass |
| Course summary → choose T1.1, 10 Mar 2027 | Participants 3, pass rate 67%, average improvement +23.3 |
| Change row 2's status to **Revoked**, then refresh the verification page | **REVOKED** |

### Delete the test afterwards
1. **Register tab:** select the three test rows (click row numbers 2–4), right-click and choose **Delete rows**.
2. **Google Drive:** open **Tsavone Certificates → T1.1** and delete the two test PDFs. Delete the T1.1 folder if it is empty, then empty the Drive **Bin**.
3. **Gmail → Drafts:** delete the two test drafts.
4. **Summary tab:** right-click the **"Summary T1.1 2027-03-10"** tab and choose **Delete**.

Once the test rows are deleted, numbering starts again from `-0001` for real participants on the same course and year. That is why you should never send out test certificates.

---

## 9. Good to know

- **Privacy:** the verification page shows the participant's name, course, dates and certificate type to anyone who has a certificate ID. IDs follow a pattern (`-0001`, `-0002` …), so someone could guess other IDs and see who completed a course. This is normal for public certificate checkers. If it ever matters for a sensitive course, keep those rows in **Draft** status; they then show NOT FOUND.
- **Where data lives:** everything stays in your Google account (Sheet, Drive, Gmail). The QR codes are made inside the script, not by an outside website.
- **Several people using the sheet:** anyone with edit access can use the menu. Certificates are saved in the Drive of whoever runs it, and Gmail drafts go to their Gmail. The verification page always reads the register as **you**.

---

## 10. If something goes wrong

| Message or problem | What to do |
|---|---|
| No **Tsavone** menu | Reload the sheet. If it is still missing, run **setup** again from Apps Script (step 2). |
| "Verification page not set up" | Finish step 5 and paste the `/exec` link into Settings → verify_url. |
| Logo missing on certificates; the message mentions the logo | Check logo_drive_file_id (step 4). The file must be PNG or JPG. |
| "These codes were not found in the course list" | Check the course code spelling (for example `T1.1`) against your "Tsavone Courses" sheet. |
| "Missing attendance or post-test score on row(s) …" | Fill in those cells, then run Calculate results again. |
| Certificate text or layout looks wrong | Send a screenshot of the PDF so the layout can be adjusted in `Code.gs`. |
| QR code opens the wrong page | Check Settings → verify_url, then remake the affected certificates (select rows → Generate → YES). |
| "Authorisation is required" | Run **setup** once more from Apps Script and click Allow. |

---

## For developers: automatic checks

The files in `tests/` run outside Google with Node.js:

| Test | What it checks |
|---|---|
| `qr.test.js` | QR codes decode correctly (versions 1–13) and the PNG is valid. Needs the `jsqr` package. |
| `logic.test.js` | Results, IDs, dates, summaries, email text, and that the verification lookup reveals no private fields. |
| `flow.test.js` | Every menu item end to end, against strict stand-ins that only accept real Apps Script method names |

Run them from the repository root:

```bash
npm install --no-save jsqr@1.4.0
```

```bash
node register/tests/qr.test.js
```

```bash
node register/tests/logic.test.js
```

```bash
node register/tests/flow.test.js
```
