# Tsavone Tender Tracker: installation and user guide

This tool runs inside your **private** Google Sheet **"Tsavone Tender Tracker"**. It adds a **Tsavone Bids** menu that:

- adds tenders through a simple form, each with a reference number such as `TND-2026-0001`
- colours each tender by how close its deadline is
- creates a private Google Drive folder and a **bid checklist** for each tender, matched against your company documents
- marks tenders as submitted
- keeps a **Dashboard** of deadlines, win rate and pipeline value
- emails you a short reminder each morning, **only** when something needs attention

**Nothing is made public.** There is no web page, folders and documents stay private to your Google account, and the reminder email goes only to `p.engoru@tsavone.com`.

---

## What is in this folder

| File | What it is |
|---|---|
| `Code.gs` | The program. Paste it into Apps Script as "Code". |
| `Sidebar.html` | The "Add new tender" form. Paste it into Apps Script as an HTML file called "Sidebar". |
| `sample-tenders.tsv` | Three sample tenders, if you prefer to paste them by hand. The **Testing** menu does this for you. |
| `tests/` | Automatic checks for developers. Not needed in Apps Script. |

---

## Step 1: Paste the two files (5 minutes)

1. Open your sheet **Tsavone Tender Tracker**. Click **Extensions → Apps Script**. A new tab opens.
2. Click **Untitled project** at the top and rename it **Tsavone Tender Tracker**.
3. **Code.gs**
   1. In a new browser tab, open https://raw.githubusercontent.com/pxengoru-bot/tsavone-tools1/main/tenders/Code.gs
   2. Press **Ctrl+A**, then **Ctrl+C**.
   3. In Apps Script, click **Code.gs** on the left, click inside the code area, press **Ctrl+A**, then **Ctrl+V**.
4. **Sidebar.html**
   1. In Apps Script, click the **+** next to "Files", choose **HTML** and name it exactly `Sidebar`.
   2. Open https://raw.githubusercontent.com/pxengoru-bot/tsavone-tools1/main/tenders/Sidebar.html and press **Ctrl+A**, **Ctrl+C**.
   3. Back in Apps Script, click inside **Sidebar.html**, press **Ctrl+A**, then **Ctrl+V**.
5. Press **Ctrl+S** to save.

You should now see two files on the left: **Code.gs** and **Sidebar.html**.

---

## Step 2: Run setup once and approve permissions

1. In the Apps Script toolbar, click the box next to **Debug**, choose **setup**, then click **▶ Run**.
2. A box says **"Authorisation required"**. Click **Review permissions** and choose your account.
3. On "Google hasn't verified this app", click **Advanced**, then **Go to Tsavone Tender Tracker (unsafe)**. This is normal for your own scripts.
4. Tick **Select all** and click **Continue**.
5. Wait for **"Execution completed"** at the bottom.
6. Go back to the sheet and press **F5**. A **Tsavone Bids** menu appears after Help.

**What setup does:**
- Names the first tab **Tenders**, styles and freezes the header, and adds the dropdowns, date formats and UGX format.
- Adds the deadline colours.
- Adds one extra column, **submitted_date**, at the end. "Mark as submitted" fills it, and the dashboard uses it to count submissions per year.
- Creates the **Documents** tab with your 20 standing company documents, and the **Dashboard** tab.
- Sets the sheet to Kampala time.

You can run it again at any time (**Tsavone Bids → Run setup again**). It never deletes your data.

**Why the permissions are needed:**

| Permission | Used to |
|---|---|
| Sheets | Read and update the tracker |
| Drive | Create private bid folders in "Tsavone Bids" |
| Docs | Write the bid checklists |
| Send email as you | Send the daily reminder to `p.engoru@tsavone.com` only |
| Run when you are not present | The daily reminder at 07:00 |
| Display content in Google apps | The Tsavone Bids menu and the "Add new tender" panel |

---

## Step 3: Turn on the daily reminder

Click **Tsavone Bids → Daily reminder → Turn daily reminder on**.

- Every morning between **07:00 and 08:00 Kampala time** (Google picks a moment within that hour), the tracker checks:
  - tenders due within **7 days** that are not yet submitted
  - company documents **expired** or **expiring within 30 days**
- If there is something, you get **one short email** with a link to the sheet. If there is nothing, no email is sent.
- **Turn it off:** Tsavone Bids → Daily reminder → **Turn daily reminder off**.
- **Test it now:** Tsavone Bids → Daily reminder → **Send reminder now (test)**.

To send the reminder to a different address, change the line `var REMINDER_EMAIL = 'p.engoru@tsavone.com';` near the top of `Code.gs`, then save.

---

## Step 4: Fill in the Documents tab (once, then keep it up to date)

The **Documents** tab lists your standing company documents. For each one, fill in:

| Column | What to put |
|---|---|
| status | **Have**, **Missing** or **Expired** |
| issue_date / expiry_date | Dates, for documents that expire (tax clearance, trading licence, NSSF compliance and so on) |
| drive_link | The Google Drive link to the current copy. Right-click the file in Drive, choose **Share → Copy link**, and paste it here. Sharing does **not** need to change; the link is just for you. |
| notes | Anything useful |

**Colours:** a row turns **orange** when its expiry date is within 30 days, and **red** when it has expired.

You can add your own documents at the bottom, such as an ISO certificate or professional body membership. The checklist matches those too.

---

## Using the menu for a real tender

### Add new tender
**Tsavone Bids → Add new tender** opens a form on the right. Fill in at least the organisation and title. Add the deadline and paste the list of required documents from the notice, one per line or separated by commas. Click **Add tender**. A new row appears with its reference number. The form stays open so you can add another.

You can also type straight into the next empty row. As soon as you type an organisation or title, the row gets its reference number, today's date found, status **Watching**, go/no-go **Undecided** and result **Pending**.

### Deadline colours (open tenders only)

| Colour | Meaning |
|---|---|
| **Red** | Due within 3 days |
| **Orange** | Due within 7 days |
| **Yellow** | Due within 14 days |
| **Grey** | Deadline passed |

Rows with status **Submitted, Won, Lost, Not pursued** or **Cancelled** are not coloured.

### Create bid folder and checklist
1. Click any cell in the tender's row.
2. Click **Tsavone Bids → Create bid folder and checklist**.

You get:
- A private folder in Google Drive: **Tsavone Bids / TND-2026-0001 - Organisation - Short title**.
- A **Bid checklist** Google Doc inside it, containing:
  - **Tender details:** the organisation, deadline (with "in N days"), submission method, lead person and links.
  - **Required documents:** each item from documents_required, matched to your Documents tab, with its **status** (green Have, red Missing/Expired, yellow other) and **Drive link**. Items not in your Documents tab are marked "Prepare for this bid", for example a signed bid form.
  - **Bid steps:** clarification deadline, go/no-go, pricing, technical proposal, signatures and stamps, compliance check, and submission with proof of submission, each with columns for who and by when.
- The folder link is added to the tender's **notes**. If **documents_ready** is empty, it is filled with "N of M ready".

If you run it again for the same tender, you are asked whether to replace the checklist. The old one goes to the Drive Bin.

### Mark as submitted
1. Click any cell in the tender's row.
2. Click **Tsavone Bids → Mark as submitted**.
3. Type the submission date, or leave it empty for today. You can use 5/11/2026 (day first), 5 Nov 2026 or 2026-11-05.

Status becomes **Submitted**, **submitted_date** is filled, result becomes **Pending**, and go/no-go becomes **Go** if it was undecided. Save the proof of submission in the bid folder.

When the result comes in, set **status** (and **result**) to **Won**, **Lost** or **No response**.

### Refresh dashboard
**Tsavone Bids → Refresh dashboard** rebuilds the **Dashboard** tab. The numbers are live formulas, so they keep updating on their own. Refresh it if you add columns or change categories. It shows:
- Key figures:
  - tenders tracked and open
  - due in 7 days
  - **submitted this year**
  - won, lost, **win rate**
  - **total value won**
  - **value in pipeline** (Preparing, Submitted and Shortlisted)
- Counts by **status**, and count and value by **category**.
- Lists of tenders **due in the next 7 days** and the **next 30 days** that are not yet submitted.
- **Documents** that are missing, expired or expiring within 30 days.

---

## Try it first with sample tenders

1. Click **Tsavone Bids → Testing → Add sample tenders**. Three sample tenders appear, due in **2, 6 and 12 days**, so they show red, orange and yellow.
2. Click a sample row, then **Create bid folder and checklist**. Open the folder link from the notes column to see the checklist.
3. Click **Refresh dashboard** and look at the Dashboard tab.
4. Click **Daily reminder → Send reminder now (test)**. You should receive an email listing the two sample tenders due within 7 days.
5. Click a sample row, then **Mark as submitted**, and see the colour disappear.

### Delete the samples afterwards
Click **Tsavone Bids → Testing → Remove sample tenders** and confirm. This removes the three sample rows and their bid folders (moved to the Drive Bin). Only rows whose notes start with "TEST ROW" are removed; real tenders are never touched.

(If you prefer to paste the samples by hand from `sample-tenders.tsv`: click the first empty cell in column A and paste. If everything lands in column A, select those cells and use **Data → Split text to columns** with separator **Tab**. Their deadlines are fixed dates in November 2026, so the colours depend on today's date.)

---

## Good to know

- **Privacy:** the sheet, folders and checklists are private. Nothing is published, and no link is shared. Anyone you share the sheet with as an editor can use the menu; their bid folders are created in their own Drive.
- **Use a computer:** the Tsavone Bids menu does not appear in the Google Sheets phone app. The reminder email can be read anywhere.
- **Reference numbers** follow the year the tender was found: TND-2026-0001, TND-2026-0002 and so on. They are never reused while a row exists.
- **Updating the program later:** paste the new Code.gs (and Sidebar.html if it changed) as in step 1 and save. Nothing needs redeploying.

## If something goes wrong

| Problem | What to do |
|---|---|
| No Tsavone Bids menu | Reload the sheet (F5) and wait 10 seconds. If it is still missing, run **setup** from Apps Script (step 2). |
| "Go to the Tenders tab and click any cell…" | Click a cell in the tender's row on the Tenders tab, then use the menu again. |
| A required document shows "Not in Documents tab" but you have it | Add it to the Documents tab with similar wording, or use the standard name. |
| No reminder email | Check it is turned on (Daily reminder → Turn daily reminder on). Use **Send reminder now (test)**. Remember that no email is sent when nothing needs attention. |
| Google asks for permission again | Approve it in the same way as step 2. This happens when a new feature is used for the first time. |

## For developers

The `tests/` folder holds two automatic checks, run with Node.js from the repository root:

| Test | What it checks |
|---|---|
| `logic.test.js` | Dates, reference numbers, document matching, reminder rules and dashboard formulas |
| `flow.test.js` | Every menu item end to end, against strict stand-ins that only accept real Apps Script method names |

```bash
node tenders/tests/logic.test.js
```

```bash
node tenders/tests/flow.test.js
```
