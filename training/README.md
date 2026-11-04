# Tsavone Training website

This folder holds the complete Tsavone Consulting training website. It is a plain website (HTML, CSS and JavaScript). It needs no build step, database or special hosting. Upload the folder as it is and it works.

**What is in the folder**

| File or folder | What it is |
|---|---|
| `index.html` | The website: header, hero, courses, how we deliver, fees, booking form, contact and footer |
| `thank-you.html` | The page people see after sending the booking form |
| `css/styles.css` | Colours, fonts and layout |
| `js/app.js` | Loads courses from the Google Sheet, runs search, filters, the fee estimator and the form. **Fee and discount rules are at the top of this file.** |
| `data/courses-fallback.json` | A saved copy of the course list, used if the Google Sheet cannot be reached |
| `assets/` | Logo, icon, favicon and the picture shown when the site is shared on social media |

---

## 1. Managing courses in the Google Sheet

The course list on the website comes **live from the Google Sheet "Tsavone Courses"**. When you change the sheet, the website shows the change the next time someone opens or refreshes the page. Google can take up to about 5 minutes to pass on a change.

**Do not rename the column headings in row 1.** The columns are:

| Column | What to type | Example |
|---|---|---|
| `code` | A short, unique course code | `T1.5` |
| `track_no` | The track number | `1` |
| `track_name` | The track name. Type it exactly the same for every course in the track. | `Evidence, MEL and data` |
| `title` | The course title | `Qualitative Research Methods` |
| `days` | Number of days (a number only) | `3` |
| `formats` | Formats, separated by commas. Use: Open, In-house, Blended, Field, ToT | `Open, in-house, field` |
| `audience` | Who the course is for | `Researchers and M&E staff` |
| `topics` | Topics, separated by a space, a vertical bar and a space: ` \| ` | `Interviews \| Focus groups \| Coding and analysis` |
| `takes_home` | What participants take home | `Interview guide template, coding sheet` |
| `fee_ugx` | Fee per participant in Uganda shillings, **numbers only, no commas** | `1050000` |
| `fee_usd` | Fee per participant in US dollars, numbers only | `350` |
| `status` | One of the words below | `Active` |

**The status column controls what visitors see:**

| Status | What happens on the website |
|---|---|
| `Active` | The course is shown |
| `New` | The course is shown with an orange **NEW** badge |
| `Full` | The course is shown with a **FULLY BOOKED** label, and the booking button is switched off |
| `Hidden` (or left empty) | The course is not shown at all |

### Add a course
1. Open the sheet and go to the first empty row at the bottom.
2. Fill in every column. Set `status` to `New` or `Active`.
3. Refresh the website. The course appears in its track, in the booking form list and in the fee estimator. The course count at the top of the page updates by itself.

### Edit a course
Change the cell in the sheet. The website updates on the next refresh.

### Hide a course (without deleting it)
Change its `status` to `Hidden`. To bring it back later, change it to `Active`.

### Mark a course as full
Change its `status` to `Full`. Change it back to `Active` when places open again.

### Add a new track
Give a course a new `track_no` (for example `9`) and a new `track_name`. The new track appears as a filter button by itself. Tracks 1 to 8 have fixed colours; new tracks get extra colours automatically.

### Keep the sheet shared
The sheet must stay shared as **"Anyone with the link: Viewer"** (Share button, top right of the sheet). If it is made private, the website cannot read it and falls back to the saved list.

### Update the saved copy (about once a term)
If the sheet cannot be reached, the website shows the saved list in `data/courses-fallback.json`, with the note "Showing saved course list". To refresh that saved copy after big changes, ask your web person to export the sheet again. Alternatively, open this link in your browser and send the downloaded CSV to them:
`https://docs.google.com/spreadsheets/d/10V8_jXPTFwRl2y940v_AtmMDPU8RGQBI6yH5qDy6N4I/gviz/tq?tqx=out:csv&gid=0`

The website still works if the saved copy is a little out of date. The live sheet always wins when it can be reached.

---

## 2. Changing fees and discounts

There are two places to change:

**a) The fee table and discount text that visitors read.** Open `index.html` in a text editor (for example Notepad, TextEdit or the Hostinger File Manager editor). Search for `fee-table` and change the numbers in the table. Search for `discount-list` and change the discount wording.

**b) The rules the fee estimator uses.** Open `js/app.js`. The settings are near the top, under `SETTINGS`:

```
onlineDiscount: 0.40,          // online about 40% less
groupSmallMin: 3,              // 3–4 people from one organisation ...
groupSmallDiscount: 0.10,      // ... get 10%
groupLargeMin: 5,              // 5 or more ...
groupLargeDiscount: 0.15,      // ... get 15%
earlyDiscount: 0.10,           // booking at least 30 days ahead
selfSponsoredDiscount: 0.20,   // self-sponsored individuals
inHouseUgxPerDayMin: 3500000, inHouseUgxPerDayMax: 5000000, inHouseMaxPeople: 25,
```

Percentages are written as decimals: 0.10 = 10%, 0.15 = 15%, 0.40 = 40%. Change only the numbers and leave the commas and the rest of the line as they are.

**Course fees themselves** come from the `fee_ugx` and `fee_usd` columns in the Google Sheet, so you change them there.

How the estimator works: it starts from the course fee × number of participants. It then takes off each discount that applies, one after the other: online, then group, then early booking, then self-sponsored. Only the larger group discount (10% or 15%) is applied, and only once. The self-sponsored rate is for individuals paying for themselves, so it replaces the group discount. Shilling totals are rounded to the nearest 1,000. The page clearly labels the result as an estimate, with the final quote sent by email.

---

## 3. The booking form (FormSubmit)

The booking form sends enquiries to **info@tsavone.com** through the free service FormSubmit (formsubmit.co). There is nothing to install.

**Important, one-time step:** the **first** time anyone submits the form on the live website, FormSubmit sends a confirmation email to info@tsavone.com titled something like "Action Required: Activate FormSubmit". **Open that email and click "Activate Form".** Until you do, enquiries are not delivered. Send a test enquiry yourself straight after publishing so this happens immediately. After activation, every enquiry arrives as a neat table by email, and you can reply directly to the sender.

Spam protection: the form has a hidden "honeypot" field that people never see. Spam robots fill it in, and FormSubmit then throws their message away. If you still get a lot of spam, open `index.html`, find `name="_captcha" value="false"` and delete that line. FormSubmit will then show a short "I am not a robot" check.

If the form ever fails, visitors can use the **"Write to info@tsavone.com"** link next to the form. It opens their email app with their details already filled in.

After sending, visitors see `thank-you.html`.

---

## 4. Publishing on GitHub Pages (free)

1. Make sure the `training` folder is on the repository's main branch (merge the pull request if the work is on another branch).
2. In the repository on GitHub, go to **Settings → Pages**.
3. Under "Build and deployment", set **Source: Deploy from a branch**, **Branch: main**, **Folder: / (root)**. Click **Save**.
4. Wait 1–2 minutes. The site will be at:
   **https://pxengoru-bot.github.io/tsavone-tools1/training/**

GitHub Pages is free for public repositories. For a private repository it needs a paid GitHub plan.

GitHub Pages is useful as a free preview. For the real site at tsavone.com/training, use Hostinger (below), because your main tsavone.com website is hosted there.

---

## 5. Uploading to Hostinger

The site is set up for **https://www.tsavone.com/training/**.

### Option A: as tsavone.com/training (set up and recommended)
1. Log in to Hostinger **hPanel → Websites → tsavone.com → Files → File Manager**, and open `public_html`. This is the folder that already holds your main tsavone.com website.
2. Create a new folder inside `public_html` called `training`, all lower case.
3. Upload **everything inside** this repository's `training` folder into `public_html/training`: `index.html`, `thank-you.html`, `README.md` (optional), and the `css`, `js`, `data` and `assets` folders. `index.html` must sit directly inside `public_html/training`, not in a further subfolder. Tip: zip the contents on your computer, upload the zip, then right-click → **Extract**.
4. Visit https://www.tsavone.com/training/ and send a test enquiry (see section 3).
5. Optional: add a "Training" link to your main website's menu, pointing to `/training/`.

Your existing website is not affected. The training site lives in its own folder and only uses its own files.

If your main site is built with WordPress or the Hostinger Website Builder, check first that no existing page already uses the address `/training`. If one does, rename that page's address, or the two will clash.

### Option B: as training.tsavone.com
1. In hPanel go to **Domains → Subdomains** and create the subdomain `training`. Note the folder Hostinger creates for it.
2. Upload the contents of the `training` folder into that folder, as in Option A, step 3.
3. Turn on the free SSL certificate for training.tsavone.com (**Security → SSL**).
4. Open `index.html` and use "Find and replace" to change every `https://www.tsavone.com/training/` to `https://training.tsavone.com/`.

### Why the web address matters
`index.html` tells Google and social media the site's official address. It appears in the `canonical`, `og:url`, `og:image` and `twitter:image` lines and in the course list for Google. All other links are relative, so the site works in any folder. If you ever move it, find and replace that address as in Option B, step 4.

---

## 6. Good to know

- **No names of people** appear anywhere on the site, by design.
- **Search engines:** the page includes Google-friendly course information (schema.org `Course`). You can check it at https://search.google.com/test/rich-results once the site is live.
- **Speed:** the page, styles and script together are about 40 KB when compressed, plus fonts and logos, and shows the saved course list after 2.5 seconds if the sheet is slow. That keeps it fast on mobile data.
- **Sharing a single course:** each course has its own link. Open a course and copy the address bar (it ends in, for example, `#course-T1.1`) to send someone straight to it.
- **Changing contact details:** search `index.html` for the phone number or email address and edit it in each place it appears. Do the same in `thank-you.html`.
