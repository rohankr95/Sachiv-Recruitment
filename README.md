# Sachiv Recruitment Surajpur – Application Correction Portal (Google Apps Script)

Applicants log in with their **registered email address or mobile number plus date of birth**. They can then review their existing application, correct details, replace documents, preview the changes and submit them. Everything stays in your **existing Google Sheet** and **Google Drive**. There is no external database, server or paid service.

---

## 1. Architecture

```
 Browser (applicant / admin)                Google Apps Script web app                Google Workspace
 ─────────────────────────────              ───────────────────────────                ─────────────────
 Index.html / Admin.html        google.script.run   Code.gs  (doGet, helpers)
  ├ Login.html            ─────────────────────▶   Authentication.gs (login, sessions,   ─▶  CacheService
  ├ ApplicationForm.html  ◀─────────────────────     rate limit)                              (sessions, limits,
  ├ Preview.html            JSON {ok,data|error}    Application.gs   (field mapping,           Field_Config cache)
  ├ Styles.html                                      load, validate, submit)          ─▶  Spreadsheet
  └ JavaScript.html                                 DocumentUpload.gs (stage/view/       │   ├ Form Responses 1 (records)
                                                     replace/archive)                    │   ├ Field_Config
 Only a random session token is held in the         AuditLog.gs      (append-only log)   │   └ Correction_Log
 browser. No sheet IDs, folder IDs or other          Admin.gs         (dashboard)    ─▶  Drive
 applicants' data are ever sent.                    Setup.gs / SampleData.gs (editor only)    ├ Replacement Uploads
                                                     LockService serialises all writes         ├ Archived Documents
                                                                                               └ Staging (temporary)
```

**How Apps Script works with Sheets and Drive**

* The web app is deployed with **Execute as: Me** and **Who has access: Anyone**. Applicants do not need Google accounts. All Sheet and Drive access runs as the owner, so the spreadsheet and documents can stay **private (Restricted)**.
* **Login** reads only the 4–7 columns it needs (email/mobile, DOB, application no., name, status), never the whole sheet. Matching happens on the server. On success a 64-character random token is created and the list of permitted application numbers is stored in `CacheService` next to it.
* **Retrieval** finds the row with `TextFinder` on the Application Number column, then reads that one row in three calls (values, display values, formulas).
* **Submission** takes a `LockService` script lock, re-reads the row, checks that nobody saved a newer version (optimistic concurrency), validates every value again on the server, archives and moves documents, writes **only the changed cells** plus the 4 system columns, and appends the audit rows in one batch.
* **Documents** are never made public. The portal streams them to the browser as base64 after it has checked the session (PDFs render with pdf.js, which also works on mobile).

## 2. Files

| File | Purpose |
|---|---|
| `Config.gs` | Spreadsheet ID, sheet names, key column aliases, Drive folder IDs, limits, timeouts, workflow switches |
| `Code.gs` | `doGet`, `include`, response wrapper, normalisation/sanitisation, hashing helpers |
| `Authentication.gs` | Email/mobile + DOB login, duplicate detection, sessions, rate limiting |
| `Application.gs` | Dynamic field mapping (`Field_Config` + auto-detection), load, validation, submission |
| `DocumentUpload.gs` | Upload staging, signature checks, viewer, archive + replace, staging cleanup |
| `AuditLog.gs` | Append-only `Correction_Log`, history, statistics |
| `Admin.gs` | Admin login (hashed passwords), dashboard, search, history, duplicates, approve/reject |
| `Setup.gs` | `setupPortal()`: system columns, Field_Config, Correction_Log, folders, trigger, sharing checks |
| `SampleData.gs` | `createSampleData()`: 2,500-row synthetic test spreadsheet |
| `Index.html`, `Login.html`, `ApplicationForm.html`, `Preview.html`, `Styles.html`, `JavaScript.html` | Applicant UI |
| `Admin.html`, `AdminJavaScript.html` | Admin dashboard UI (`?page=admin`) |
| `appsscript.json` | Manifest (V8, web app settings, OAuth scopes) |
| `local-test/` | Offline test harness (Node + Playwright). Not uploaded to Apps Script |

## 3. Login rules (email or mobile + DOB, no OTP)

* The applicant enters **either** their email address **or** their mobile number, plus their date of birth. The server treats input containing `@` as an email and anything else as a mobile number. Mobile numbers are normalised, so `+91 98765 43210`, `098765 43210` and `9876543210` all match. Email matching ignores case.
* The DOB is compared as a calendar date, whether the sheet holds a real date, `15/08/2000`, `2000-08-15`, `15-Aug-2000` or a serial number. Set `CONFIG.DATE.TEXT_DATE_ORDER` to `'MDY'` if your text dates are month-first.
* **Duplicates.** When the same email or mobile appears on more than one row:
  * **`MATCH_DOB` (default):** every application with that email/mobile **and** the entered DOB is listed, and each one is marked **Duplicate** with the number of applications sharing the email. This covers one person who applied more than once. Rows with the same email but a *different* DOB belong to a different person (for example a shared cyber-café email), so they are **not** shown.
  * **`SHOW_ALL`:** once the DOB matches any row, **all** applications on that email/mobile are listed and marked Duplicate. Use this only if you are sure shared emails always belong to the same person, because otherwise one applicant can edit another's record.
  * The admin dashboard has a **Duplicate registrations** tab that lists every shared email/mobile. It flags groups whose DOBs differ ("likely different persons").
* After 5 failed attempts the email/mobile is locked for 15 minutes, and every failure adds a 1-second server delay (configurable in `RATE_LIMIT`). Failure messages never say whether the email exists.
* Sessions expire after 15 minutes of inactivity (with a warning at 2 minutes, and typing in the form keeps the session alive), and after 90 minutes in any case.

## 4. Setup (about 15 minutes)

1. **Back up the spreadsheet** (File → Make a copy).
2. Open <https://script.google.com> → **New project** (or, from the spreadsheet, **Extensions → Apps Script** to bind it).
3. Create the files from this repository with the **same names**. In the editor, `.gs` files are *Script* files and `.html` files are *HTML* files (type the name without the extension). Turn on **Project Settings → Show "appsscript.json"** and paste `appsscript.json`.
   *With [clasp](https://github.com/google/clasp):* run `clasp create --type standalone` in this folder, then `clasp push` (`.claspignore` leaves out the test files).
4. Edit **`Config.gs`**:
   * `SPREADSHEET_ID`: the long ID from the sheet URL `https://docs.google.com/spreadsheets/d/`**`<ID>`**`/edit`. Leave the placeholder if the script is bound to the sheet.
   * `APPLICATION_SHEET_NAME`: the tab that holds the ~2,500 records.
   * `KEY_COLUMNS`: add your exact header text if it differs from the aliases listed (for example `'Mobile No. (Registered)'`).
   * Optional: `DRIVE.*_FOLDER_ID`. Leave them blank to have folders created. To use your own folder, copy the ID from `https://drive.google.com/drive/folders/`**`<ID>`**.
   * Optional: `UPLOAD.MAX_FILE_SIZE_MB`, `SESSION`, `RATE_LIMIT`, `REQUIRE_APPROVAL`, `PORTAL_OPEN`, `MAX_CORRECTIONS_PER_APPLICATION`, branding.
5. Select **`setupPortal`** in the function list → **Run** → authorise. Read the execution log, which reports the columns added, detected document columns, created folders, duplicate application numbers and any sharing problems.
6. Review the new **`Field_Config`** tab (see §5) and adjust it.
7. Create an admin: **Project Settings → Script Properties** → add `NEW_ADMIN_USERNAME` and `NEW_ADMIN_PASSWORD` (10+ characters) → run **`setupAdminAccount`**. The password is stored as a salted, iterated SHA-256 hash and the plain-text properties are deleted. To remove an admin, set `REMOVE_ADMIN_USERNAME` and run `removeAdminAccount`.

### Changes made to your existing sheet
Only four columns are **appended at the end**: `Last Modified`, `Correction Count`, `Correction Status`, `Last Correction ID`. Existing columns and rows are never reordered, deleted or duplicated. Two new tabs are added: `Field_Config` and the protected `Correction_Log`.

## 5. Field mapping and how document columns are detected

Nothing is hard-coded per column. Fields come from the **header row**. `setupPortal()` writes one row per header to `Field_Config`, with these settings filled in automatically:

| Column | Meaning |
|---|---|
| Column Header | Exact header text in the application sheet (the link between the two) |
| Display Label | Label shown to applicants |
| Section | Personal Information, Address Details, Educational Qualifications, Category and Reservation Details, Contact Information, Other Application Details, Uploaded Documents (or any name you choose) |
| Field Type | `text`, `textarea`, `email`, `mobile`, `date`, `number`, `select`, `pincode`, `document`, `readonly` (shown, locked), `hidden` (never shown) |
| Required / Editable | `Yes` / `No` |
| Options | Dropdown values, comma-separated (use `|` if an option contains a comma) |
| Validation Regex | Optional extra rule, e.g. `^[A-Z]{5}[0-9]{4}[A-Z]$` |
| Max Length, Help Text, Display Order | As named |

**Document detection.** A column becomes `document` when at least half of the sampled values (first 300 rows) contain `drive.google.com` / `docs.google.com` links. If the column has no data yet, the header is used instead (upload, document, certificate, photo, signature, marksheet, proof, scan…). The link parser understands `…/file/d/<id>/view`, `open?id=<id>`, `uc?id=<id>`, the comma-separated multi-file links that Google Forms produces, `=HYPERLINK()` formulas and bare file IDs.

**Other detection.** Email, mobile, date (by header or because the values are dates), PIN code, address → textarea, and short value lists for gender/category/religion/board/yes-no → `select` with the values found in the data.

**Enforced on the server no matter what Field_Config says:** the Application Number and the four system columns can never be edited.

Changes to `Field_Config` take effect within 10 minutes, or straight away via **Refresh field configuration** on the admin dashboard. New columns added to the sheet later still appear automatically. Run `setupPortal()` again to add them to `Field_Config`, and `showFieldMapping()` to print the current mapping.

## 6. Deploying as a web app

1. **Deploy → New deployment →** type **Web app**.
2. **Execute as: Me (owner)** · **Who has access: Anyone**.
3. **Deploy**, then copy the `/exec` URL.
   * Applicant portal: `https://script.google.com/macros/s/<deployment-id>/exec`
   * Admin dashboard: the same URL followed by `?page=admin`
4. After any code change: **Deploy → Manage deployments → Edit → Version: New version**. The URL stays the same.
5. Make sure the spreadsheet and the three Drive folders have **General access: Restricted**. `setupPortal()` warns if they don't.

Quotas: Apps Script allows 6 minutes per execution and about 30 simultaneous executions per user. A typical submission finishes in 1–5 seconds. The script lock waits up to 30 seconds, after which the applicant sees "server busy, please submit again".

## 7. Workflow

Login → Verify Application (list, with duplicate flags) → Edit Details (modified fields highlighted, original value shown with Undo) → Preview (original ➜ corrected, documents existing ➜ new, declaration) → Submit (confirmation dialog) → Confirmation (transaction ID, printable).

Replacing a document: choose a file. The browser checks the type and size, then the server checks the extension, MIME type, **file signature** (`%PDF` / JPEG `FFD8FF`) and size, and stores the file in the *Staging* folder. **The sheet is unchanged until final submission.** On submit, the previous file is **copied** into *Archived Documents* (or moved, with `ARCHIVE_MODE: 'MOVE'`), the new file goes to *Replacement Uploads* named `<AppNo>_<Field>_<TransactionID>.pdf`, and the cell gets the new link. **Originals are never deleted.** Unsubmitted uploads are trashed on logout or by the daily cleanup trigger, which only touches the Staging folder.

`Correction_Log` gets one row per changed field: transaction ID, application number, field, previous value, updated value, timestamp, document details (new file ID and name, archived copy ID), status and who made the change. The portal never updates or deletes log rows, and the tab is protected.

Optional approval (`REQUIRE_APPROVAL: true`): submissions get the status *Pending Approval*. The admin can **Approve**, or **Reject & revert**, which restores the previous values of that correction (only if they haven't changed since). Each decision is logged as a `[STATUS]` row.

## 8. Testing

### A. With a synthetic sample (recommended before using real data)
1. Run **`createSampleData()`**. It creates *Correction Portal – Sample Data* with **2,500 rows × 50 columns**, Drive links to sample PDF/JPG files, deliberate duplicates and imperfections. If `SPREADSHEET_ID` is not set it points the portal at this sheet automatically.
2. Run `setupPortal()` and `setupAdminAccount()`, then deploy (§6).
3. Use the logins printed in the execution log. Rows `APP202600025/26` share an email and DOB (same person, both listed as duplicates). Rows `APP202600040/41` share an email but have different DOBs (each person sees only their own).
4. To switch to the real sheet, delete the Script Property `SPREADSHEET_ID` and set `CONFIG.SPREADSHEET_ID`.

### B. Functional checklist
| # | Test | Expected |
|---|---|---|
| 1 | Correct email + DOB; correct mobile (with +91/spaces) + DOB | Application opens |
| 2 | Wrong DOB, unknown email | Generic "do not match" message, attempts remaining |
| 3 | 5 wrong attempts | Locked 15 min, even with the correct DOB |
| 4 | Duplicate email, same DOB | All matching applications listed with **Duplicate** badges |
| 5 | Change name, gender, PIN, mobile → Preview | Each shows original ➜ corrected, highlighted |
| 6 | Invalid mobile/email/PIN, empty mandatory field | Inline error, Preview blocked |
| 7 | Upload `.png`, a renamed `.exe`, a 5 MB PDF | Rejected with a clear message |
| 8 | Replace a document, view it, then submit | New link in sheet, old copy in Archive, original still exists |
| 9 | Submit | Count +1, status *Modified*, Last Modified set, log rows written, **no new row** |
| 10 | Open the same application in two tabs and submit both | Second submission gets "updated after you opened it" |
| 11 | Leave idle for 15 min | Warning at 2 min, then logged out |
| 12 | Admin: dashboard, search by app no/email/mobile, history, documents, duplicates | Data correct |

### C. Security checks
* Open the spreadsheet and folder links in a private window: access should be denied.
* In the browser console, call `google.script.run.loadApplication('<your token>', '<another app no>')`: `FORBIDDEN`.
* Call `setupPortal`, `setupAdminAccount` or `createSampleData` through `google.script.run`: refused (owner-in-editor only).
* Enter `=IMPORTXML(...)` or `<script>` in a text field: stored as plain text, shown escaped, never executed.
* View the page source: no spreadsheet ID, folder ID or other applicant data.
* Admin login with a wrong password 5 times: locked for 30 minutes.

### D. Performance checks
* With the 2,500-row sample, open **Executions** in the Apps Script dashboard. Expected: login ~1–2 s, load ~1–3 s (mostly Drive name lookups for document columns), submit ~2–6 s depending on how many documents are replaced.
* Have 10 or more testers submit at the same moment: all succeed (serialised by LockService), or a few get "busy, try again". There should be no lost updates and no duplicate rows.

### E. Offline automated tests (developers)
`local-test/` runs the **real `.gs` code** against an in-memory mock of SpreadsheetApp, DriveApp, CacheService, LockService and the other services:

```bash
node local-test/backend-test.js
```

There are 33 backend tests covering login, duplicates, rate limiting, access control, validation, uploads, archiving, the audit log, optimistic locking, admin and approval. For a full browser walkthrough with screenshots (needs Playwright):

```bash
node local-test/ui-test.js /tmp/screens
```

To test against **your own sheet**, export it as JSON (`{sheetName, values, formulas}`, dates as `{"d": ISO}`) **outside the repository** and run:

```bash
node local-test/sheet-data-test.js /path/outside/repo/sheet.json
```

It prints the detected key columns and field mapping, then checks data quality, tries login for every row (email + DOB and mobile + DOB), loads every application, submits sample corrections, and confirms that no untouched cell changed. Never commit real applicant data.

The mock is a test aid. Always finish with tests A–D on real Google infrastructure.

## 9. Troubleshooting
| Message | Fix |
|---|---|
| "missing column(s): …" | Add your header text to `CONFIG.KEY_COLUMNS` |
| "Document storage is not configured" | Run `setupPortal()` |
| "appears more than once in the records" | Two rows share an application number. Fix the sheet (the admin Duplicates tab lists them) |
| A field shows in the wrong section or type | Edit `Field_Config`, then *Refresh field configuration* |
| A document shows "Not accessible" | The owner account cannot open that Drive file. Share it with the owner, or let the applicant upload a replacement |
| Error with a reference number | Find the reference under **Executions** in the Apps Script editor to see the stack trace |
