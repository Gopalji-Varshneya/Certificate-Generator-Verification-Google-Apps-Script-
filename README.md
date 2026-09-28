# Certificate Generator & Verification (Google Apps Script)

Bulk-generate PDF certificates from a Google Sheet and a Google Slides template, email them to recipients, and let anyone verify a certificate by scanning its QR code or by looking it up with their email address.

Built with **Google Apps Script** only. No server, no paid services, no Google Cloud project needed.

---

## Features

- **Bulk generation** from a Google Sheet, using a Google Slides certificate template
- **Unique certificate IDs** (`CERT-XXXXXXXX`) created automatically
- **QR code on every certificate** that opens a verification page
- **Public verification page** showing whether a certificate is valid
- **Find-my-certificate page**: recipients enter their email and get their certificate link
- **Batch email sending** with the PDF attached
- **Optimized for speed**: QR codes and PDF exports are fetched in parallel, sheet writes are batched, and long runs stop safely before the Apps Script time limit

---

## How it works

```
Google Sheet (Students)
        |
        v
  generateBatch()  --->  copies Slides template, fills placeholders,
        |                inserts QR code, exports PDF to Drive
        v
  PDF_URL / Verify_URL / Status written back to the Sheet
        |
        v
  sendEmailBatch() --->  emails the PDF to each recipient

Web app (doGet):
   /exec            -> Index.html (find certificate by email)
   /exec?id=CERT-x  -> verification page (opened by scanning the QR)
```

---

## Repository files

| File | Purpose |
|---|---|
| `Code.js` | All backend logic: ID generation, certificate generation, email, verification |
| `Index.html` | Front-end page for looking up a certificate by email |
| `README.md` | This file |

---

## Setup

### 1. Prepare the Google Sheet

Create a spreadsheet with a tab named **`Students`** and this exact header row (column names are case-sensitive):

| Certificate_ID | Team | Role | Name | Email | Stream | Academic_Year | IssueDate | PDF_URL | Verify_URL | Status | Email_Status |
|---|---|---|---|---|---|---|---|---|---|---|---|

- Fill in `Name`, `Email`, `Stream` and `Academic_Year` for each person.
- Leave `Certificate_ID`, `PDF_URL`, `Verify_URL`, `Status` and `Email_Status` empty. The script fills them.
- `IssueDate` is optional. If it is blank, today's date is used.

### 2. Create the Slides certificate template

Create a Google Slides presentation with **one slide** and add these placeholders as text (one per text box):

| Placeholder | Replaced with |
|---|---|
| `{{NAME}}` | Recipient name |
| `{{COURSE}}` | Stream |
| `{{Academic_Year}}` | Academic year |
| `{{DATE}}` | Issue date |
| `{{CERTIFICATE_ID}}` | Certificate ID |
| `{{QR}}` | QR code image (the text box is replaced by the QR, so size the box the way you want the QR to appear) |

Tips to avoid overlapping text:
- Give each placeholder its own text box with clear spacing.
- In **Format options > Text fitting**, choose **Shrink text on overflow**.
- Make the `{{NAME}}` box wide enough for your longest name.
- Do not leave trailing spaces after a placeholder in a centered box.

### 3. Create a Drive output folder

Create an empty folder in Google Drive. Certificates are saved here.

### 4. Create the Apps Script project

1. Open [script.google.com](https://script.google.com) and create a new project.
2. Add the two files from this repo:
   - `Code.js` (the default `Code.gs` file)
   - `Index.html` (**File +** > **HTML**, name it exactly `Index`)
3. In `Code.js`, replace every `xxx` in `CONFIG`:

```js
const CONFIG = {
  SPREADSHEET_ID:     'xxx',  // from the sheet URL: /spreadsheets/d/<THIS PART>/edit
  STUDENT_SHEET:      'Students',
  SLIDES_TEMPLATE_ID: 'xxx',  // from the Slides URL: /presentation/d/<THIS PART>/edit
  OUTPUT_FOLDER_ID:   'xxx',  // from the Drive folder URL: /folders/<THIS PART>
  VERIFY_URL:         'https://script.google.com/macros/s/xxx/exec', // filled in after step 5
  BATCH_SIZE:         1,      // raise after testing (see Usage)
  SEND_EMAIL:         true
};
```

### 5. Deploy the web app

1. Click **Deploy > New deployment > Web app**.
2. Set **Execute as**: *Me*.
3. Set **Who has access**: **Anyone**.
   - If it is set to your organization only, people outside it will be sent to a Google sign-in page and QR verification will fail for them.
4. Click **Deploy**, authorize the permissions, and copy the **Web app URL** (ends in `/exec`).
5. Paste that URL into `CONFIG.VERIFY_URL` in `Code.js`.
6. **Deploy > Manage deployments > Edit > Version: New version > Deploy** so the change goes live.

> Important: the QR codes are generated from `VERIFY_URL`. If that URL is wrong or outdated, every QR code will point to a dead link.

### 6. First run

1. In the Apps Script editor, run `generateMissingIDs` once and approve the permission prompts.
2. Run `testOneStudent` and check the sheet: the first row should get a `PDF_URL` and `Status = GENERATED`.
3. Open the PDF in Drive and check the layout.
4. Scan the QR code, or open `<VERIFY_URL>?id=<Certificate_ID>`, to confirm it shows **CERTIFICATE VALID**.

---

## Usage

If the script is bound to your sheet, a **Certificates** menu appears with:

| Menu item | Function | What it does |
|---|---|---|
| Test / Generate next certificate | `generateBatch` | Generates certificates for pending rows |
| Send next email batch | `sendEmailBatch` | Emails generated certificates (20 per run by default) |
| Generate missing Certificate IDs | `generateMissingIDs` | Assigns IDs to rows that have none |

**Generating many certificates**

1. After testing with `BATCH_SIZE: 1`, raise it (for example to `25` or `50`).
2. Run `generateBatch`. Each run stops safely before the Apps Script time limit and logs how many are left. Run it again to continue; finished rows are skipped.
3. Run `sendEmailBatch` to email them. Rows marked `SENT` are skipped.

**Regenerating a certificate:** clear `PDF_URL` (and `Status`) on that row and run `generateBatch` again.

---

## Status values

| Column | Value | Meaning |
|---|---|---|
| `Status` | `PROCESSING` | Being generated |
| `Status` | `GENERATED` | Ready and valid |
| `Status` | `ERROR: ...` | Failed; it is retried on the next run |
| `Email_Status` | `SENT` | Email delivered to the mail service |
| `Email_Status` | `ERROR: ...` | Failed; retried on the next run |

Only rows with `Status = GENERATED` are shown as valid on the verification page.

---

## Troubleshooting

| Problem | Likely cause and fix |
|---|---|
| **"Page not found"** on the web app URL | The deployment ID is stale. In **Manage deployments**, copy the current `/exec` URL. |
| **QR code opens a dead or sign-in page** | `VERIFY_URL` does not match the live public deployment, or access is not set to **Anyone**. Fix it, redeploy, then clear `PDF_URL` on affected rows and regenerate. |
| **Placeholder text appears on the certificate** (for example `{{Academic_Year}}`) | The text in the slide does not match the script exactly (check for `}]` instead of `}}`). |
| **"Check your column headings"** | A header is misspelled. Use the exact names from Setup step 1, including `Email_Status` with an underscore. |
| **`QR placeholder {{QR}} was not found`** | The `{{QR}}` box is missing from slide 1, or it contains extra text. |
| **Code changes have no effect on the live URL** | Redeploy: **Manage deployments > Edit > New version**. |
| **Emails stop sending** | The daily Apps Script mail quota is used up. Try again the next day. |

---

## Security and privacy notes

- **Never commit your real IDs.** Keep `xxx` placeholders in the repo, and put real values only in your own Apps Script project.
- The web app runs as you and is public by design. The email lookup page returns the name, stream, issue date, ID and PDF link for any email address that exists in the sheet, so consider whether that suits your data.
- Certificate PDFs are shared as "anyone with the link can view".

---

## License

MIT license.
