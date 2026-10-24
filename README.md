# MailFlow

Personalized bulk-email dashboard for GitHub Pages with Google Apps Script backend.

## Safety defaults
- Live email sending is disabled by default (`ENABLE_LIVE_SENDING=false`).
- Keep it disabled while testing; use preview and test data first.
- The app limits each send run to 90 recipients and requires a confirmation step.

## Project layout
- `index.html` — GitHub Pages entry page
- `app.js` — frontend logic
- `config.js.example` — frontend configuration template
- `backend/Code.gs` — Apps Script backend
- `backend/appsscript.json` — Apps Script manifest

## Publish the site
1. Open this repository's **Settings → Pages**.
2. Under Build and deployment, choose **Deploy from a branch**.
3. Choose branch `main`, folder `/ (root)`, and Save.
4. The site will be published at `https://getquicker.github.io/mailflow/` after the Pages build finishes.

## Connect Google Sheets and Apps Script
1. Create/open your Google Sheet. The default layout expects headings on row 4 and contact records from row 5. Columns: A Company, B Name, C Role, D Email, E Outreach status, F Email status.
2. Create a Google Apps Script project and paste `backend/Code.gs` into `Code.gs`; set the manifest from `backend/appsscript.json` if needed.
3. In Apps Script **Project Settings → Script properties**, add `SPREADSHEET_ID` (the ID from the Sheet URL), `MAILFLOW_ACCESS_CODE` (a long random secret), and `ENABLE_LIVE_SENDING` set to `false`.
4. Deploy as a **Web app**, execute as yourself, and set access according to your organization's needs. Copy the `/exec` URL.
5. Copy `config.js.example` to `config.js` and set `APP_SCRIPT_URL` and `ACCESS_CODE`. Do not commit a sensitive access code to a public repository; use a private repo or a safer authenticated backend for real contacts.

## Test safely
- Start with dummy contacts and sample addresses you control.
- Keep live sending disabled until the configuration and permissions are verified.
- Never publish real contact data or access codes in this repository.

## Important limitations
This is a starter tool, not an enterprise email platform. A shared access code is not per-user identity; public frontend code cannot keep secrets. For team use with real contacts, use an authenticated backend and access controls. Sending is subject to Google Apps Script quotas and Gmail policies.