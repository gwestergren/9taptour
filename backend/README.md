# 9 Tap Tour data service

The website keeps Tour data in the Google Sheet **9 Tap Tour Data** (Greg's Drive, folder `9taptour`).
`Code.gs` is a small Google Apps Script web app that sits inside that sheet: the site reads data from it and sends
check-ins, scores, Setup, venues, member edits, pots and finalized results to it.

## Install it (about 5 minutes, once)

1. Open the sheet **9 Tap Tour Data**. In the menu choose **Extensions → Apps Script**. A new tab opens.
2. At the top left, click **Untitled project** and rename it **9 Tap Tour API**.
3. In `Code.gs`, select everything and delete it. Paste in the whole of `backend/Code.gs` from this repo. Press **Ctrl+S**.
4. In the toolbar, pick **setup** in the function list (next to Debug), then click **Run**.
   - Google asks for permission: **Review permissions** → pick your Google account → **Advanced** →
     **Go to 9 Tap Tour API (unsafe)** → **Allow**. (It says "unsafe" only because the script is yours and not published by Google.)
   - The Execution log at the bottom shows `Staff key: ...`. Keep it private, like a password. Don't paste it in chat or email.
     It will go into the website's private Azure settings when staff sign-in is turned on.
5. Click **Deploy → New deployment**. Click the gear next to "Select type" and choose **Web app**.
   - Description: `v1`
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, then **Copy** the **Web app URL** (it ends in `/exec`).
6. Test it: paste the URL in a browser and add `?ping=1` to the end. You should see `{"ok":true,...}`.
   Add `?tabs=Pots` instead and you should see the pots from the sheet.

## Updating the script later

Paste the new code, press **Ctrl+S**, then **Deploy → Manage deployments** → pencil icon →
Version: **New version** → **Deploy**. The web app URL stays the same.

## What it allows

- Anyone can read the public tabs. Phone and Email columns are never sent out.
- Anyone can add a check-in (bowlers check in from their phones).
- Everything else needs the staff key. If it ever leaks, run **newStaffKey** once and the old key stops working.
