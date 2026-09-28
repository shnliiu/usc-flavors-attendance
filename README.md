# USC Flavors Attendance

A free QR attendance system using **Google Sheets + Google Apps Script + GitHub Pages**.

- `/` = president scanner
- `/pass/?id=CODE` = member pass
- `/apps-script/Code.gs` = backend to paste into the Google Sheet's Apps Script project

## One-time setup

1. In your Google Sheet, rename the member tab to **Roster**. First five columns must be:
   `Code | First Name | Last Name | Grade | Major`
2. Open **Extensions → Apps Script**, paste `apps-script/Code.gs`, save, then run `setupAttendanceSystem` once.
3. In the new **Config** tab, change `SCANNER_PIN`. Keep `TIME_ZONE` = `America/Los_Angeles`.
4. Use **Flavors Attendance → Generate missing member codes**.
5. Use **Flavors Attendance → Install 5-minute finalizer trigger** once.
6. Apps Script: **Deploy → New deployment → Web app**. Execute as **Me**; access **Anyone**. Copy the URL ending in `/exec`.
7. Replace `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` in both `index.html` and `pass/index.html` with that `/exec` URL.
8. GitHub: **Settings → Pages → Deploy from a branch → main → /(root) → Save**.
9. Your scanner URL will be `https://shnliiu.github.io/usc-flavors-attendance/`.
10. Confirm Config `SITE_BASE_URL` is `https://shnliiu.github.io/usc-flavors-attendance`, then run **Generate pass links**.

## Meeting behavior

A meeting is created when the president taps **Start meeting**, or when the first **valid** member code is scanned that day. Unknown/random codes never create a column. Successful scans write `Present`, add a timestamp note, and append to hidden `Scan Log`. Duplicate scans are harmless. Conditional formatting handles green Present / red Absent.

## Test plan

Add 5 fake members; generate codes; start a meeting; scan one member; scan the same member twice; enter a bad code; confirm Present is green and no bad write occurs. To test the cutoff without waiting until midnight, use **Flavors Attendance → TEST: finalize open meeting now** and confirm blank cells become `Absent`.

## Known limitations

Apps Script quotas apply. The scanner PIN is remembered in that device's local storage. Member-pass URLs are possession-based: anyone who gets a member's private URL can display that QR. Camera permission can be revoked, so the scanner includes manual code entry.
