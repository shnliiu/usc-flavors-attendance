# Firebase → Google Sheet mirror

This Apps Script is intentionally **not** the live attendance backend.

Firestore records attendance first. The scanner then sends a background mirror event here. If Apps Script or Sheets is unavailable, the scan still succeeds in Firebase and the member pass still updates.

## Install
1. Open the existing attendance Google Sheet.
2. Extensions → Apps Script.
3. Replace the Apps Script code with `Code.gs` from this folder.
4. Deploy → New deployment → Web app.
5. Execute as: Me.
6. Who has access: Anyone.
7. Copy the new `/exec` URL.
8. Put that URL into the scanner's `MIRROR_URL`.

The endpoint verifies the caller's Firebase ID token and only accepts authorized officer UID(s).
