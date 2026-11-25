# USC Flavors Firebase migration

## Goal
Firebase/Firestore becomes the live attendance backend. GitHub Pages stays the frontend for now. Google Sheets becomes a reporting mirror and is no longer required for a scan to succeed.

## Collections
- `members/{code}`: private roster record: name, email, active, passLink
- `publicMembers/{code}`: member-facing record: name, active
- `usedCodes/{code}`: historical issued-code registry
- `officers/{uid}`: authorized Firebase Authentication user IDs
- `meetings/{meetingId}`: meeting metadata
- `meetings/{meetingId}/attendance/{code}`: Present/Absent + timestamp
- `publicState/current`: current meeting id/label/open state

## Migration sequence
1. Create Firebase project on Spark.
2. Register a Web app and copy its Firebase config into `firebase-config.js`.
3. Create Firestore in Production mode.
4. Publish `firestore.rules`.
5. Enable Google Authentication.
6. Sign in once with each officer Google account, then add an `officers/{uid}` document for each approved officer.
7. Import the current roster. Existing member codes remain unchanged. Each row creates:
   - private `members/{code}`
   - public `publicMembers/{code}`
   - permanent `usedCodes/{code}`
8. Point scanner/pass pages to `firebase-db.js`.
9. Test with a small group.
10. Switch production traffic to Firestore.
11. Keep Apps Script only as a non-critical Sheet synchronization/export layer.

## Sheet behavior after cutover
The Sheet can remain shared with members, but it is no longer the database. Attendance sync can run after scans or on a schedule. If Google Sheets/Apps Script has an outage, scanning and member passes continue to work; the Sheet catches up later.

## Security
Do not put email addresses in `publicMembers`. The Firebase web configuration is intentionally public; Firestore Security Rules and Firebase Authentication protect the data.
