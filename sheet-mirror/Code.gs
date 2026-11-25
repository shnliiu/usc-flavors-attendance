const SPREADSHEET_ID = '1K2o0ep9P9DnOF2vrBBRSgfwU1tE3OiCOvoTQW12dpNc';
const ROSTER_SHEET = 'Roster';
const FIREBASE_API_KEY = ['AIzaSyDQX_','PZdxq4Ag1cJ6LS7Ay-zLSm-lsEvLs'].join('');
const OFFICER_UIDS = new Set(['pEXSlBY1L8QdiAaZTNgwooMVGGQ2']);

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const uid = verifyFirebaseToken_(body.idToken);
    if (!OFFICER_UIDS.has(uid)) return json_({ok:false,error:'Unauthorized'}, 403);

    if (body.action === 'startMeeting') return json_(mirrorStartMeeting_(body));
    if (body.action === 'scan') return json_(mirrorScan_(body));
    if (body.action === 'closeMeeting') return json_(mirrorCloseMeeting_(body));
    return json_({ok:false,error:'Unknown action'}, 400);
  } catch (err) {
    return json_({ok:false,error:String(err && err.message || err)}, 500);
  }
}

function verifyFirebaseToken_(idToken) {
  if (!idToken) throw new Error('Missing Firebase ID token');
  const url = 'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + encodeURIComponent(FIREBASE_API_KEY);
  const r = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({idToken:idToken}),
    muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) throw new Error('Invalid Firebase session');
  const data = JSON.parse(r.getContentText());
  const user = data.users && data.users[0];
  if (!user || !user.localId) throw new Error('Firebase user not found');
  return user.localId;
}

function roster_() {
  const s = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ROSTER_SHEET);
  if (!s) throw new Error('Roster sheet not found');
  return s;
}

function meetingColumn_(sheet, meetingId, createIfMissing) {
  const lastCol = Math.max(sheet.getLastColumn(), 4);
  const headers = sheet.getRange(1,1,1,lastCol).getDisplayValues()[0];
  for (let i = 4; i < headers.length; i++) if (headers[i] === meetingId) return i + 1;
  if (!createIfMissing) return 0;
  const col = Math.max(5, lastCol + 1);
  sheet.getRange(1,col).setValue(meetingId);
  return col;
}

function rowForCode_(sheet, code) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const codes = sheet.getRange(2,1,lastRow-1,1).getDisplayValues().flat();
  const idx = codes.indexOf(String(code || '').trim().toUpperCase());
  return idx < 0 ? 0 : idx + 2;
}

function mirrorStartMeeting_(body) {
  const meetingId = String(body.meetingId || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(meetingId)) throw new Error('Invalid meeting id');
  const s = roster_();
  const col = meetingColumn_(s, meetingId, true);
  return {ok:true,column:col};
}

function mirrorScan_(body) {
  const meetingId = String(body.meetingId || '').trim();
  const code = String(body.code || '').trim().toUpperCase();
  if (!meetingId || !code) throw new Error('Missing meeting or code');
  const s = roster_();
  const col = meetingColumn_(s, meetingId, true);
  const row = rowForCode_(s, code);
  if (!row) throw new Error('Member code not found in Sheet');
  const cell = s.getRange(row,col);
  cell.setValue('Present');
  cell.setNote('Synced from Firebase at ' + new Date().toISOString());
  return {ok:true,row:row,column:col};
}

function mirrorCloseMeeting_(body) {
  const meetingId = String(body.meetingId || '').trim();
  if (!meetingId) throw new Error('Missing meeting id');
  const s = roster_();
  const col = meetingColumn_(s, meetingId, true);
  const lastRow = s.getLastRow();
  if (lastRow >= 2) {
    const range = s.getRange(2,col,lastRow-1,1);
    const vals = range.getValues();
    for (let i=0;i<vals.length;i++) if (!vals[i][0]) vals[i][0] = 'Absent';
    range.setValues(vals);
  }
  return {ok:true,column:col};
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
