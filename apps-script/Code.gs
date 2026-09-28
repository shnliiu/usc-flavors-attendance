const ROSTER_SHEET='Roster';
const CONFIG_SHEET='Config';
const LOG_SHEET='Scan Log';
const DEFAULT_TZ='America/Los_Angeles';
const CODE_CHARS='ABCDEFGHJKMNPQRSTUVWXYZ23456789';

function doGet(e){
  try{
    const action=String(e?.parameter?.action||'').toLowerCase();
    if(action==='member') return json_(getMemberPublic_(String(e.parameter.id||'')));
    if(action==='health') return json_({ok:true,service:'USC Flavors Attendance'});
    return json_({ok:false,error:'Unknown action'});
  }catch(err){return json_({ok:false,error:err.message||String(err)});}
}
function doPost(e){
  try{
    const body=JSON.parse(e?.postData?.contents||'{}');
    const action=String(body.action||'').toLowerCase();
    if(action==='verifypin'){requirePin_(body.pin);return json_({ok:true});}
    if(action==='status'){requirePin_(body.pin);return json_(getStatus_());}
    if(action==='startmeeting'){requirePin_(body.pin);return json_(startMeeting_());}
    if(action==='closemeeting'){requirePin_(body.pin);return json_(closeMeeting_());}
    if(action==='scan'){requirePin_(body.pin);return json_(recordScan_(String(body.code||'')));}
    return json_({ok:false,error:'Unknown action'});
  }catch(err){return json_({ok:false,error:err.message||String(err)});}
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function onOpen(){SpreadsheetApp.getUi().createMenu('Flavors Attendance')
.addItem('Setup / repair sheets','setupAttendanceSystem')
.addItem('Generate missing member codes','generateMissingCodes')
.addItem('Generate pass links','generatePassLinks')
.addSeparator().addItem('Install 5-minute finalizer trigger','installFinalizerTrigger')
.addItem('TEST: finalize open meeting now','testFinalizeOpenMeetingNow').addToUi();}
function setupAttendanceSystem(){
  const ss=SpreadsheetApp.getActive();
  let r=ss.getSheetByName(ROSTER_SHEET); if(!r) r=ss.insertSheet(ROSTER_SHEET);
  if(r.getLastRow()===0) r.getRange(1,1,1,5).setValues([['Code','First Name','Last Name','Grade','Major']]);
  let c=ss.getSheetByName(CONFIG_SHEET); if(!c)c=ss.insertSheet(CONFIG_SHEET);
  if(c.getLastRow()===0)c.getRange(1,1,6,2).setValues([
    ['Key','Value'],['SCANNER_PIN','2468'],['TIME_ZONE',DEFAULT_TZ],
    ['SITE_BASE_URL','https://shnliiu.github.io/usc-flavors-attendance'],
    ['OPEN_MEETING_DATE',''],['OPEN_MEETING_COLUMN','']
  ]);
  let l=ss.getSheetByName(LOG_SHEET); if(!l){l=ss.insertSheet(LOG_SHEET);l.getRange(1,1,1,5).setValues([['Timestamp','Meeting Date','Code','First Name','Last Name']]);l.hideSheet();}
  SpreadsheetApp.getUi().alert('Setup complete. Change SCANNER_PIN in Config.');
}
function getConfigMap_(){const s=SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET);if(!s)throw new Error('Config sheet is missing. Run setupAttendanceSystem first.');const v=s.getDataRange().getDisplayValues(),m={};for(let i=1;i<v.length;i++){const k=String(v[i][0]||'').trim();if(k)m[k]=String(v[i][1]??'').trim();}return m;}
function setConfig_(k,val){const s=SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET),v=s.getDataRange().getDisplayValues();for(let i=1;i<v.length;i++){if(String(v[i][0]).trim()===k){const cell=s.getRange(i+1,2);cell.setNumberFormat('@');cell.setValue(String(val));return;}}const row=s.getLastRow()+1;s.getRange(row,1).setValue(k);s.getRange(row,2).setNumberFormat('@').setValue(String(val));}
function tz_(){return getConfigMap_().TIME_ZONE||DEFAULT_TZ;}
function today_(){return Utilities.formatDate(new Date(),tz_(),'yyyy-MM-dd');}
function requirePin_(pin){const e=getConfigMap_().SCANNER_PIN;if(!e)throw new Error('Scanner PIN is not configured.');if(String(pin||'')!==String(e))throw new Error('Incorrect PIN');}
function randomCode_(){let s='';for(let i=0;i<8;i++)s+=CODE_CHARS[Math.floor(Math.random()*CODE_CHARS.length)];return s;}
function generateMissingCodes(){
  const s=SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET),n=s.getLastRow();if(n<2)return;
  const existing=new Set(s.getRange(2,1,n-1,1).getDisplayValues().flat().filter(Boolean));
  const names=s.getRange(2,2,n-1,2).getDisplayValues(),out=s.getRange(2,1,n-1,1).getDisplayValues();
  for(let i=0;i<out.length;i++){if(!(names[i][0]||names[i][1])||out[i][0])continue;let c;do{c=randomCode_();}while(existing.has(c));existing.add(c);out[i][0]=c;}
  s.getRange(2,1,out.length,1).setValues(out);
}
function startMeeting_(){
  const cfg=getConfigMap_(),d=today_();if(cfg.OPEN_MEETING_DATE===d&&Number(cfg.OPEN_MEETING_COLUMN)>0)return getStatus_();
  if(cfg.OPEN_MEETING_DATE)throw new Error('A previous meeting is still open.');
  const s=SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET),col=Math.max(s.getLastColumn(),5)+1;
  s.getRange(1,col).setValue(d).setNumberFormat('@');if(s.getLastRow()>1)s.getRange(2,col,s.getLastRow()-1,1).clearContent().clearNote();
  applyAttendanceFormatting_(s,col);setConfig_('OPEN_MEETING_DATE',d);setConfig_('OPEN_MEETING_COLUMN',col);return getStatus_();
}
function recordScan_(raw){
  const code=String(raw||'').trim().toUpperCase(),s=SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET),n=s.getLastRow();
  if(!code||n<2)return{ok:false,type:'unknown',message:'Code not recognized.'};
  const codes=s.getRange(2,1,n-1,1).getDisplayValues().flat(),idx=codes.findIndex(c=>String(c).trim().toUpperCase()===code);
  if(idx<0)return{ok:false,type:'unknown',message:'Code not recognized.'};
  let cfg=getConfigMap_();if(!cfg.OPEN_MEETING_DATE)startMeeting_();cfg=getConfigMap_();
  if(cfg.OPEN_MEETING_DATE!==today_())return{ok:false,type:'closed',message:'No meeting is open for today.'};
  const row=idx+2,col=Number(cfg.OPEN_MEETING_COLUMN),first=s.getRange(row,2).getDisplayValue().trim(),last=s.getRange(row,3).getDisplayValue().trim(),cell=s.getRange(row,col);
  if(cell.getDisplayValue()==='Present')return{ok:true,type:'already',firstName:first,lastName:last,message:'Already checked in.',status:getStatus_()};
  const now=new Date(),stamp=Utilities.formatDate(now,tz_(),'yyyy-MM-dd h:mm:ss a z');
  cell.setValue('Present').setNote('Checked in: '+stamp);
  SpreadsheetApp.getActive().getSheetByName(LOG_SHEET).appendRow([now,cfg.OPEN_MEETING_DATE,code,first,last]);
  return{ok:true,type:'success',firstName:first,lastName:last,status:getStatus_()};
}
function getStatus_(){
  const s=SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET),cfg=getConfigMap_(),total=Math.max(s.getLastRow()-1,0);
  if(!cfg.OPEN_MEETING_DATE||!cfg.OPEN_MEETING_COLUMN)return{ok:true,open:false,date:null,checkedIn:0,total};
  const col=Number(cfg.OPEN_MEETING_COLUMN),checked=total?s.getRange(2,col,total,1).getDisplayValues().flat().filter(v=>v==='Present').length:0;
  return{ok:true,open:cfg.OPEN_MEETING_DATE===today_(),date:cfg.OPEN_MEETING_DATE,checkedIn:checked,total};
}
function closeMeeting_(){
  const cfg=getConfigMap_(),col=Number(cfg.OPEN_MEETING_COLUMN||0);if(!cfg.OPEN_MEETING_DATE||!col)throw new Error('No meeting is open.');
  finalizeColumn_(col);
  return{ok:true,closed:true,date:cfg.OPEN_MEETING_DATE,status:getStatus_()};
}
function finalizeAbsences(){
  const cfg=getConfigMap_(),d=cfg.OPEN_MEETING_DATE,col=Number(cfg.OPEN_MEETING_COLUMN||0);if(!d||!col)return;
  const now=new Date(),today=Utilities.formatDate(now,tz_(),'yyyy-MM-dd'),hm=Number(Utilities.formatDate(now,tz_(),'HHmm'));
  if(d<today||(d===today&&hm>=2359))finalizeColumn_(col);
}
function finalizeColumn_(col){
  const s=SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET),total=Math.max(s.getLastRow()-1,0);
  if(total){const r=s.getRange(2,col,total,1),v=r.getDisplayValues().map(x=>[String(x[0]).trim()?x[0]:'Absent']);r.setValues(v);}
  setConfig_('OPEN_MEETING_DATE','');setConfig_('OPEN_MEETING_COLUMN','');
}
function installFinalizerTrigger(){ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='finalizeAbsences').forEach(t=>ScriptApp.deleteTrigger(t));ScriptApp.newTrigger('finalizeAbsences').timeBased().everyMinutes(5).create();SpreadsheetApp.getUi().alert('Installed.');}
function testFinalizeOpenMeetingNow(){const c=Number(getConfigMap_().OPEN_MEETING_COLUMN||0);if(!c)throw new Error('No meeting is open.');finalizeColumn_(c);SpreadsheetApp.getUi().alert('Blank cells changed to Absent.');}
function applyAttendanceFormatting_(s,col){
  const range=s.getRange(2,col,Math.max(s.getMaxRows()-1,1),1),keep=s.getConditionalFormatRules().filter(rule=>!rule.getRanges().some(r=>r.getColumn()===col&&r.getSheet().getName()===ROSTER_SHEET));
  const p=SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Present').setBackground('#D9EAD3').setFontColor('#1F5F3A').setRanges([range]).build();
  const a=SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Absent').setBackground('#F4CCCC').setFontColor('#8A1C1C').setRanges([range]).build();
  s.setConditionalFormatRules(keep.concat([p,a]));
}
function getMemberPublic_(code){
  const id=String(code||'').trim().toUpperCase(),s=SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET);if(!id||s.getLastRow()<2)return{ok:false,found:false};
  const rows=s.getRange(2,1,s.getLastRow()-1,3).getDisplayValues();
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    if(String(r[0]).trim().toUpperCase()===id){
      const cfg=getConfigMap_(),col=Number(cfg.OPEN_MEETING_COLUMN||0),open=cfg.OPEN_MEETING_DATE===today_()&&col>0;
      const checkedIn=open&&s.getRange(i+2,col).getDisplayValue()==='Present';
      return{ok:true,found:true,firstName:String(r[1]).trim(),lastName:String(r[2]).trim(),meetingOpen:open,meetingDate:open?cfg.OPEN_MEETING_DATE:null,checkedIn};
    }
  }
  return{ok:false,found:false};
}
function generatePassLinks(){
  const cfg=getConfigMap_(),base=(cfg.SITE_BASE_URL||'').replace(/\/$/,'');if(!base)throw new Error('Set SITE_BASE_URL in Config first.');
  const s=SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET),n=s.getLastRow();if(n<2)return;
  const h=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0];let col=h.indexOf('Pass Link')+1;if(!col){col=s.getLastColumn()+1;s.getRange(1,col).setValue('Pass Link');}
  const codes=s.getRange(2,1,n-1,1).getDisplayValues().flat();s.getRange(2,col,codes.length,1).setValues(codes.map(c=>[c?base+'/pass/?id='+encodeURIComponent(c):'']));
}