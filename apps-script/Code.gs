const ROSTER_SHEET='Roster';
const CONFIG_SHEET='Config';
const LOG_SHEET='Scan Log';
const USED_CODES_SHEET='Used Codes';
const DEFAULT_TZ='America/Los_Angeles';
const CODE_CHARS='ABCDEFGHJKMNPQRSTUVWXYZ23456789';

function doGet(e){
  try{
    const action=String(e?.parameter?.action||'').toLowerCase();
    if(action==='member')return json_(getMemberPublic_(String(e.parameter.id||'')));
    if(action==='health')return json_({ok:true,service:'USC Flavors Attendance'});
    return json_({ok:false,error:'Unknown action'});
  }catch(err){return json_({ok:false,error:err.message||String(err)});}
}

function doPost(e){
  try{
    const body=JSON.parse(e?.postData?.contents||'{}');
    const action=String(body.action||'').toLowerCase();
    const cfg=getConfigMap_();
    requirePinFromConfig_(body.pin,cfg);

    if(action==='bootstrap'||action==='status')return json_(getStatus_(cfg));
    if(action==='startmeeting')return json_(startMeeting_(cfg));
    if(action==='closemeeting')return json_(closeMeeting_(cfg));
    if(action==='scan')return json_(recordScan_(String(body.code||''),cfg));
    if(action==='verifypin')return json_({ok:true});
    return json_({ok:false,error:'Unknown action'});
  }catch(err){return json_({ok:false,error:err.message||String(err)});}
}

function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Flavors Attendance')
    .addItem('Setup / repair sheets','setupAttendanceSystem')
    .addItem('Generate missing member codes','generateMissingCodes')
    .addItem('Generate pass links','generatePassLinks')
    .addSeparator()
    .addItem('Install 5-minute finalizer trigger','installFinalizerTrigger')
    .addItem('TEST: finalize open meeting now','testFinalizeOpenMeetingNow')
    .addToUi();
}

function setupAttendanceSystem(){
  const ss=SpreadsheetApp.getActive();
  let r=ss.getSheetByName(ROSTER_SHEET);
  if(!r)r=ss.insertSheet(ROSTER_SHEET);
  if(r.getLastRow()===0)r.getRange(1,1,1,4).setValues([['Code','Pass Link','Name','Email']]);

  let c=ss.getSheetByName(CONFIG_SHEET);
  if(!c)c=ss.insertSheet(CONFIG_SHEET);
  if(c.getLastRow()===0)c.getRange(1,1,6,2).setValues([
    ['Key','Value'],['SCANNER_PIN','2468'],['TIME_ZONE',DEFAULT_TZ],
    ['SITE_BASE_URL','https://shnliiu.github.io/usc-flavors-attendance'],
    ['OPEN_MEETING_DATE',''],['OPEN_MEETING_COLUMN','']
  ]);

  let l=ss.getSheetByName(LOG_SHEET);
  if(!l){
    l=ss.insertSheet(LOG_SHEET);
    l.getRange(1,1,1,4).setValues([['Timestamp','Meeting Date','Code','Name']]);
    l.hideSheet();
  }

  getUsedCodesSheet_();
  SpreadsheetApp.getUi().alert('Setup complete. Change SCANNER_PIN in Config if needed.');
}

function getConfigMap_(){
  const s=SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET);
  if(!s)throw new Error('Config sheet is missing. Run setupAttendanceSystem first.');
  const v=s.getDataRange().getDisplayValues(),m={};
  for(let i=1;i<v.length;i++){
    const k=String(v[i][0]||'').trim();
    if(k)m[k]=String(v[i][1]??'').trim();
  }
  return m;
}

function setConfigValues_(updates){
  const s=SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET);
  const values=s.getDataRange().getDisplayValues();
  const rows={};
  for(let i=1;i<values.length;i++)rows[String(values[i][0]||'').trim()]=i+1;
  Object.keys(updates).forEach(k=>{
    let row=rows[k];
    if(!row){row=s.getLastRow()+1;s.getRange(row,1).setValue(k);rows[k]=row;}
    s.getRange(row,2).setNumberFormat('@').setValue(String(updates[k]));
  });
}

function tzFromConfig_(cfg){return cfg.TIME_ZONE||DEFAULT_TZ;}
function todayFromConfig_(cfg){return Utilities.formatDate(new Date(),tzFromConfig_(cfg),'yyyy-MM-dd');}

function requirePinFromConfig_(pin,cfg){
  const expected=cfg.SCANNER_PIN;
  if(!expected)throw new Error('Scanner PIN is not configured.');
  if(String(pin||'')!==String(expected))throw new Error('Incorrect PIN');
}

function rosterInfo_(){
  const s=SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET);
  if(!s)throw new Error('Roster sheet is missing.');
  const lastCol=Math.max(s.getLastColumn(),1);
  const headers=s.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v=>String(v).trim());
  const col=name=>{
    const i=headers.indexOf(name);
    if(i<0)throw new Error('Roster is missing the "'+name+'" column.');
    return i+1;
  };
  return{s,headers,codeCol:col('Code'),nameCol:col('Name'),emailCol:col('Email'),passCol:headers.indexOf('Pass Link')+1};
}

function getUsedCodesSheet_(){
  const ss=SpreadsheetApp.getActive();
  let u=ss.getSheetByName(USED_CODES_SHEET);
  if(!u){
    u=ss.insertSheet(USED_CODES_SHEET);
    u.getRange(1,1,1,2).setValues([['Code','Issued At']]);
  }
  if(!u.isSheetHidden())u.hideSheet();
  return u;
}

function randomCode_(){
  let s='';
  for(let i=0;i<8;i++)s+=CODE_CHARS[Math.floor(Math.random()*CODE_CHARS.length)];
  return s;
}

function generateMissingCodes(){
  const {s,codeCol,nameCol}=rosterInfo_(),n=s.getLastRow();
  if(n<2)return;

  const u=getUsedCodesSheet_();
  const current=s.getRange(2,codeCol,n-1,1).getDisplayValues().flat()
    .map(v=>String(v||'').trim().toUpperCase()).filter(Boolean);
  const historical=u.getLastRow()>1?u.getRange(2,1,u.getLastRow()-1,1).getDisplayValues().flat()
    .map(v=>String(v||'').trim().toUpperCase()).filter(Boolean):[];

  const used=new Set(historical);
  const issued=[];
  current.forEach(code=>{
    if(!used.has(code)){used.add(code);issued.push([code,new Date()]);}
  });

  const names=s.getRange(2,nameCol,n-1,1).getDisplayValues().flat();
  const out=s.getRange(2,codeCol,n-1,1).getDisplayValues();

  for(let i=0;i<out.length;i++){
    if(!String(names[i]||'').trim()||String(out[i][0]||'').trim())continue;
    let code;
    do{code=randomCode_();}while(used.has(code));
    used.add(code);
    out[i][0]=code;
    issued.push([code,new Date()]);
  }

  s.getRange(2,codeCol,out.length,1).setValues(out);
  if(issued.length)u.getRange(u.getLastRow()+1,1,issued.length,2).setValues(issued);
}

function getStatus_(cfg){
  const {s}=rosterInfo_(),total=Math.max(s.getLastRow()-1,0);
  const d=cfg.OPEN_MEETING_DATE,col=Number(cfg.OPEN_MEETING_COLUMN||0);
  if(!d||!col)return{ok:true,open:false,date:null,checkedIn:0,total};
  const checked=total?s.getRange(2,col,total,1).getDisplayValues().flat().filter(v=>v==='Present').length:0;
  return{ok:true,open:d===todayFromConfig_(cfg),date:d,checkedIn:checked,total};
}

function startMeeting_(cfg){
  const d=todayFromConfig_(cfg);
  if(cfg.OPEN_MEETING_DATE===d&&Number(cfg.OPEN_MEETING_COLUMN)>0)return getStatus_(cfg);
  if(cfg.OPEN_MEETING_DATE)throw new Error('A previous meeting is still open.');

  const {s}=rosterInfo_();
  const col=s.getLastColumn()+1;
  s.getRange(1,col).setNumberFormat('@').setValue(d);
  if(s.getLastRow()>1)s.getRange(2,col,s.getLastRow()-1,1).clearContent().clearNote();
  applyAttendanceFormatting_(s,col);

  setConfigValues_({OPEN_MEETING_DATE:d,OPEN_MEETING_COLUMN:col});
  return{ok:true,open:true,date:d,checkedIn:0,total:Math.max(s.getLastRow()-1,0),column:col};
}

function recordScan_(raw,cfg){
  const code=String(raw||'').trim().toUpperCase();
  const {s,codeCol,nameCol}=rosterInfo_(),n=s.getLastRow();
  if(!code||n<2)return{ok:false,type:'unknown',message:'Code not recognized.'};

  const codes=s.getRange(2,codeCol,n-1,1).getDisplayValues().flat();
  const idx=codes.findIndex(v=>String(v||'').trim().toUpperCase()===code);
  if(idx<0)return{ok:false,type:'unknown',message:'Code not recognized.'};

  let meetingDate=cfg.OPEN_MEETING_DATE;
  let meetingCol=Number(cfg.OPEN_MEETING_COLUMN||0);
  const today=todayFromConfig_(cfg);

  if(!meetingDate){
    const started=startMeeting_(cfg);
    meetingDate=started.date;
    meetingCol=Number(started.column);
  }

  if(meetingDate!==today)return{ok:false,type:'closed',message:'No meeting is open for today.'};

  const row=idx+2;
  const name=s.getRange(row,nameCol).getDisplayValue().trim();
  const cell=s.getRange(row,meetingCol);

  if(cell.getDisplayValue()==='Present')
    return{ok:true,type:'already',name,message:'Already checked in.',date:meetingDate};

  const now=new Date();
  const stamp=Utilities.formatDate(now,tzFromConfig_(cfg),'yyyy-MM-dd h:mm:ss a z');
  cell.setValue('Present').setNote('Checked in: '+stamp);

  const log=SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
  if(log)log.appendRow([now,meetingDate,code,name]);

  return{ok:true,type:'success',name,date:meetingDate};
}

function closeMeeting_(cfg){
  const col=Number(cfg.OPEN_MEETING_COLUMN||0);
  if(!cfg.OPEN_MEETING_DATE||!col)throw new Error('No meeting is open.');
  const total=finalizeColumn_(col);
  return{ok:true,closed:true,status:{ok:true,open:false,date:null,checkedIn:0,total}};
}

function finalizeAbsences(){
  const cfg=getConfigMap_(),d=cfg.OPEN_MEETING_DATE,col=Number(cfg.OPEN_MEETING_COLUMN||0);
  if(!d||!col)return;
  const now=new Date(),tz=tzFromConfig_(cfg);
  const today=Utilities.formatDate(now,tz,'yyyy-MM-dd');
  const hm=Number(Utilities.formatDate(now,tz,'HHmm'));
  if(d<today||(d===today&&hm>=2359))finalizeColumn_(col);
}

function finalizeColumn_(col){
  const {s}=rosterInfo_(),total=Math.max(s.getLastRow()-1,0);
  if(total){
    const r=s.getRange(2,col,total,1);
    const v=r.getDisplayValues().map(x=>[String(x[0]).trim()?x[0]:'Absent']);
    r.setValues(v);
  }
  setConfigValues_({OPEN_MEETING_DATE:'',OPEN_MEETING_COLUMN:''});
  return total;
}

function installFinalizerTrigger(){
  ScriptApp.getProjectTriggers()
    .filter(t=>t.getHandlerFunction()==='finalizeAbsences')
    .forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('finalizeAbsences').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert('Installed.');
}

function testFinalizeOpenMeetingNow(){
  const c=Number(getConfigMap_().OPEN_MEETING_COLUMN||0);
  if(!c)throw new Error('No meeting is open.');
  finalizeColumn_(c);
  SpreadsheetApp.getUi().alert('Blank cells changed to Absent.');
}

function applyAttendanceFormatting_(s,col){
  const range=s.getRange(2,col,Math.max(s.getLastRow()-1,1),1);
  const keep=s.getConditionalFormatRules().filter(rule=>!rule.getRanges().some(r=>r.getColumn()===col&&r.getSheet().getName()===ROSTER_SHEET));
  const p=SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Present').setBackground('#D9EAD3').setFontColor('#1F5F3A').setRanges([range]).build();
  const a=SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Absent').setBackground('#F4CCCC').setFontColor('#8A1C1C').setRanges([range]).build();
  s.setConditionalFormatRules(keep.concat([p,a]));
}

function getMemberPublic_(code){
  const id=String(code||'').trim().toUpperCase();
  const {s,codeCol,nameCol}=rosterInfo_(),n=s.getLastRow();
  if(!id||n<2)return{ok:false,found:false};

  const codes=s.getRange(2,codeCol,n-1,1).getDisplayValues().flat();
  const idx=codes.findIndex(v=>String(v||'').trim().toUpperCase()===id);
  if(idx<0)return{ok:false,found:false};

  const cfg=getConfigMap_(),col=Number(cfg.OPEN_MEETING_COLUMN||0);
  const open=cfg.OPEN_MEETING_DATE===todayFromConfig_(cfg)&&col>0;
  const checkedIn=!!(open&&s.getRange(idx+2,col).getDisplayValue()==='Present');
  const name=s.getRange(idx+2,nameCol).getDisplayValue().trim();

  return{ok:true,found:true,name,meetingOpen:open,meetingDate:open?cfg.OPEN_MEETING_DATE:null,checkedIn};
}

function generatePassLinks(){
  const cfg=getConfigMap_(),base=(cfg.SITE_BASE_URL||'').replace(/\/$/,'');
  if(!base)throw new Error('Set SITE_BASE_URL in Config first.');

  const info=rosterInfo_(),s=info.s,n=s.getLastRow();
  if(n<2)return;

  let passCol=info.passCol;
  if(!passCol){
    s.insertColumnAfter(info.codeCol);
    passCol=info.codeCol+1;
    s.getRange(1,passCol).setValue('Pass Link');
  }else if(passCol!==info.codeCol+1){
    s.moveColumns(s.getRange(1,passCol,s.getMaxRows(),1),info.codeCol+1);
    passCol=info.codeCol+1;
  }

  const refreshed=rosterInfo_();
  const codes=s.getRange(2,refreshed.codeCol,n-1,1).getDisplayValues().flat();
  s.getRange(2,passCol,codes.length,1).setValues(codes.map(c=>[
    c?base+'/pass/?id='+encodeURIComponent(String(c).trim()):''
  ]));
}
