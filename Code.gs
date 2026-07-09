// Yasou - לוגיקת 3 גלים נסתרת מהלקוח
// הלקוח רואה רק שעות, המערכת ממפה: 20:00/20:05/20:10=covered, 20:30/20:35=inside, 21:00+=outside
const SHEET_NAME = "reservations";
const DRIVE_FOLDER_NAME = "Yasou Receipts";
const HEADER = ["id","customerName","phone","date","time","guests","tableId","status","notes","createdAt","source","reservationType","depositAmount","receiptUrl","agree","zone"];

function getSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){ sh = ss.insertSheet(SHEET_NAME); sh.getRange(1,1,1,HEADER.length).setValues([HEADER]); }
  if(sh.getLastRow()===0) sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
  if(sh.getLastColumn() < HEADER.length) sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
  return sh;
}
function getOrCreateFolder_(){ const it=DriveApp.getFoldersByName(DRIVE_FOLDER_NAME); return it.hasNext()?it.next():DriveApp.createFolder(DRIVE_FOLDER_NAME); }
function parseBody_(e){ try{ if(e.postData&&e.postData.contents) return JSON.parse(e.postData.contents);}catch(e){} return e.parameter||{}; }
function jsonResponse_(obj, cb){ const j=JSON.stringify(obj); if(cb) return ContentService.createTextOutput(`${cb}(${j})`).setMimeType(ContentService.MimeType.JAVASCRIPT); return ContentService.createTextOutput(j).setMimeType(ContentService.MimeType.JSON); }
function zoneByTime_(t){ if(["20:00","20:05","20:10"].indexOf(t)>-1) return "covered"; if(["20:30","20:35"].indexOf(t)>-1) return "inside"; return "outside"; }

function doGet(e){
  const action=(e.parameter.action||"").toLowerCase(), cb=e.parameter.callback||"";
  const sh=getSheet_();
  const rows = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1, HEADER.length).getValues() : [];
  if(action==="getreservations"){
    const reservations = rows.map(r=>{ const o={}; HEADER.forEach((h,i)=>o[h]=r[i]); o.guests=Number(o.guests)||0; o.tableId=Number(o.tableId)||0; return o; }).filter(r=>r.id);
    return jsonResponse_({ok:true, reservations}, cb);
  }
  if(action==="getavailability"){
    const date=e.parameter.date; if(!date) return jsonResponse_({ok:false},cb);
    const capacity={covered:62, inside:66, outside:40};
    const used={covered:0, inside:0, outside:0};
    rows.forEach(r=>{ const d=String(r[3]), status=String(r[7]), zone=String(r[15]||zoneByTime_(String(r[4]))), guests=Number(r[5])||0; if(d===date && status!=="cancelled" && used[zone]!==undefined) used[zone]+=guests; });
    const available={ covered:capacity.covered-used.covered, inside:capacity.inside-used.inside, outside:capacity.outside-used.outside };
    return jsonResponse_({ok:true, used, available, capacity}, cb);
  }
  return jsonResponse_({ok:true, msg:"Yasou 3-waves hidden running"}, cb);
}

function doPost(e){
  const action=(e.parameter.action||"").toLowerCase();
  const data=parseBody_(e);
  const sh=getSheet_();
  if(action==="addpublicorder"){
    let receiptUrl="";
    try{
      if(data.receiptBase64 && data.receiptName){
        const folder=getOrCreateFolder_();
        const bytes=Utilities.base64Decode(data.receiptBase64);
        const blob=Utilities.newBlob(bytes, data.receiptMime||"image/jpeg", `${data.date}_${data.time}_${data.customerName}_${data.receiptName}`.replace(/[^\w.\-]/g,"_"));
        const file=folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); receiptUrl=file.getUrl();
      }
    }catch(err){ receiptUrl="error:"+err; }
    let zone = String(data.zone||"").toLowerCase();
    if(!zone || ["covered","inside","outside"].indexOf(zone)===-1) zone = zoneByTime_(String(data.time||""));
    const row=[ String(data.id||Utilities.getUuid()), String(data.customerName).slice(0,80), String(data.phone).slice(0,30), String(data.date), String(data.time), Number(data.guests)||2, 0, "pending_public", String(data.notes||`[${zone}] ${data.time}`).slice(0,320), data.createdAt||new Date().toISOString(), "public_site", String(data.reservationType||"private"), Number(data.depositAmount)||0, receiptUrl, data.agree?"yes":"no", zone ];
    sh.appendRow(row);
    return jsonResponse_({ok:true, zone, receiptUrl});
  }
  if(action==="addreservation"||action==="updatereservationstatus"||action==="cancelreservation"){
    const id=String(data.id||""); if(!id) return jsonResponse_({ok:false});
    const ids = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat().map(String) : [];
    const idx=ids.indexOf(id);
    if(action==="cancelreservation"&&idx>=0){ sh.getRange(idx+2,8).setValue("cancelled"); return jsonResponse_({ok:true}); }
    if(action==="updatereservationstatus"&&idx>=0){ sh.getRange(idx+2,8).setValue(String(data.status||"")); if(data.tableId!==undefined) sh.getRange(idx+2,7).setValue(Number(data.tableId)); if(data.zone) sh.getRange(idx+2,16).setValue(String(data.zone)); return jsonResponse_({ok:true}); }
    const row = HEADER.map(h=> data[h]!==undefined?data[h]:""); if(!row[15]&&data.time) row[15]=zoneByTime_(String(data.time));
    if(idx>=0) sh.getRange(idx+2,1,1,HEADER.length).setValues([row]); else sh.appendRow(row);
    return jsonResponse_({ok:true});
  }
  return jsonResponse_({ok:false, error:"unknown"});
}
