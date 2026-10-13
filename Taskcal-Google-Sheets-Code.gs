/**
 * TASKCAL GOOGLE SHEETS BACKEND
 *
 * 1. Create a Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Replace the default code with this file.
 * 4. Save.
 * 5. Deploy -> New deployment -> Web app.
 *    Execute as: Me
 *    Who has access: Anyone with the link
 * 6. Copy the /exec URL and paste it into Taskcal using "☁ Google Sheets".
 *
 * Sheet tab used: Tasks
 *
 * Columns:
 * ID | Title | Date | Time | Priority | Category | Repeat | End |
 * Notes | DoneJSON | SkippedJSON | Created | Updated
 */

const SHEET_NAME = "Tasks";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const headers = ["ID","Title","Date","Time","Priority","Category","Repeat","End",
                   "Notes","DoneJSON","SkippedJSON","Created","Updated"];
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = p.action || "getTasks";
    let result;

    if (action === "getTasks") {
      result = {ok:true, tasks:readTasks_()};
    } else if (action === "saveAll") {
      const payload = p.payload ? JSON.parse(decodeURIComponent(p.payload)) : [];
      result = saveAll_(payload);
    } else if (action === "upsertTask") {
      const task = p.payload ? JSON.parse(decodeURIComponent(p.payload)) : null;
      result = upsertTask_(task);
    } else if (action === "deleteTask") {
      result = deleteTask_(String(p.id || ""));
    } else {
      result = {ok:false,error:"Unknown action"};
    }

    return jsonp_(result, p.callback);
  } catch (err) {
    return jsonp_({ok:false,error:String(err)}, e && e.parameter && e.parameter.callback);
  }
}

function jsonp_(obj, callback) {
  const body = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + body + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function readTasks_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2,1,last-1,13).getValues();
  return values.filter(r => String(r[0] || "").trim()).map(r => ({
    id:String(r[0]),
    title:String(r[1]),
    date:sheetDate_(r[2]),
    time:String(r[3] || ""),
    pri:Number(r[4]) || 2,
    cat:String(r[5] || "Other"),
    repeat:String(r[6] || "none"),
    end:sheetDate_(r[7]),
    notes:String(r[8] || ""),
    done:String(r[9] || "{}"),
    skipped:String(r[10] || "{}"),
    created:Number(r[11]) || Date.now(),
    updated:Number(r[12]) || Date.now()
  }));
}

function sheetDate_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v);
}

function saveAll_(incoming) {
  const sh = getSheet_();
  const existing = readTasks_();
  const byId = {};
  existing.forEach(t => byId[t.id] = t);

  incoming.forEach(t => {
    if (!t || !t.id) return;
    byId[t.id] = t;
  });

  const incomingIds = {};
  incoming.forEach(t => { if (t && t.id) incomingIds[t.id] = true; });

  // Taskcal sends the complete current list, so IDs absent from the client
  // are treated as deleted.
  Object.keys(byId).forEach(id => {
    if (!incomingIds[id]) delete byId[id];
  });

  const all = Object.keys(byId).map(id => byId[id]);
  const rows = all.map(t => [
    t.id, t.title, t.date, t.time, t.pri, t.cat, t.repeat, t.end || "",
    t.notes || "", t.done || "{}", t.skipped || "{}",
    t.created || Date.now(), Date.now()
  ]);

  // Replace data below the header. This keeps the sheet easy to inspect.
  const oldLast = sh.getLastRow();
  if (oldLast > 1) sh.getRange(2,1,oldLast-1,13).clearContent();
  if (rows.length) sh.getRange(2,1,rows.length,13).setValues(rows);

  return {ok:true,count:rows.length};
}


function upsertTask_(t) {
  if (!t || !t.id) return {ok:false,error:"Missing task ID"};
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2,1,last-1,1).getValues().flat().map(String);
    const idx = ids.indexOf(String(t.id));
    const row = [[
      t.id,t.title,t.date,t.time,t.pri,t.cat,t.repeat,t.end||"",
      t.notes||"",t.done||"{}",t.skipped||"{}",t.created||Date.now(),Date.now()
    ]];
    if (idx >= 0) {
      sh.getRange(idx+2,1,1,13).setValues(row);
      return {ok:true,action:"updated"};
    }
  }
  sh.appendRow([
    t.id,t.title,t.date,t.time,t.pri,t.cat,t.repeat,t.end||"",
    t.notes||"",t.done||"{}",t.skipped||"{}",t.created||Date.now(),Date.now()
  ]);
  return {ok:true,action:"created"};
}

function deleteTask_(id) {
  if (!id) return {ok:false,error:"Missing task ID"};
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return {ok:true,action:"not-found"};
  const ids = sh.getRange(2,1,last-1,1).getValues().flat().map(String);
  const idx = ids.indexOf(id);
  if (idx < 0) return {ok:true,action:"not-found"};
  sh.deleteRow(idx+2);
  return {ok:true,action:"deleted"};
}


https://script.google.com/macros/s/AKfycbyrPpFxP3LF1Ib_Z7kQtJwDk1EHXwjhUX5LXIzb8awIR8YbNVVrIJf_2xRktTXVkeSQSg/exec



AKfycbyrPpFxP3LF1Ib_Z7kQtJwDk1EHXwjhUX5LXIzb8awIR8YbNVVrIJf_2xRktTXVkeSQSg