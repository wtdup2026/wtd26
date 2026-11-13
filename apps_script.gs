/**
 * World Toilet Day 2026 — Registration receiver
 * Paste this whole file into: Google Sheet > Extensions > Apps Script
 * Then: Deploy > New deployment > Web app
 *   Execute as: Me    |    Who has access: Anyone
 */
const SHEET_NAME = 'Registrations';
const FIELDS  = ['name','designation','category','organisation','district','mobile','email','gender','sessions','lunch','access','remarks'];
const HEADERS = ['Timestamp','Registration ID','Name','Designation','Category','Organisation / ULB','District','Mobile','Email','Gender','Sessions of interest','Lunch','Accessibility needs','Remarks'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#e6f2ee');
      sh.setFrozenRows(1);
    }
    const p = e.parameter || {};
    const mobile = String(p.mobile || '').trim();
    if (!/^[6-9]\d{9}$/.test(mobile)) return out({ok:false, error:'invalid mobile'});

    // If this mobile number is already registered, return the same ID (no duplicate row)
    const last = sh.getLastRow();
    if (last > 1) {
      const rows = sh.getRange(2,1,last-1,8).getValues();
      for (const r of rows) {
        if (String(r[7]) === mobile) return out({ok:true, id:r[1], duplicate:true});
      }
    }
    const id = 'WTD26-' + String(last).padStart(4,'0');
    sh.appendRow([new Date(), id].concat(FIELDS.map(f => String(p[f] || '').slice(0,500))));
    sh.getRange(sh.getLastRow(), 8).setNumberFormat('@');
    return out({ok:true, id:id});
  } catch (err) {
    return out({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out({ok:true, msg:'WTD 2026 registration endpoint is running'}); }

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
