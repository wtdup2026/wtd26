/**
 * World Toilet Day 2026 — Registration receiver (v2)
 * - Any mobile number can register any number of times
 * - Every submission gets a new unique ID: WTD26-0001, WTD26-0002 ...
 *
 * Replace ALL code in your Apps Script project with this file, Save, then:
 * Deploy > Manage deployments > (pencil) Edit > Version: New version > Deploy
 * (the Web app URL stays the same)
 */
const SHEET_TITLE = 'WTD 2026 Registrations';
const FIELDS  = ['name','designation','category','organisation','district','mobile','email','gender','sessions','lunch','access','remarks'];
const HEADERS = ['Timestamp','Registration ID','Name','Designation','Category','Organisation / ULB','District','Mobile','Email','Gender','Sessions of interest','Lunch','Accessibility needs','Remarks'];

// Run once from the editor if the sheet does not exist yet (safe to run again)
function setup() { const sh = getSheet_(); Logger.log(sh.getParent().getUrl()); }

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID'), ss;
  try { ss = id ? SpreadsheetApp.openById(id) : null; } catch (e) { ss = null; }
  if (!ss) { ss = SpreadsheetApp.create(SHEET_TITLE); props.setProperty('SHEET_ID', ss.getId()); }
  const sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#e6f2ee');
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return out_({ok:false, error:'busy, please try again'});
  try {
    const p = (e && e.parameter) || {};
    const sh = getSheet_();
    // Unique serial: one more than the highest ID already in the sheet
    let n = 0;
    const last = sh.getLastRow();
    if (last > 1) {
      sh.getRange(2, 2, last - 1, 1).getValues().forEach(r => {
        const m = String(r[0]).match(/(\d+)$/); if (m) n = Math.max(n, +m[1]);
      });
    }
    const id = 'WTD26-' + String(n + 1).padStart(4, '0');
    const row = [new Date(), id].concat(FIELDS.map(f => String(p[f] || '').slice(0, 500)));
    sh.appendRow(row);
    sh.getRange(sh.getLastRow(), 8).setNumberFormat('@').setValue(String(p.mobile || ''));
    SpreadsheetApp.flush();
    return out_({ok:true, id:id});
  } catch (err) {
    return out_({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out_({ok:true, msg:'WTD 2026 registration endpoint is running'}); }

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
