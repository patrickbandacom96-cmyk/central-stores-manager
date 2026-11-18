// Google Sheet link-up for Central Stores Manager (two-way).
// 1. Open your Google Sheet. Extensions > Apps Script. Delete everything and paste this. Save.
// 2. Deploy > New deployment > type: Web app. Execute as: Me. Who has access: Anyone. Deploy.
// 3. Allow access when asked. Copy the Web app URL.
// 4. In the app: Settings: Google Sheet and PIN. Paste the URL and tap Save.
var STOCK = ['item', 'company', 'unit', 'qty', 'min'];
var LOG = ['date', 'item', 'company', 'type', 'qty', 'person', 'note'];
var HEAD_STOCK = ['Item', 'Company', 'Unit', 'In hand', 'Low stock level'];
var HEAD_LOG = ['Date', 'Item', 'Company', 'Type', 'Qty', 'Person', 'Note'];

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return out({ ok: true, stock: read(ss, 'Stock', STOCK), log: read(ss, 'History', LOG) });
}
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    write(ss, 'Stock', HEAD_STOCK, STOCK, d.stock || []);
    write(ss, 'History', HEAD_LOG, LOG, d.log || []);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}
function read(ss, name, keys) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, keys.length).getValues();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var o = {};
    keys.forEach(function (k, i) {
      var x = r[i];
      o[k] = (x instanceof Date) ? x.toISOString() : x;
    });
    return o;
  });
}
function write(ss, name, head, keys, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (!rows.length) return;
  var vals = rows.map(function (r) { return keys.map(function (k) { return r[k] == null ? '' : r[k]; }); });
  var dateCol = keys.indexOf('date');
  if (dateCol > -1) sh.getRange(2, dateCol + 1, vals.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, vals.length, keys.length).setValues(vals);
}
