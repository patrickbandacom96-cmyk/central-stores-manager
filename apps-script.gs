// Google Sheets link-up for Central Stores Manager.
// 1. Make a Google Sheet. Open Extensions > Apps Script. Paste this, replacing everything.
// 2. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Deploy.
// 3. Copy the Web app URL and paste it in the app under More > Google Sheet link.
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  write(ss, 'Stock', ['Item', 'Company', 'Unit', 'In stock', 'Minimum'],
    d.items.map(function (i) { return [i.name, i.company, i.unit, i.qty, i.min]; }));
  write(ss, 'History', ['Date', 'Type', 'Item', 'Qty', 'Person / note', 'Balance after'],
    d.log.map(function (l) { return [l.date, l.type, l.name, l.qty, l.note, l.after]; }));
  return ContentService.createTextOutput('ok');
}
function write(ss, name, head, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
}
