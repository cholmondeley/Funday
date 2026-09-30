// Paste this into the Google Sheet: Extensions > Apps Script, replacing everything.
// Then Deploy > New deployment > type "Web app" > Execute as: Me > Who has access: Anyone.
// Copy the Web app URL into SYNC_URL at the top of index.html.

const TAB = 'Tracking';
const HEADERS = ['id', 'timestamp', 'date', 'type', 'player', 'spinId', 'item'];

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1).map(r => ({
    id: String(r[0]), ts: r[1], date: String(r[2]), type: r[3],
    player: r[4], spinId: String(r[5] || ''), item: r[6],
  }));
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(String) : [];
    if (!ids.includes(String(d.id))) {
      // date is stored as text so Sheets doesn't turn it into a Date
      sh.appendRow([d.id, d.ts, "'" + d.date, d.type, d.player, d.spinId || '', d.item || '']);
    }
  } finally { lock.releaseLock(); }
  return ContentService.createTextOutput('ok');
}
