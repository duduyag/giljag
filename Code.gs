const SPREADSHEET_ID = '1hg3nTgKufgrwbmwuzrxg779qbu_uKTU-pcewsy5cFuI';
const SHEET_NAME = 'Sheet1';
const HEADERS = ['Timestamp', 'First name', 'Last name', 'Region', 'Phone', 'Source', 'Page URL', 'User agent'];

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'ICD antisnore leads endpoint' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(body);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    }

    sheet.appendRow([
      data.ts || new Date().toISOString(),
      data.firstName || '',
      data.lastName || '',
      data.region || '',
      data.phone || '',
      data.source || 'ICD-antisnore',
      data.pageUrl || '',
      data.userAgent || '',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message ? err.message : err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
