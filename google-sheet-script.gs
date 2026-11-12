/**
 * Falegnametro → Google Sheet
 * Incolla questo codice in Estensioni → Apps Script del foglio e pubblicalo come "App web".
 * Ogni test compilato aggiunge una riga al foglio "Risposte". Le colonne si creano da sole.
 */
const SHEET_NAME = 'Risposte';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const incoming = JSON.parse(e.postData.contents);
    const data = Object.assign({ data_ora: new Date() }, incoming);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    let headers = sh.getLastColumn() > 0
      ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      : [];
    const missing = Object.keys(data).filter(k => headers.indexOf(k) === -1);
    if (missing.length) {
      headers = headers.concat(missing);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    const row = headers.map(h => {
      const v = data[h];
      if (v === undefined || v === null) return '';
      if (Array.isArray(v)) return v.join(', ');
      if (typeof v === 'object' && !(v instanceof Date)) return JSON.stringify(v);
      return v;
    });
    sh.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Per controllare dal browser che l'app web risponda.
function doGet() {
  return ContentService.createTextOutput('Falegnametro: collegamento attivo');
}
