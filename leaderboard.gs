// Crea il progetto Apps Script da Estensioni > Apps Script in un foglio Google.
function preparaClassifica() {
  const file = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = file.getSheetByName('Classifica') || file.insertSheet('Classifica');
  if (!sheet.getLastRow()) sheet.appendRow(['Nickname', 'Punteggio']);
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', file.getId());
}

function classifica() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return SpreadsheetApp.openById(id).getSheetByName('Classifica');
}

function righe(sheet) {
  const count = sheet.getLastRow() - 1;
  return count > 0 ? sheet.getRange(2, 1, count, 2).getValues() : [];
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function chiave(name) {
  return name.normalize('NFKC').toLowerCase();
}

function doGet() {
  const rows = righe(classifica()).map(([name, score]) => ({ name, score }));
  rows.sort((a, b) => b.score - a.score || a.name.localeCompare(b.name, 'it'));
  return json(rows);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (_) {
    return json({ ok: false });
  }
  const name = typeof data?.name === 'string' ? data.name.trim() : '';
  const score = data?.score;
  if (!name || name.length > 14 || /[\u0000-\u001f\u007f]/.test(name) ||
      !Number.isSafeInteger(score) || score < 0) return json({ ok: false });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = classifica();
    const rows = righe(sheet);
    const index = rows.findIndex(([savedName]) => chiave(savedName) === chiave(name));
    if (index === -1) {
      const row = rows.length + 2;
      sheet.getRange(row, 1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(name).build());
      sheet.getRange(row, 2).setValue(score);
    } else if (score > rows[index][1]) {
      sheet.getRange(index + 2, 2).setValue(score);
    }
    SpreadsheetApp.flush();
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
