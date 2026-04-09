// Google Apps Script — MGG Aggregate Sports

function aggregateSports() {
  const CONFIG = {
    SOURCE_SPREADSHEET_ID: '', // ID таблиці-джерела
    DEST_SPREADSHEET_ID: '', // якщо порожній — використовує активну таблицю
    DEST_SHEET_NAME: 'MGG Aggregated'
  };

  const sourceSS = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
  const destSS = CONFIG.DEST_SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.DEST_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const limitDate = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000); // +7 днів

  let allRows = [];

  // Отримуємо всі аркуші з таблиці-джерела
  const allSheets = sourceSS.getSheets();

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheet.getLastRow() < 2) return;

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => h.toString().trim().toLowerCase());

    const findHeader = (...possibleNames) => {
      for (const name of possibleNames) {
        const index = headers.indexOf(name.toLowerCase());
        if (index !== -1) return index;
      }
      return -1;
    };

    const idx = {
      name:  findHeader('название матча', 'назва матчу', 'название события'),
      start: findHeader('начало матча', 'начало, utc+3 (киев)', 'начало, utc+2 (киев)', 'time utc', 'start time'),
      date:  findHeader('date', 'дата'),
      cdn:   findHeader('cdn id'),
      mgg:   findHeader('mgg id')
    };

    if (idx.name === -1 || idx.start === -1 || idx.date === -1 || idx.cdn === -1) {
      Logger.log(`⚠️ Пропуск: на аркуші "${sheetName}" немає всіх необхідних заголовків.`);
      return;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateStr = row[idx.date];
      if (!dateStr) continue;

      const parts = dateStr.split('.');
      if (parts.length !== 3) continue;

      const matchDate = new Date(parts[2], parts[1] - 1, parts[0]);
      if (isNaN(matchDate.getTime())) continue;

      matchDate.setHours(0, 0, 0, 0);
      if (matchDate < today || matchDate > limitDate) continue;

      const timeObj = parseMatchTime(row[idx.start], matchDate);
      const mggId = idx.mgg !== -1 ? row[idx.mgg] : '';

      allRows.push([matchDate, row[idx.cdn], row[idx.name], timeObj, mggId]);
    }
  });

  // Сортування за датою, потім за часом
  allRows.sort((a, b) => {
    const dateDiff = a[0].getTime() - b[0].getTime();
    if (dateDiff !== 0) return dateDiff;
    return a[3].getTime() - b[3].getTime();
  });

  // Перетворюємо час у текст перед записом
  const outputRows = allRows.map(r => {
    const date = r[0];
    const cdn = r[1];
    const name = r[2];
    const timeStr = formatTime(r[3]); // "HH:mm"
    const mggId = r[4];
    return [date, cdn, name, timeStr, mggId];
  });

  // Виводимо у цільовий аркуш
  const destSheet = destSS.getSheetByName(CONFIG.DEST_SHEET_NAME) || destSS.insertSheet(CONFIG.DEST_SHEET_NAME);
  destSheet.clear();
  destSheet.getRange('1:1').setFontWeight('bold');
  destSheet.appendRow(['Date', 'CDN ID', 'Название матча', 'Начало матча', 'MGG ID']);

  if (outputRows.length > 0) {
    const range = destSheet.getRange(2, 1, outputRows.length, outputRows[0].length);
    range.setValues(outputRows);
    destSheet.getRange(2, 1, outputRows.length, 1).setNumberFormat('dd.MM.yyyy'); // формат тільки для дати
    destSheet.autoResizeColumns(1, 5);
  }

  Logger.log(`✅ Імпорт завершено. Оброблено ${allRows.length} матчів із ${allSheets.length} листів.`);
}

// Розбір часу з клітинки у Date-об’єкт
function parseMatchTime(timeStr, dateObj) {
  const d = new Date(dateObj);
  if (!timeStr) {
    d.setHours(0, 0, 0, 0);
    return d;
  }
  const match = timeStr.toString().trim().match(/^(\d{1,2}):(\d{2})/);
  if (!match) {
    d.setHours(0, 0, 0, 0);
    return d;
  }
  const hours = parseInt(match[1], 10);
  const minutes = parseInt(match[2], 10);
  d.setHours(hours, minutes, 0, 0);
  return d;
}

// Перетворення часу з Date у текст "HH:mm"
function formatTime(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj)) return '';
  const hours = String(dateObj.getHours()).padStart(2, '0');
  const minutes = String(dateObj.getMinutes()).padStart(2, '0');
  return `${hours}:${minutes}`;
}
