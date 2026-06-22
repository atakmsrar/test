const SPREADSHEET_ID = '18WXyXuFPx00LN6zmH7FVZKIiQ2jzazChJ5AoiA9KU8s';
const SHEET_GID = 1615180526;

const HEADERS = [
  'weekKey',
  'Неделя',
  'Дата',
  'Описание',
  'Сумма',
  'Валюта',
  'Курс ₴ за $',
  'Сумма в $',
  'Кошелёк',
  'Итого день $',
  'Основной день $',
  'Доп. день $',
  'Отправитель',
  'Сохранено в',
  'Статус'
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getSheetByGid_(ss, SHEET_GID);
    ensureHeaders_(sheet);

    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    if (!rows.length) return json_({ ok: true, written: 0 });

    if (payload.mode === 'replaceWeek') {
      deleteMatchingRows_(sheet, payload.weekKey, null);
    }
    if (payload.mode === 'replaceDay') {
      const dayDate = rows[0].date;
      deleteMatchingRows_(sheet, payload.weekKey, dayDate);
    }

    const values = rows.map(r => [
      r.weekKey || payload.weekKey || '',
      r.weekRange || payload.weekRange || '',
      r.date || '',
      r.description || '',
      r.amount || '',
      r.currency || '',
      r.rate || '',
      r.amountUsd || '',
      r.wallet || '',
      r.dayTotalUsd || '',
      r.mainTotalUsd || '',
      r.altTotalUsd || '',
      r.sender || '',
      r.savedAt || new Date().toISOString(),
      r.status || ''
    ]);

    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, HEADERS.length).setValues(values);
    return json_({ ok: true, written: values.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function getSheetByGid_(ss, gid) {
  const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid));
  if (!sheet) throw new Error('Лист с gid=' + gid + ' не найден');
  return sheet;
}

function ensureHeaders_(sheet) {
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  const current = range.getValues()[0];
  const empty = current.every(v => v === '');
  if (empty) {
    range.setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }
  HEADERS.forEach((h, i) => {
    if (!current[i]) sheet.getRange(1, i + 1).setValue(h);
  });
}

function deleteMatchingRows_(sheet, weekKey, date) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const rowWeekKey = values[i][0];
    const rowDate = values[i][2];
    if (rowWeekKey === weekKey && (!date || rowDate === date)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================
// ДИАГНОСТИКА: запусти вручную в редакторе GAS, чтобы увидеть
// и удалить все триггеры — именно они отправляли "." в Telegram.
// =============================================================
function listAndDeleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.length) {
    Logger.log('Триггеры не найдены.');
    return;
  }
  triggers.forEach(t => {
    Logger.log(
      'Удаляю триггер: ' + t.getHandlerFunction() +
      ' | тип: ' + t.getEventType() +
      ' | источник: ' + t.getTriggerSource()
    );
    ScriptApp.deleteTrigger(t);
  });
  Logger.log('Все триггеры удалены.');
}
