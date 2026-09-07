/**
 * Приёмник заявок с лендинга https://sorvalliny.github.io/astashovo/
 * Пишет каждую заявку строкой в таблицу «Гастро-Терем 2026 — заявки».
 *
 * Как поставить:
 *   1. Открыть таблицу → Расширения → Apps Script
 *   2. Вставить этот код вместо содержимого Code.gs, сохранить
 *   3. Развернуть → Новое развёртывание → тип «Веб-приложение»
 *        Запуск от имени: Я
 *        У кого есть доступ: Все
 *   4. Скопировать URL развёртывания (вида https://script.google.com/macros/s/…/exec)
 *      и вписать его в index.html в константу SHEET_ENDPOINT
 */

var SHEET_ID = '15Lw7_CcqgFSMmE0jFC0-64vAytibcZ9khQrMO1o8sVE';

var HEADERS = [
  'Дата заявки', 'Кто бронирует', 'Телеграм', 'Телефон', 'Тариф', 'Человек',
  'Цена с человека', 'Итого', 'Оплата', 'К оплате сейчас', 'Участники (ФИО)',
  'Комментарий', 'Статус', 'Оплачено'
];

function sheet_() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Таблица считает формулой всё, что начинается с = + - @ — телефон «+7 900…»
 * превращался в #ERROR!. Апостроф впереди заставляет хранить значение текстом.
 */
function text_(v) {
  var s = (v === null || v === undefined) ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var d = JSON.parse(e.postData.contents);
    var names = Array.isArray(d.names) ? d.names.filter(String) : [];

    sheet_().appendRow([
      new Date(),
      text_(d.who),
      text_(d.tg),
      text_(d.phone),
      text_(d.tariff),
      Number(d.people) || '',
      Number(d.price) || '',
      Number(d.total) || '',
      d.payMode || '',
      Number(d.due) || '',
      text_(names.join('\n')),
      text_(d.comment),
      'Новая',
      ''
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** Проверка, что развёртывание живо: открыть URL в браузере. */
function doGet() {
  return json_({ ok: true, service: 'Гастро-Терем · приём заявок' });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Разовая уборка: удаляет строки, где «Кто бронирует» начинается с «ТЕСТ».
 * Запускать вручную из редактора: выбрать функцию в списке и нажать «Выполнить».
 */
function cleanupTestRows() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return;
  var col = sh.getRange(2, 2, last - 1, 1).getValues();
  var killed = 0;
  for (var i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]).indexOf('ТЕСТ') === 0) { sh.deleteRow(i + 2); killed++; }
  }
  Logger.log('Удалено строк: ' + killed);
}
