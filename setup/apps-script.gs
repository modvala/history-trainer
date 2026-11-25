/**
 * Приёмник результатов тренажёра по истории.
 * Вставляется в Google-таблицу: Расширения → Apps Script.
 * Каждая тренировка добавляет строку на лист «Сессии», все ответы — на лист «Ответы»,
 * последний блок ПРОГРЕСС — на лист «Прогресс», присылает письмо и (по желанию)
 * дописывает файл «История — результаты.md» на Google Диске.
 */

const TOKEN = 'jKfBacutjmCuIp4b';   // должен совпадать с token в config.js
const EMAIL = '';                   // пусто = письма приходят владельцу скрипта
const SEND_EMAIL = true;            // письмо после каждой тренировки
const APPEND_MD = true;             // дописывать MD-файл на Google Диске
const MD_NAME = 'История — результаты.md';

// Запустите один раз из редактора: Google попросит разрешения и пришлёт тестовое письмо.
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, 'Сессии', ['Дата', 'Устройство', 'Тренировка №', 'Верно', 'Всего', 'С сомнением', 'Повторить', 'ID']);
  if (SEND_EMAIL) MailApp.sendEmail(EMAIL || Session.getEffectiveUser().getEmail(),
    'История: тестовое письмо', 'Скрипт настроен. Таблица: ' + ss.getUrl());
  if (APPEND_MD) appendMd_(new Date(), 'Тест: скрипт настроен.');
}

// doPost вызывает сайт сам. Запускать его из редактора не нужно.
function doPost(e) {
  if (!e || !e.postData) throw new Error('doPost вызывается сайтом. Из редактора запускайте функцию setup.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return out_('bad token');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const when = new Date(d.sentAt || Date.now());
    const sid = d.device + '-' + d.session;

    const sessions = sheet_(ss, 'Сессии',
      ['Дата', 'Устройство', 'Тренировка №', 'Верно', 'Всего', 'С сомнением', 'Повторить', 'ID']);
    const last = sessions.getLastRow();
    if (last > 1) {
      const ids = sessions.getRange(2, 8, last - 1, 1).getValues().flat();
      if (ids.indexOf(sid) !== -1) return out_('duplicate');
    }
    sessions.appendRow([when, d.device, d.session, d.good, d.total, d.unsure,
      safe_((d.weakTopics || []).join(', ')), sid]);

    const answers = sheet_(ss, 'Ответы',
      ['Дата', 'ID тренировки', '№ вопроса', 'Тема', 'Вопрос', 'Результат', 'Ответ ребёнка', 'Правильный ответ', 'Повтор']);
    const rows = (d.answers || []).map(a => [when, sid, a.id, safe_(a.topic), safe_(a.question),
      a.result, safe_(a.chosen), safe_(a.correct), a.repeat ? 'да' : '']);
    if (rows.length) answers.getRange(answers.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    const prog = sheet_(ss, 'Прогресс', ['Обновлено', 'Блок ПРОГРЕСС']);
    prog.getRange(2, 1, 1, 2).setValues([[when, safe_(d.progress)]]);

    if (SEND_EMAIL) {
      MailApp.sendEmail(EMAIL || Session.getEffectiveUser().getEmail(),
        'История: ' + d.good + ' из ' + d.total,
        d.report + '\n\nВсе результаты: ' + ss.getUrl());
    }
    if (APPEND_MD) appendMd_(when, d.report);
    return out_('ok');
  } finally {
    lock.releaseLock();
  }
}

// Открыть адрес веб-приложения в браузере: должно показать «работает».
function doGet() { return out_('работает'); }

function appendMd_(when, text) {
  const files = DriveApp.getFilesByName(MD_NAME);
  const f = files.hasNext() ? files.next()
    : DriveApp.createFile(MD_NAME, '# Результаты тренировок по истории\n', MimeType.PLAIN_TEXT);
  const stamp = Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  f.setContent(f.getBlob().getDataAsString() + '\n\n## ' + stamp + '\n\n' + text + '\n');
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

// Текст, начинающийся с = + - @, таблица приняла бы за формулу.
function safe_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out_(s) { return ContentService.createTextOutput(s); }
