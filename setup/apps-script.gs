/**
 * Приёмник результатов тренажёра по истории.
 * Вставляется в Google-таблицу: Расширения → Apps Script.
 * Каждая тренировка добавляет строку на лист «Сессии», все ответы — на лист «Ответы»,
 * блок ПРОГРЕСС ученика — на лист «Прогресс», присылает письмо и (по желанию)
 * дописывает файл «История — результаты.md» на Google Диске.
 *
 * После изменения кода: Начать развертывание → Управление развертываниями →
 * карандаш → Версия: «Новая версия» → Развернуть. Адрес /exec останется прежним.
 */

const TOKEN = 'jKfBacutjmCuIp4b';   // должен совпадать с token в config.js
const EMAIL = '';                   // пусто = письма приходят владельцу скрипта
const SEND_EMAIL = true;            // письмо после каждой тренировки
const APPEND_MD = true;             // дописывать MD-файл на Google Диске
const MD_NAME = 'История — результаты.md';
const TEST_NAMES = ['тест', 'test', 'тест1'];  // имена для пробных прогонов (см. deleteTestRuns)

const H_SESS = ['Дата', 'Имя', 'Устройство', 'Тренировка №', 'Верно', 'Всего', 'С сомнением', 'Повторить', 'ID'];
const H_ANS = ['Дата', 'Имя', 'ID тренировки', '№ вопроса', 'Тема', 'Вопрос', 'Результат', 'Ответ ученика', 'Правильный ответ', 'Повтор'];
const H_PROG = ['Обновлено', 'Имя', 'Блок ПРОГРЕСС'];

// Запустите один раз из редактора: Google попросит разрешения и пришлёт тестовое письмо.
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, 'Сессии', H_SESS);
  if (SEND_EMAIL) MailApp.sendEmail(EMAIL || Session.getEffectiveUser().getEmail(),
    'История: тестовое письмо', 'Скрипт настроен. Таблица: ' + ss.getUrl());
  if (APPEND_MD) appendMd_(new Date(), 'Тест', 'Скрипт настроен.');
}

// Стирает ВСЕ результаты: листы «Сессии», «Ответы», «Прогресс» и MD-файл.
// Нужен, чтобы убрать пробные прогоны перед началом настоящих тренировок.
function clearAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  resetSheet_(ss, 'Сессии', H_SESS);
  resetSheet_(ss, 'Ответы', H_ANS);
  resetSheet_(ss, 'Прогресс', H_PROG);
  const files = DriveApp.getFilesByName(MD_NAME);
  if (APPEND_MD && files.hasNext()) files.next().setContent('# Результаты тренировок по истории\n');
}

// Удаляет только строки с тестовыми именами из TEST_NAMES (из всех трёх листов).
// Записи в MD-файле не трогает.
function deleteTestRuns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isTest = n => TEST_NAMES.indexOf(String(n).trim().toLowerCase()) !== -1;
  ['Сессии', 'Ответы', 'Прогресс'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const col = col_(sh, 'Имя');
    if (!col) return;
    const vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().flat();
    for (let i = vals.length - 1; i >= 0; i--) if (isTest(vals[i])) sh.deleteRow(i + 2);
  });
}

// doPost вызывает сайт сам. Запускать его из редактора не нужно.
function doPost(e) {
  if (!e || !e.postData) throw new Error('doPost вызывается сайтом. Из редактора запускайте setup, clearAll или deleteTestRuns.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return out_('bad token');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const when = new Date(d.sentAt || Date.now());
    const name = String(d.name || 'без имени').slice(0, 30);
    const sid = d.device + '-' + name + '-' + d.session;

    const sessions = sheet_(ss, 'Сессии', H_SESS);
    const idCol = col_(sessions, 'ID');
    if (idCol && sessions.getLastRow() > 1) {
      const ids = sessions.getRange(2, idCol, sessions.getLastRow() - 1, 1).getValues().flat();
      if (ids.indexOf(sid) !== -1) return out_('duplicate');
    }
    sessions.appendRow([when, safe_(name), d.device, d.session, d.good, d.total, d.unsure,
      safe_((d.weakTopics || []).join(', ')), sid]);

    const answers = sheet_(ss, 'Ответы', H_ANS);
    const rows = (d.answers || []).map(a => [when, safe_(name), sid, a.id, safe_(a.topic), safe_(a.question),
      a.result, safe_(a.chosen), safe_(a.correct), a.repeat ? 'да' : '']);
    if (rows.length) answers.getRange(answers.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    const prog = sheet_(ss, 'Прогресс', H_PROG);
    const names = prog.getLastRow() > 1 ? prog.getRange(2, 2, prog.getLastRow() - 1, 1).getValues().flat() : [];
    const r = names.indexOf(name);
    const row = r === -1 ? prog.getLastRow() + 1 : r + 2;
    prog.getRange(row, 1, 1, 3).setValues([[when, safe_(name), safe_(d.progress)]]);

    if (SEND_EMAIL) {
      MailApp.sendEmail(EMAIL || Session.getEffectiveUser().getEmail(),
        'История — ' + name + ': ' + d.good + ' из ' + d.total,
        d.report + '\n\nВсе результаты: ' + ss.getUrl());
    }
    if (APPEND_MD) appendMd_(when, name, d.report);
    return out_('ok');
  } finally {
    lock.releaseLock();
  }
}

// Открыть адрес веб-приложения в браузере: должно показать «работает».
function doGet() { return out_('работает'); }

function appendMd_(when, name, text) {
  const files = DriveApp.getFilesByName(MD_NAME);
  const f = files.hasNext() ? files.next()
    : DriveApp.createFile(MD_NAME, '# Результаты тренировок по истории\n', MimeType.PLAIN_TEXT);
  const stamp = Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  f.setContent(f.getBlob().getDataAsString() + '\n\n## ' + stamp + ' — ' + name + '\n\n' + text + '\n');
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    header_(sh, header);
  }
  return sh;
}

function resetSheet_(ss, name, header) {
  const sh = ss.getSheetByName(name);
  if (!sh) return sheet_(ss, name, header);
  sh.clear();
  header_(sh, header);
  return sh;
}

function header_(sh, header) {
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function col_(sh, title) {
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  return head.indexOf(title) + 1;
}

// Текст, начинающийся с = + - @, таблица приняла бы за формулу.
function safe_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out_(s) { return ContentService.createTextOutput(s); }
