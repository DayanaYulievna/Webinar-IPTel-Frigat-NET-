/**
 * Приймає заявки та відповіді міні-тесту з лендінгу, записує їх у Google Таблицю
 * і створює лид у Bitrix24 (якщо підключено).
 *
 * Як підключити: див. README.md у цій папці.
 * Скрипт потрібно створити з самої таблиці: Розширення → Apps Script.
 *
 * Налаштування Bitrix24 зберігаються у «Властивостях скрипта» (Налаштування проєкту → Властивості скрипта):
 *   BITRIX_WEBHOOK       адреса вхідного вебхука, наприклад https://портал.bitrix24.ua/rest/1/код/   (обов'язково)
 *   BITRIX_ASSIGNED_ID   id відповідального (Олексій Гордєєв). Без нього лид створиться без відповідального
 *   BITRIX_SOURCE_ID     код джерела лида, за замовчуванням WEB
 */

// Куди надсилати сповіщення про нові реєстрації. Залиште порожнім, якщо пошта не потрібна.
const NOTIFY_EMAIL = "";

const SHEET_FORMS = "Реєстрації";
const SHEET_QUIZ = "Міні-тест";

const FORM_HEADERS = ["Час", "Джерело", "Ім'я", "Телефон", "Email", "Ким працює", "Сфера бізнесу", "Запитання до спікерів",
  "Бал міні-тесту", "Сесія", "Сторінка", "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content", "gclid", "fbclid",
  "Звідки прийшов", "Bitrix24"];
const BITRIX_COL = FORM_HEADERS.indexOf("Bitrix24") + 1; // номер колонки, 1-based

const QUIZ_HEADERS = ["Час", "Сесія", "Бал", "Результат", "1. Скільки дзвінків без відповіді", "2. Дзвінки потрапляють у CRM",
  "3. Передзвонюють того ж дня", "4. Є запис або підсумок розмови", "5. Є наступний крок із датою", "Сторінка",
  "utm_source", "utm_medium", "utm_campaign"];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();

    if (d.source === "quiz") {
      const answers = (d.answers || []).map(function (x) { return clean(x.a); });
      while (answers.length < 5) answers.push("");
      withLock(function () {
        getSheet(ss, SHEET_QUIZ, QUIZ_HEADERS).appendRow(
          [now, clean(d.sid), Number(d.score) || 0, clean(d.result)].concat(
            answers.slice(0, 5), [clean(d.page), clean(d.utm_source), clean(d.utm_medium), clean(d.utm_campaign)])
        );
      });
    } else {
      let sh, rowNum;
      withLock(function () {
        sh = getSheet(ss, SHEET_FORMS, FORM_HEADERS);
        sh.appendRow([
          now, clean(d.source), clean(d.name), asText(d.phone), clean(d.email), clean(d.role), clean(d.industry),
          clean(d.question), clean(d.quiz), clean(d.sid), clean(d.page),
          clean(d.utm_source), clean(d.utm_medium), clean(d.utm_campaign), clean(d.utm_term), clean(d.utm_content),
          clean(d.gclid), clean(d.fbclid), clean(d.ref), ""
        ]);
        rowNum = sh.getLastRow();
      });

      // Bitrix24: помилка тут не повинна втратити заявку, вона вже в таблиці
      const status = sendToBitrix(d);
      sh.getRange(rowNum, BITRIX_COL).setValue(status);

      if (NOTIFY_EMAIL) {
        try {
          MailApp.sendEmail(
            NOTIFY_EMAIL,
            "Нова реєстрація на вебінар: " + clean(d.name),
            ["Ім'я: " + clean(d.name), "Телефон: " + clean(d.phone), "Email: " + clean(d.email),
             "Ким працює: " + clean(d.role), "Сфера: " + clean(d.industry),
             "Запитання: " + clean(d.question), "Бал міні-тесту: " + clean(d.quiz),
             "UTM: " + [d.utm_source, d.utm_medium, d.utm_campaign].filter(Boolean).join(" / "),
             "Bitrix24: " + status].join("\n")
          );
        } catch (mailErr) { /* ліміт пошти не повинен ламати прийом заявки */ }
      }
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Перевірка в браузері: відкрийте адресу веб-застосунку, має з'явитися {"ok":true,"service":"webinar-leads"}
function doGet() {
  return json({ ok: true, service: "webinar-leads" });
}

/* ===================== Bitrix24 ===================== */

// Повертає короткий статус для колонки «Bitrix24». Ніколи не кидає помилку.
function sendToBitrix(d) {
  const props = PropertiesService.getScriptProperties();
  let hook = props.getProperty("BITRIX_WEBHOOK");
  if (!hook) return "Bitrix не налаштовано";
  hook = hook.replace(/\/?$/, "/");
  try {
    const phone = normalizePhone(d.phone);
    const email = str(d.email);

    // 1) дубль: не створюємо другий лид, а додаємо коментар у знайдений
    const dup = findDuplicateLead(hook, phone, email);
    if (dup) {
      bx(hook, "crm.timeline.comment.add", {
        fields: { ENTITY_ID: dup, ENTITY_TYPE: "lead", COMMENT: "Повторна реєстрація на вебінар.\n" + commentText(d) }
      });
      return "Дубль: додано коментар у лід #" + dup;
    }

    // 2) новий лид у воронці лідів
    const fields = {
      TITLE: "Вебінар: " + (str(d.name) || email || phone),
      NAME: str(d.name),
      STATUS_ID: "NEW",
      OPENED: "Y",
      SOURCE_ID: props.getProperty("BITRIX_SOURCE_ID") || "WEB",
      SOURCE_DESCRIPTION: "Лендінг вебінару IPTel × Frigat NET (" + (d.source === "modal" ? "швидка картка" : "форма внизу сторінки") + ")",
      COMMENTS: commentHtml(d),
      UTM_SOURCE: str(d.utm_source),
      UTM_MEDIUM: str(d.utm_medium),
      UTM_CAMPAIGN: str(d.utm_campaign),
      UTM_CONTENT: str(d.utm_content),
      UTM_TERM: str(d.utm_term)
    };
    if (phone) fields.PHONE = [{ VALUE: phone, VALUE_TYPE: "WORK" }];
    if (email) fields.EMAIL = [{ VALUE: email, VALUE_TYPE: "WORK" }];
    const assigned = props.getProperty("BITRIX_ASSIGNED_ID");
    if (assigned) fields.ASSIGNED_BY_ID = Number(assigned);

    const id = bx(hook, "crm.lead.add", { fields: fields, params: { REGISTER_SONET_EVENT: "Y" } });
    return "OK: лід #" + id + (assigned ? "" : " (відповідального не задано)");
  } catch (err) {
    return "ПОМИЛКА: " + safe(err, hook);
  }
}

function findDuplicateLead(hook, phone, email) {
  const checks = [];
  if (phone) checks.push({ type: "PHONE", values: [phone] });
  if (email) checks.push({ type: "EMAIL", values: [email] });
  for (let i = 0; i < checks.length; i++) {
    const res = bx(hook, "crm.duplicate.findbycomm", { type: checks[i].type, values: checks[i].values, entity_type: "LEAD" });
    const ids = res && res.LEAD;
    if (ids && ids.length) return ids[0];
  }
  return null;
}

function bx(hook, method, body) {
  const resp = UrlFetchApp.fetch(hook + method + ".json", {
    method: "post", contentType: "application/json", payload: JSON.stringify(body), muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  let j = {};
  try { j = JSON.parse(resp.getContentText()); } catch (e) { /* не JSON */ }
  if (code >= 400 || j.error) {
    throw new Error((j.error || "HTTP " + code) + (j.error_description ? ": " + j.error_description : ""));
  }
  return j.result;
}

// Повторна спроба для рядків зі статусом «ПОМИЛКА…». Запустіть вручну (Виконати → retryFailed) або повісьте на тригер.
function retryFailed() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FORMS);
  if (!sh || sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, FORM_HEADERS.length).getValues();
  rows.forEach(function (r, i) {
    if (String(r[BITRIX_COL - 1] || "").indexOf("ПОМИЛКА") !== 0) return;
    const d = {
      source: r[1], name: r[2], phone: String(r[3]).replace(/^'/, ""), email: r[4], role: r[5], industry: r[6], question: r[7], quiz: r[8],
      utm_source: r[11], utm_medium: r[12], utm_campaign: r[13], utm_term: r[14], utm_content: r[15], ref: r[18]
    };
    sh.getRange(i + 2, BITRIX_COL).setValue(sendToBitrix(d));
  });
}

// Діагностика: Виконати → checkBitrix, потім «Журнал виконання». Показує, чи працює вебхук і хто відповідальний.
function checkBitrix() {
  const props = PropertiesService.getScriptProperties();
  const hook = (props.getProperty("BITRIX_WEBHOOK") || "").replace(/\/?$/, "/");
  if (hook === "/") { Logger.log("BITRIX_WEBHOOK не задано"); return; }
  try {
    const me = bx(hook, "profile", {});
    Logger.log("Вебхук працює. Користувач вебхука: " + me.NAME + " " + me.LAST_NAME);
  } catch (err) { Logger.log("Вебхук не працює: " + safe(err, hook)); return; }
  const id = props.getProperty("BITRIX_ASSIGNED_ID");
  if (!id) { Logger.log("BITRIX_ASSIGNED_ID не задано: лиди створюватимуться без відповідального"); return; }
  try {
    const u = bx(hook, "user.get", { ID: id });
    Logger.log("Відповідальний: " + (u && u[0] ? u[0].NAME + " " + u[0].LAST_NAME : "користувача з id " + id + " не знайдено"));
  } catch (err) { Logger.log("Не вдалося перевірити відповідального: " + safe(err, hook)); }
}

/* ===================== допоміжне ===================== */

// +380XXXXXXXXX: приймає 0671234567, 380671234567, +380 (67) 123-45-67 та інші записи
function normalizePhone(v) {
  const raw = str(v);
  if (!raw) return "";
  let digits = raw.replace(/\D/g, "");
  if (digits.length === 10 && digits.charAt(0) === "0") digits = "38" + digits;
  else if (digits.length === 9) digits = "380" + digits;
  return digits ? "+" + digits : "";
}

function commentText(d) {
  return commentLines(d).join("\n");
}
function commentHtml(d) {
  return commentLines(d).map(esc).join("<br>");
}
function commentLines(d) {
  const l = [];
  if (str(d.role)) l.push("Ким працює: " + str(d.role));
  if (str(d.industry)) l.push("Сфера бізнесу: " + str(d.industry));
  if (str(d.question)) l.push("Запитання до спікерів: " + str(d.question));
  if (str(d.quiz)) l.push("Бал міні-тесту: " + str(d.quiz));
  if (str(d.ref)) l.push("Звідки прийшов: " + str(d.ref));
  if (str(d.gclid)) l.push("gclid: " + str(d.gclid));
  if (str(d.fbclid)) l.push("fbclid: " + str(d.fbclid));
  l.push("Форма: " + (d.source === "modal" ? "швидка картка" : "внизу сторінки"));
  return l;
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { fn(); } finally { lock.releaseLock(); }
}

function getSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  // Таблиця, створена з готовим рядком заголовків: перейменовуємо першу вкладку замість створення нової.
  if (!sh && name === SHEET_FORMS) {
    const first = ss.getSheets()[0];
    if (first && first.getLastRow() >= 1 && first.getRange(1, 1).getValue() === "Час") {
      first.setName(name);
      sh = first;
    }
  }
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  } else if (sh.getLastColumn() < headers.length) {
    // старіша таблиця з меншою кількістю колонок: дописуємо нові заголовки
    const have = sh.getLastColumn();
    sh.getRange(1, have + 1, 1, headers.length - have).setValues([headers.slice(have)]).setFontWeight("bold");
  }
  return sh;
}

function str(v) {
  return v == null ? "" : String(v).trim().slice(0, 2000);
}

// Захист від формул: значення, що починаються на = + - @, записуються як текст.
function clean(v) {
  const s = str(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Телефон завжди пишемо як текст, інакше таблиця перетворить 0671234567 на число й загубить початковий нуль.
function asText(v) {
  const s = str(v);
  return s ? "'" + s : "";
}

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Прибирає з тексту помилки адресу вебхука (у ній секретний код)
function safe(err, hook) {
  let m = String(err && err.message ? err.message : err);
  if (hook) m = m.split(hook).join("***/");
  return m.replace(/\/rest\/\d+\/[^\/\s"']+\/?/g, "/rest/***/").slice(0, 300);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
