/**
 * Приймає заявки та відповіді міні-тесту з лендінгу і записує їх у Google Таблицю.
 *
 * Як підключити: див. README.md у цій папці.
 * Скрипт потрібно створити з самої таблиці: Розширення → Apps Script.
 */

// Куди надсилати сповіщення про нові реєстрації. Залиште порожнім, якщо пошта не потрібна.
const NOTIFY_EMAIL = "";

const SHEET_FORMS = "Реєстрації";
const SHEET_QUIZ = "Міні-тест";

const FORM_HEADERS = ["Час", "Джерело", "Ім'я", "Телефон", "Email", "Ким працює", "Сфера бізнесу", "Запитання до спікерів", "Бал міні-тесту", "Сесія", "Сторінка"];
const QUIZ_HEADERS = ["Час", "Сесія", "Бал", "Результат", "1. Скільки дзвінків без відповіді", "2. Дзвінки потрапляють у CRM", "3. Передзвонюють того ж дня", "4. Є запис або підсумок розмови", "5. Є наступний крок із датою", "Сторінка"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();

    if (d.source === "quiz") {
      const answers = (d.answers || []).map(function (x) { return clean(x.a); });
      while (answers.length < 5) answers.push("");
      getSheet(ss, SHEET_QUIZ, QUIZ_HEADERS).appendRow(
        [now, clean(d.sid), Number(d.score) || 0, clean(d.result)].concat(answers.slice(0, 5), [clean(d.page)])
      );
    } else {
      getSheet(ss, SHEET_FORMS, FORM_HEADERS).appendRow([
        now, clean(d.source), clean(d.name), asText(d.phone), clean(d.email), clean(d.role),
        clean(d.industry), clean(d.question), clean(d.quiz), clean(d.sid), clean(d.page)
      ]);
      if (NOTIFY_EMAIL) {
        MailApp.sendEmail(
          NOTIFY_EMAIL,
          "Нова реєстрація на вебінар: " + clean(d.name),
          ["Ім'я: " + clean(d.name), "Телефон: " + clean(d.phone), "Email: " + clean(d.email),
           "Ким працює: " + clean(d.role), "Сфера: " + clean(d.industry),
           "Запитання: " + clean(d.question), "Бал міні-тесту: " + clean(d.quiz)].join("\n")
        );
      }
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Перевірка в браузері: відкрийте адресу веб-застосунку, має з'явитися {"ok":true,"service":"webinar-leads"}
function doGet() {
  return json({ ok: true, service: "webinar-leads" });
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
  }
  return sh;
}

// Захист від формул: значення, що починаються на = + - @, записуються як текст.
function clean(v) {
  const s = v == null ? "" : String(v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Телефон завжди пишемо як текст, інакше таблиця перетворить 0671234567 на число й загубить початковий нуль.
function asText(v) {
  const s = v == null ? "" : String(v).slice(0, 2000);
  return s ? "'" + s : "";
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
