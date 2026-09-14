/*  АНКЕТА БФ «МИР І ДОБРОБУТ 777» → GOOGLE ТАБЛИЦЯ
    ─────────────────────────────────────────────────────────────────────
    Сторінка на GitHub Pages шле сюди кожну відповідь окремо, щойно людина
    перестала друкувати. Тут відповідь записується в таблицю: один рядок
    на питання, повторне збереження оновлює той самий рядок.

    РОЗГОРТАННЯ (один раз):
    1. Таблиця → Розширення → Apps Script → вставити цей файл.
    2. Project Settings → Script properties:
         ACCESS_KEY  — довгий випадковий рядок; той самий стоїть у посиланні ?k=…
         BOT_TOKEN   — (необов'язково) токен @Gregory_Ivd_bot, лише тут, не в коді
         CHAT_IDS    — (необов'язково) chat_id через кому, кому писати про «Готово»
    3. Deploy → New deployment → Web app:
         Execute as: Me;  Who has access: Anyone
    4. URL /exec вписати в CONFIG.endpoint у index.html.
    Після кожної правки коду — Manage deployments → Edit → New version.
*/

const SHEET_ID = "1cd8-b2Oz0LPjwv7W8dMN4JIKgNm3HGEI2l9S1WPfQis";
const SHEET_NAME = "Відповіді";
const HEADER = ["qid", "Розділ", "Питання", "Відповідь", "Оновлено"];

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (!keyOk(p.k)) return json({ ok: false, error: "key" });
    const qid = String(p.qid || "").slice(0, 40);
    if (!/^[\w.-]+$/.test(qid)) return json({ ok: false, error: "qid" });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = sheet();
      const row = findRow(sh, qid);
      const values = [[qid, String(p.section || ""), String(p.question || ""),
                       String(p.answer || "").slice(0, 20000), new Date()]];
      if (row) sh.getRange(row, 1, 1, 5).setValues(values);
      else sh.appendRow(values[0]);
    } finally {
      lock.releaseLock();
    }

    if (qid === "_done") notify("Катерина натиснула «Готово» в анкеті БФ 777. Відповіді — у таблиці.");
    return json({ ok: true, savedAt: new Date().toISOString() });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "load") {
    if (!keyOk(p.k)) return json({ ok: false, error: "key" });
    const sh = sheet();
    const data = sh.getDataRange().getValues().slice(1);
    const answers = {};
    data.forEach(r => { if (r[0]) answers[r[0]] = { answer: r[3], updated: r[4] }; });
    return json({ ok: true, answers: answers });
  }
  return ContentService.createTextOutput("BF777 anketa alive");
}

function keyOk(k) {
  const expected = PropertiesService.getScriptProperties().getProperty("ACCESS_KEY");
  return Boolean(expected) && k === expected;
}

function sheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
  }
  if (sh.getLastRow() === 0 || sh.getRange(1, 1).getValue() !== "qid") {
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRow(sh, qid) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === qid) return i + 2;
  return 0;
}

function notify(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("BOT_TOKEN");
  const chats = (props.getProperty("CHAT_IDS") || "").split(",").map(s => s.trim()).filter(Boolean);
  if (!token || !chats.length) return;
  chats.forEach(chat => {
    try {
      UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
        method: "post", payload: { chat_id: chat, text: text }, muteHttpExceptions: true
      });
    } catch (err) {
      console.warn("Telegram: " + err.message);
    }
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* Разовий запуск у редакторі: створює ACCESS_KEY, якщо його ще немає,
   і пише в журнал готове посилання для Катерини. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty("ACCESS_KEY");
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "").slice(0, 8);
    props.setProperty("ACCESS_KEY", key);
  }
  sheet();
  console.log("Посилання: https://gregory-ivd.github.io/bf777-pytannya/?k=" + key);
}
