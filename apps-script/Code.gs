/**
 * צד השרת של דף ההרשמה ליום מהו"ת.
 * יש להדביק את הקובץ הזה בעורך Apps Script שמחובר לגיליון (Extensions > Apps Script),
 * ולהגדיר את סיסמת המנהל ב-Project Settings > Script Properties בשם ADMIN_PASSWORD.
 * הוראות מלאות בקובץ SETUP.md.
 */

// חייב להיות זהה לרשימת השעות ב-index.html
const SLOTS = [
  "16:00", "16:10", "16:20", "16:30", "16:40", "16:50",
  "17:00", "17:10", "17:20", "17:30", "17:40", "17:50",
  "18:00", "18:10", "18:20", "18:30", "18:40", "18:50",
  "19:00", "19:10", "19:20", "19:30", "19:40", "19:50"
];

const SHEET_NAME = "הרשמות";
const HEADERS = ["שעה", "שם ההורה", "שם הילד/ה", "קוד ביטול", "זמן הרשמה"];
const MAX_NAME_LENGTH = 60;

// ---------- נקודות כניסה ----------

// קריאה ציבורית: מחזירה רק אילו שעות תפוסות, בלי שמות
function doGet() {
  const rows = readRows_(getSheet_());
  return json_({ ok: true, taken: rows.map(r => r.time) });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: "בקשה לא תקינה" });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json_({ ok: false, error: "המערכת עמוסה כרגע, נסו שוב בעוד רגע" });
  }

  try {
    switch (req.action) {
      case "book":        return json_(book_(req));
      case "cancel":      return json_(cancel_(req));
      case "adminList":   return json_(withAdmin_(req, adminList_));
      case "adminCancel": return json_(withAdmin_(req, adminCancel_));
      case "clearAll":    return json_(withAdmin_(req, clearAll_));
      default:            return json_({ ok: false, error: "פעולה לא מוכרת" });
    }
  } catch (err) {
    return json_({ ok: false, error: "שגיאה בשרת: " + err.message });
  } finally {
    lock.releaseLock();
  }
}

// ---------- פעולות ----------

function book_(req) {
  const time = String(req.time || "");
  const parent = clean_(req.parent);
  const child = clean_(req.child);

  if (SLOTS.indexOf(time) === -1) return { ok: false, error: "שעה לא תקינה" };
  if (!parent || !child) return { ok: false, error: "יש למלא שם הורה ושם ילד/ה" };

  const sheet = getSheet_();
  const rows = readRows_(sheet);

  if (rows.some(r => r.time === time)) {
    return { ok: false, error: "השעה " + time + " נתפסה זה עתה. בחרו שעה אחרת.", code: "TAKEN" };
  }
  const same = rows.find(r => norm_(r.child) === norm_(child) && norm_(r.parent) === norm_(parent));
  if (same) {
    return { ok: false, error: child + " כבר רשום/ה לשעה " + same.time + ". כדי להחליף שעה, בטלו קודם את ההרשמה הקיימת." };
  }

  const code = String(Math.floor(1000 + Math.random() * 9000));
  sheet.appendRow([time, parent, child, code, new Date()]);
  sortSheet_(sheet);
  return { ok: true, time: time, code: code };
}

function cancel_(req) {
  const time = String(req.time || "");
  const code = String(req.code || "").trim();
  const sheet = getSheet_();
  const row = readRows_(sheet).find(r => r.time === time);
  if (!row || row.code !== code) {
    return { ok: false, error: "השעה וקוד הביטול לא תואמים" };
  }
  sheet.deleteRow(row.row);
  return { ok: true };
}

function adminList_() {
  const rows = readRows_(getSheet_());
  return {
    ok: true,
    bookings: rows.map(r => ({ time: r.time, parent: r.parent, child: r.child }))
  };
}

function adminCancel_(req) {
  const sheet = getSheet_();
  const row = readRows_(sheet).find(r => r.time === String(req.time || ""));
  if (!row) return { ok: false, error: "לא נמצאה הרשמה בשעה הזו" };
  sheet.deleteRow(row.row);
  return adminList_();
}

function clearAll_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  return adminList_();
}

// ---------- עזר ----------

function withAdmin_(req, fn) {
  const password = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  if (!password) return { ok: false, error: "לא הוגדרה סיסמת מנהל בשרת (ADMIN_PASSWORD)" };
  if (String(req.password || "") !== password) {
    Utilities.sleep(1000); // מאט ניחוש סיסמאות
    return { ok: false, error: "סיסמה שגויה", code: "BAD_PASSWORD" };
  }
  return fn(req);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setFontWeight("bold");
    sheet.setRightToLeft(true);
  }
  // שמירה כטקסט, כדי ש"16:00" לא יהפוך לשעה ו"0123" לא יהפוך למספר
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("D:D").setNumberFormat("@");
  return sheet;
}

function readRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 4).getDisplayValues()
    .map((v, i) => ({ row: i + 2, time: v[0].trim(), parent: v[1], child: v[2], code: v[3].trim() }))
    .filter(r => r.time);
}

function sortSheet_(sheet) {
  const last = sheet.getLastRow();
  if (last > 2) sheet.getRange(2, 1, last - 1, HEADERS.length).sort(1);
}

function clean_(s) {
  return String(s || "").replace(/\s+/g, " ").trim().slice(0, MAX_NAME_LENGTH);
}

function norm_(s) {
  return clean_(s).toLowerCase();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
