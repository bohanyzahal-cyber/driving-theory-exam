// ווטשדוג בוקר-בחינות — סקריפט Apps Script נפרד (פעולה 7 מהבדיקה של 18/09/2026).
//
// למה: בשלושה ימי בחינות (15–17/09) הוויכוח היה תמיד "גוגל או אנחנו?" — ונענה
// רק בדיעבד, מצילומי מסך של דף הביצועים. הסקריפט הזה מודד כל דקה, מסקריפט
// אחר ומגיליון אחר, שלושה דברים בנפרד ורושם אותם עם חותמת זמן:
//   (א) health           — פעולה ריקה בשרת הבחינות: אפס עבודה, רק ההזנקה והקבלה של גוגל
//   (ב) health&deep=1    — אותה פעולה + קריאת תא אחד מהמסמך של הבחינות (sheetMs)
//   (ג) control          — ping לסקריפט אחר באותו חשבון (bohan-site-server), גיליון אחר
//   (ד) ownSheet         — קריאת תא אחד מהגיליון של הווטשדוג עצמו (Sheets כשירות)
// הפירוש (עמודת verdict):
//   הכול איטי                      → google/account   (גוגל או החשבון)
//   רק sheetMs של (ב) איטי           → our-document     (המסמך שלנו נתקע — ההפרדה היא הפתרון)
//   (א)+(ג) איטיים, sheetMs מהיר     → dispatch         (הזנקה/קבלה של גוגל, לא המסמך)
//   הכול מהיר בזמן שהבוחנים תקועים  → client/network   (הבעיה אצל הלקוח או ברשת)
//
// התקנה (פעם אחת, ~3 דקות):
//   1. גיליון Google חדש בשם "ווטשדוג בחינות" → תוספים → Apps Script (סקריפט צמוד לגיליון).
//   2. להדביק את הקובץ הזה, לשמור, להריץ installWatchdog פעם אחת מהעורך (יבקש הרשאות:
//      גיליונות + UrlFetch).
//   3. לוודא שבגיליון "log" מופיעה שורה, ושמחר בבוקר מופיעות שורות כל 2 דקות בין 07:00 ל-13:00.
//
// תקציב הטריגרים (ביקורת C, R3): הגרסה הקודמת יצרה טריגר של כל דקה לכל היממה —
// 1,440 הרצות ביום, מהן ~360 בתוך החלון עם 3 קריאות UrlFetch כל אחת ≈ 36–48 דקות
// מתוך 90 הדקות שיש לחשבון ליום, ועוד 1,080 הרצות-סרק. עכשיו: טריגר יומי ב-07:00
// יוצר טריגר של כל 2 דקות, וטריגר יומי ב-13:00 מוחק אותו. ≈180 מדידות × ~6 שנ' ≈ 18 דקות,
// ואפס הרצות מחוץ לחלון. בדיקת הבוקר (07:00) מסמנת שורה בולטת אם כבר אז משהו לא תקין.
// בבוקר בחינות: זה המקום הראשון להסתכל בו, לפני דף הביצועים ולפני גיליון אבחון.

var WATCHDOG_EXAM_EXEC = 'https://script.google.com/macros/s/AKfycbzOI0zrDEngP-GvlRblhOk8tQsYBvWZ2gGliIQHTpS67WrDZl4la8NPpwtJr_Vjsh3Gzg/exec';
var WATCHDOG_CONTROL_EXEC = 'https://script.google.com/macros/s/AKfycbwWtoSd7ivgXZi0luvYmX8FIZdSGevAHfbCmvKJElgU5egF2rPlC7m9f-k6OILbqIFT/exec?action=ping';
var WATCHDOG_WINDOW_START_HOUR = 7;    // שעון ישראל, כולל
var WATCHDOG_WINDOW_END_HOUR = 13;     // לא כולל
var WATCHDOG_SLOW_MS = 5000;           // מעל זה = "איטי" לצורך ה-verdict
var WATCHDOG_SHEET = 'log';
var WATCHDOG_TICK_MINUTES = 2;
// עמודות חדשות נוספות תמיד בסוף, כדי שהשורות הישנות בגיליון יישארו קריאות.
var WATCHDOG_HEADERS = ['זמן (ישראל)', 'health ms', 'health', 'deep ms', 'deep sheetMs', 'deep', 'control ms', 'control', 'ownSheet ms', 'verdict', 'build', 'deep totalMs', 'outside ms'];
var WATCHDOG_HANDLERS = ['watchdogTick', 'watchdogOpenWindow', 'watchdogCloseWindow', 'watchdogMorningCheck'];

// הרצה אחת מהעורך. משאירה בדיוק שלושה טריגרים יומיים ואף טריגר דקות.
function installWatchdog() {
  var triggers = ScriptApp.getProjectTriggers(), removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (WATCHDOG_HANDLERS.indexOf(triggers[i].getHandlerFunction()) === -1) continue;
    ScriptApp.deleteTrigger(triggers[i]);
    removed++;
  }
  ScriptApp.newTrigger('watchdogMorningCheck').timeBased().atHour(WATCHDOG_WINDOW_START_HOUR).everyDays(1).inTimezone('Asia/Jerusalem').create();
  ScriptApp.newTrigger('watchdogOpenWindow').timeBased().atHour(WATCHDOG_WINDOW_START_HOUR).everyDays(1).inTimezone('Asia/Jerusalem').create();
  ScriptApp.newTrigger('watchdogCloseWindow').timeBased().atHour(WATCHDOG_WINDOW_END_HOUR).everyDays(1).inTimezone('Asia/Jerusalem').create();
  var row = watchdogMeasure();
  var msg = 'installWatchdog: removed ' + removed + ' old trigger(s); daily ' + WATCHDOG_WINDOW_START_HOUR +
    ':00 morning check + window open, daily ' + WATCHDOG_WINDOW_END_HOUR + ':00 window close; first row: ' + row.join(' | ');
  Logger.log(msg);
  return row;
}

// 07:00 — פותח את חלון המדידה: יוצר טריגר של כל 2 דקות (ומוחק כפילויות).
function watchdogOpenWindow() {
  watchdogDeleteTickTriggers();
  ScriptApp.newTrigger('watchdogTick').timeBased().everyMinutes(WATCHDOG_TICK_MINUTES).create();
  Logger.log('watchdogOpenWindow: tick every ' + WATCHDOG_TICK_MINUTES + ' minutes');
  return 'opened';
}

// 13:00 — סוגר את החלון: מוחק את טריגר הדקות. אין הרצות עד מחר ב-07:00.
function watchdogCloseWindow() {
  var n = watchdogDeleteTickTriggers();
  Logger.log('watchdogCloseWindow: removed ' + n + ' tick trigger(s)');
  return 'closed:' + n;
}

function watchdogDeleteTickTriggers() {
  var triggers = ScriptApp.getProjectTriggers(), removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() !== 'watchdogTick') continue;
    ScriptApp.deleteTrigger(triggers[i]);
    removed++;
  }
  return removed;
}

// 07:00 — מדידה אחת לפני שהבוחנים מתחברים. אם כבר עכשיו ה-verdict אינו 'ok',
// השורה מסומנת (רקע אדום) כדי שתיראה מיד בפתיחת הגיליון.
function watchdogMorningCheck() {
  var row = watchdogMeasure();
  var verdict = row[9];
  if (verdict !== 'ok') {
    try {
      var sheet = getWatchdogSheet();
      sheet.getRange(sheet.getLastRow(), 1, 1, WATCHDOG_HEADERS.length).setBackground('#fde2e1');
    } catch (e) { Logger.log('watchdogMorningCheck: highlight failed: ' + e); }
  }
  Logger.log('watchdogMorningCheck: ' + row.join(' | '));
  return row;
}

// הטריגר. מחוץ לחלון הבוקר חוזר מיד; בתוכו — מדידה אחת ושורה אחת. שומר הזמן
// בקוד נשאר גם אחרי המעבר לטריגר-לפי-חלון: אם טריגר הדקות שרד סגירה (כשל,
// התקנה ידנית), הוא לא יתחיל למדוד כל היום.
function watchdogTick() {
  var hour = Number(Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'H'));
  if (hour < WATCHDOG_WINDOW_START_HOUR || hour >= WATCHDOG_WINDOW_END_HOUR) { watchdogCloseWindow(); return 'outside window'; }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return 'previous tick still running';   // מדידה תקועה לא מצטברת
  try { return watchdogMeasure().join(' | '); } finally { lock.releaseLock(); }
}

// מדידה ידנית מהעורך, בכל שעה.
function watchdogNow() { return watchdogMeasure().join(' | '); }

function watchdogMeasure() {
  var stamp = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm:ss');
  var health = timedFetch(WATCHDOG_EXAM_EXEC + '?action=health&origin=examinee-app');
  var deep = timedFetch(WATCHDOG_EXAM_EXEC + '?action=health&deep=1&origin=examinee-app');
  var control = timedFetch(WATCHDOG_CONTROL_EXEC);
  var own = timedOwnSheetRead();
  var sheetMs = deep.json && typeof deep.json.sheetMs === 'number' ? deep.json.sheetMs : -1;
  var build = (health.json && health.json.build) || (deep.json && deep.json.build) || '';
  // totalMs = כמה זמן ההרצה של health&deep=1 ארכה בתוך Apps Script (השרת מודד את עצמו).
  // outside = deep.ms − totalMs = כל מה שאינו ההרצה: הזנקת המכולה, ה-302, הרשת.
  // זה מה שהבדיל ב-14–15/09 בין "המסמך שלנו נתקע" לבין "שער הכניסה של גוגל איטי".
  var totalMs = deep.json && typeof deep.json.totalMs === 'number' ? deep.json.totalMs : -1;
  var outsideMs = (totalMs >= 0 && deep.ms >= 0) ? Math.max(0, deep.ms - totalMs) : -1;
  var row = [stamp, health.ms, health.kind, deep.ms, sheetMs, deep.kind, control.ms, control.kind, own,
    watchdogVerdict(health, deep, sheetMs, control, own), build, totalMs, outsideMs];
  try { getWatchdogSheet().appendRow(row); } catch (e) { Logger.log('watchdog: appendRow failed: ' + e); }
  return row;
}

// kind: 'json' (תשובת JSON תקינה), 'html' (דף שגיאה של גוגל), 'http:<code>', 'error:<msg>'
function timedFetch(url) {
  var t0 = Date.now(), out = { ms: -1, kind: '', json: null };
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    out.ms = Date.now() - t0;
    var code = res.getResponseCode(), body = res.getContentText() || '';
    if (code !== 200) { out.kind = 'http:' + code; return out; }
    try { out.json = JSON.parse(body); out.kind = 'json'; }
    catch (eParse) { out.kind = 'html'; }
  } catch (e) {
    out.ms = Date.now() - t0;
    out.kind = 'error:' + String(e && e.message ? e.message : e).slice(0, 60);
  }
  return out;
}

function timedOwnSheetRead() {
  var t0 = Date.now();
  try { getWatchdogSheet().getRange(1, 1).getValue(); return Date.now() - t0; }
  catch (e) { return -1; }
}

function watchdogVerdict(health, deep, sheetMs, control, own) {
  var slow = function(ms, kind) { return ms < 0 || ms > WATCHDOG_SLOW_MS || (kind && kind !== 'json'); };
  var h = slow(health.ms, health.kind), d = slow(deep.ms, deep.kind), c = slow(control.ms, control.kind);
  var doc = sheetMs < 0 || sheetMs > WATCHDOG_SLOW_MS, o = own < 0 || own > WATCHDOG_SLOW_MS;
  if (!h && !d && !c && !doc && !o) return 'ok';
  if (doc && !h && !c) return 'our-document';
  // dispatch-ours: שתי הבקשות לסקריפט שלנו איטיות, אבל גם המסמך שלנו מהיר וגם
  // סקריפט הבקרה (אותו חשבון, מסמך אחר) מהיר — כלומר התור הוא של הפרויקט הזה
  // (מכולות תפוסות / קוד גדול), לא של גוגל ולא של הגיליון. זה בדיוק המצב של
  // 14–15/09 שבו 'mixed' לא אמר כלום.
  if (h && d && !doc && !c) return 'dispatch-ours';
  if (h && c && !doc) return o ? 'google/account' : 'dispatch';
  if (h && c && doc) return 'google/account';
  if (o && !h && !c) return 'sheets-service';
  return 'mixed';
}

function getWatchdogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(WATCHDOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WATCHDOG_SHEET);
    sheet.getRange(1, 1, 1, WATCHDOG_HEADERS.length).setValues([WATCHDOG_HEADERS]);
    sheet.getRange(1, 1, 1, WATCHDOG_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}
