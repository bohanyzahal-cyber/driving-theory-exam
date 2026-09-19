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
//      גיליונות + UrlFetch). זה יוצר טריגר של כל דקה ומריץ מדידה ראשונה.
//   3. לוודא שבגיליון "log" מופיעה שורה. מחוץ ל-07:00–13:00 שעון ישראל הטיק חוזר מיד
//      (כמה מאיות שנייה), כדי לשמור על תקציב הטריגרים (90 דק'/יום בחשבון gmail).
// בבוקר בחינות: זה המקום הראשון להסתכל בו, לפני דף הביצועים ולפני גיליון אבחון.

var WATCHDOG_EXAM_EXEC = 'https://script.google.com/macros/s/AKfycbzOI0zrDEngP-GvlRblhOk8tQsYBvWZ2gGliIQHTpS67WrDZl4la8NPpwtJr_Vjsh3Gzg/exec';
var WATCHDOG_CONTROL_EXEC = 'https://script.google.com/macros/s/AKfycbwWtoSd7ivgXZi0luvYmX8FIZdSGevAHfbCmvKJElgU5egF2rPlC7m9f-k6OILbqIFT/exec?action=ping';
var WATCHDOG_WINDOW_START_HOUR = 7;    // שעון ישראל, כולל
var WATCHDOG_WINDOW_END_HOUR = 13;     // לא כולל
var WATCHDOG_SLOW_MS = 5000;           // מעל זה = "איטי" לצורך ה-verdict
var WATCHDOG_SHEET = 'log';
var WATCHDOG_HEADERS = ['זמן (ישראל)', 'health ms', 'health', 'deep ms', 'deep sheetMs', 'deep', 'control ms', 'control', 'ownSheet ms', 'verdict', 'build'];

function installWatchdog() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'watchdogTick') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('watchdogTick').timeBased().everyMinutes(1).create();
  var row = watchdogMeasure();
  Logger.log('installWatchdog: trigger every minute created; first row: ' + row.join(' | '));
  return row;
}

// הטריגר. מחוץ לחלון הבוקר חוזר מיד; בתוכו — מדידה אחת ושורה אחת.
function watchdogTick() {
  var hour = Number(Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'H'));
  if (hour < WATCHDOG_WINDOW_START_HOUR || hour >= WATCHDOG_WINDOW_END_HOUR) return 'outside window';
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
  var row = [stamp, health.ms, health.kind, deep.ms, sheetMs, deep.kind, control.ms, control.kind, own,
    watchdogVerdict(health, deep, sheetMs, control, own), build];
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
