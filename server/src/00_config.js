// © 2026 Vitaly Gitelman. All Rights Reserved.
// Unauthorized copying, modification or distribution is prohibited.
// ===== Google Apps Script — מערכת בחינות חיצונית =====
// הדבק את הקוד הזה ב-Apps Script של גיליון Google Sheets חדש
// Deploy → New deployment → Web app
// Execute as: Me | Who has access: Anyone
// העתק את ה-URL שמקבלים והדבק ב-examiner.html וב-examinee.html

// ========== פונקציות עזר ==========

// The three sheets archiveSheets() moves old rows into (14_pending_archive.js,
// target `reports`). The NAMES live here because 12_reads readHistorySince and
// 22_util read them on the live path and both are `both` — a split deployment
// without the archive job must still know where the history is (DESIGN §13.3).
var PENDING_ARCHIVE_SHEET = 'ממתינים_ארכיון';
var RESULTS_ARCHIVE_SHEET = 'תוצאות_ארכיון';
var EXAMS_ARCHIVE_SHEET = 'מבחנים_ארכיון';

var SHEET_HEADERS = {
  'בוחנים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'מס בוחן', 'תפקיד', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'אתרים מנוהלים'],
  'אתרים': ['שם אתר', 'מזהה', 'טלפון מנהל', 'כיתות'],
  // 15 columns: createSession appends the default population (idx 14) and
  // getSessionInfo / listSessions read it (same drift as ממתינים, review E §5).
  'סשנים': ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה', 'תקף עד', 'פעיל', 'כמויות JSON', 'מאושרים JSON', 'בוחן אחראי', 'אוכלוסיית ברירת מחדל'],
  // 19 columns: the code reads 15-18 and writes 19 (review E §5 — a sheet
  // re-created from a 15-column header would be born four columns short and the
  // warning counter / site / "finished on device" writes would land outside it).
  'ממתינים': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע', 'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'],
  // Rows moved out of ממתינים by archiveSheets (full 19-col width, nothing deleted).
  'ממתינים_ארכיון': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע', 'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'],
  'תוצאות': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'],
  // Rows moved out of תוצאות by archiveSheets (30 days) — same 30 columns.
  'תוצאות_ארכיון': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'],
  // The question map of one attempt. Written only by the exam-start path and
  // read only while that attempt is scored, so it is archived after 2 days.
  'מבחנים': ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'],
  'מבחנים_ארכיון': ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'],
  'הארכות זמן': ['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן'],
  // 10 columns: role (idx 8) and site (idx 9) are read by teacherLogin,
  // teacherVerifyLogin, teacherCommanderDashboard and adminDashboard.
  'מורים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'תפקיד', 'אתר'],
  // 8 columns: createClass appends the site (idx 7) and every commander report reads it.
  'כיתות': ['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'תאריך יצירה', 'פעיל', 'אתר'],
  'כיתות שנמחקו': ['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'אתר', 'תאריך מחיקה'],
  'תלמידי כיתות': ['קוד כיתה', 'שם תלמיד', 'מזהה תלמיד', 'תאריך הצטרפות'],
  'תוצאות תרגול': ['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז', 'עבר/נכשל', 'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא', 'טלפון'],
  'התקדמות תלמידים': ['שם תלמיד', 'קוד כיתה', 'מפתח', 'streak', 'wrong_qs', 'history', 'עדכון אחרון'],
  'חיזוי סיכון': ['חושב בתאריך', 'שם', 'דרגה', 'קוד כיתה', 'מורה ת.ז.', 'שם מורה', 'שם כיתה', 'אתר', 'ציון תרגול', 'תרגולים', 'מגמה', 'ניסיון צפוי', 'ניגש בעבר', 'סיכוי מעבר', 'רמת סיכון', 'ביטחון', 'זוהה בטלפון', 'טלפון']
};

// Sites used ONLY for system testing by examiners (not real exams). Their rows are
// EXCLUDED from the commander dashboard statistics so test data doesn't pollute the
// real numbers. They are NOT filtered from the live examiner dashboard — a tester
// still needs to see their own test session. Add more names here if needed.
var TEST_SITES = ['בדיקת נתונים', 'דימונה דוגית 35'];
function isTestSite(site) {
  return TEST_SITES.indexOf(String(site || '').trim()) !== -1;
}

// Person-name normalizer for fuzzy matching: strip punctuation, collapse spaces,
// lowercase, token-sort (so "ישראל ישראלי" and "ישראלי ישראל" hash the same).
function normalizeNameKey(s) {
  if (!s) return '';
  var t = String(s).replace(/[׳״'".\-]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  if (!t) return '';
  var tokens = t.split(' ').filter(function(x) { return x; });
  tokens.sort();
  return tokens.join(' ');
}

// Examiner identity sets (names + IDs) from the בוחנים sheet, for excluding examiners
// who registered as EXAMINEES to test the system. ID match is exact (no false positives,
// since a ת.ז. is unique to the examiner); name match is fuzzy (normalizeNameKey) and can
// rarely catch a real same-named candidate. Read once per report.
function getExaminerExclusion() {
  var names = {}, ids = {};
  try {
    var d = getSheet('בוחנים').getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var nk = normalizeNameKey(d[i][0]);   // col 0 = שם
      if (nk) names[nk] = true;
      var ik = normalizeId(d[i][1]);         // col 1 = ת.ז.
      if (ik) ids[ik] = true;
    }
  } catch (e) {}
  return { names: names, ids: ids };
}
function isExaminerSelfTest(name, id, excl) {
  if (!excl) return false;
  var ik = normalizeId(id);
  if (ik && excl.ids[ik]) return true;       // ת.ז. match — precise
  var nk = normalizeNameKey(name);
  return !!(nk && excl.names[nk]);           // name match — fuzzy fallback
}

