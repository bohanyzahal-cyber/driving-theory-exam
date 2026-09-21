// Key prefix shared by every CacheService / ScriptProperties entry this script
// owns (pending snapshots, extra-minutes maps, token verdicts, diagnostics).
// Declared here — a module that nothing in the roadmap deletes — so the live
// keys keep their names no matter which subsystem is retired; the question
// cache declares the same literal value for as long as it exists.
var CACHE_KEY_PREFIX = 'qv2_';

// Attempt number = how many non-'בוטל' result rows this examinee already has
// for this licence — in the LIVE sheet and in the archive (B5: 'תוצאות' is
// archived after 30 days, and without the archive a retake three months later
// would be recorded as attempt 1).
// `liveRows` is an optional rows array the caller already holds (rows[0] =
// header); without it we read the live sheet ourselves in the three columns
// this needs: B (ת.ז.), E (דרגה), H (עבר/נכשל). The old version re-read the
// whole sheet up to three times per call to re-validate its own input.
var RESULTS_ATTEMPT_COLSPEC = [[2, 1], [5, 1], [8, 1]];
function countAttempts(idNumber, license, liveRows) {
  var wantId = normalizeId(idNumber), wantLic = String(license);
  var count = countAttemptRows(liveRows || readAttemptColumns(getSheet('תוצאות')), wantId, wantLic);
  var arch = getSheetIfExists(RESULTS_ARCHIVE_SHEET);
  if (arch) count += countAttemptRows(readAttemptColumns(arch), wantId, wantLic);
  return count;
}
// Live + archive attempt columns as ONE table (oldest first), for a caller that
// counts attempts for SEVERAL examinees in one request — the dashboard's
// reconciliation used to re-read the whole 'תוצאות' sheet once per stale row
// (review C R1: 40 stale rows = 6.8 M cells in one poll).
function readAttemptHistory() {
  var live = readAttemptColumns(getSheet('תוצאות'));
  var arch = getSheetIfExists(RESULTS_ARCHIVE_SHEET);
  if (!arch) return live;
  var archRows = readAttemptColumns(arch);
  var header = live.length ? live.slice(0, 1) : archRows.slice(0, 1);
  return header.concat(archRows.slice(1), live.slice(1));
}

function readAttemptColumns(sheet) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return readSheetSlice(sheet, 1, lastRow, lastCol, RESULTS_ATTEMPT_COLSPEC);
}
function countAttemptRows(rows, wantId, wantLic) {
  var count = 0;
  for (var i = 1; i < rows.length; i++) {
    if (normalizeId(rows[i][1]) !== wantId || String(rows[i][4]) !== wantLic) continue;
    if (String(rows[i][7] || '').trim() === 'בוטל') continue;   // overturned DQ is not a real attempt
    count++;
  }
  return count;
}

function formatPhoneForWA(phone) {
  phone = String(phone || '').replace(/[^0-9]/g, '');
  if (phone.charAt(0) === '0') phone = '972' + phone.substring(1);
  else if (phone.length === 9 && phone.charAt(0) === '5') phone = '972' + phone;
  return phone;
}

function normalizeId(val) {
  var s = String(val || '').replace(/[^0-9]/g, '');
  while (s.length < 9) s = '0' + s;
  return s;
}

// "מפקד קד״ץ" gets typed in many forms: with Hebrew gershayim ״, ASCII " or ',
// no separator at all ("מפקד קדץ"), with extra spaces. Match all of them so a
// sheet entry typed casually still resolves to the role.
function isKdtzRole(role) {
  return /^\s*מפקד\s+קד[\s׳״'"]*ץ\s*$/.test(String(role || ''));
}

function nowISO() {
  return new Date().toISOString();
}

function todayStr() {
  var d = new Date();
  var dd = ('0' + d.getDate()).slice(-2);
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  var yyyy = d.getFullYear();
  var hh = ('0' + d.getHours()).slice(-2);
  var mi = ('0' + d.getMinutes()).slice(-2);
  return dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi;
}

