// ---- One scan for "this examinee's current row" ----------------------------
// Ten handlers wrote this reverse loop by hand, which is how their status and
// 'בוטל' filters drifted apart (review E S7, C R12). rows[0] is a header.
//
// findLatestPendingRow: newest ממתינים row of (session, id). `statuses`, when
// given, keeps scanning past rows in other states instead of stopping at the
// first match — that is what "reset every stuck row" and "approve the waiting
// one" need. Returns { idx, row, status }; idx is an index into `rows` (the
// sheet row is idx + 1 + off) and is -1 when nothing matched.
function findLatestPendingRow(rows, sessionCode, idNumber, statuses) {
  var code = String(sessionCode || '').trim(), id = normalizeId(idNumber);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() !== code || normalizeId(rows[i][1]) !== id) continue;
    var st = String(rows[i][5] || '').trim();
    if (statuses && statuses.indexOf(st) === -1) continue;
    return { idx: i, row: rows[i], status: st };
  }
  return { idx: -1, row: null, status: '' };
}

// Newest 'תוצאות' row of (session, id). skipCancelled leaves 'בוטל' rows out:
// a correction must never land on a row that was already overturned (E S7).
function findLatestResultRow(rows, sessionCode, idNumber, skipCancelled) {
  var code = String(sessionCode || '').trim(), id = normalizeId(idNumber);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]).trim() !== code || normalizeId(rows[i][1]) !== id) continue;
    var st = String(rows[i][7] || '').trim();
    if (skipCancelled && st === 'בוטל') continue;
    return { idx: i, row: rows[i], status: st };
  }
  return { idx: -1, row: null, status: '' };
}

function findRow(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) return i + 1;
  }
  return -1;
}

function findAllRows(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) results.push({ row: i + 1, data: data[i] });
  }
  return results;
}

function generateSessionCode() {
  var sessSheet = getSheet('סשנים');
  var data = sessSheet.getDataRange().getValues();
  var existingCodes = {};
  for (var i = 1; i < data.length; i++) {
    // Check ALL session codes (not just active) to prevent data mixing with closed sessions
    existingCodes[String(data[i][0]).trim()] = true;
  }
  // 8-character alphanumeric code (unambiguous chars: no O/0/I/1/L)
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var code;
  do {
    code = '';
    for (var c = 0; c < 8; c++) {
      code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  } while (existingCodes[code]);
  return code;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

