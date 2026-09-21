// ========== Writes to ממתינים shared by several modules ==========
// Every status change of an examinee row goes through setPendingStatus so that
// (a) the per-session snapshot the pollers read is dropped at once — a status
// written past the snapshot was the r23 gap (review C R8), and (b) the flush
// happens exactly once. Column numbers are 1-based getRange columns.
var PENDING_STATUS_COL = 6;
var PENDING_COLS = { status: 6, language: 7, population: 8, license: 9, audio: 10, timeExtension: 11, examStart: 12, token: 13, dqCount: 14, extScreen: 15, warnCount: 16, lastWarning: 17, site: 18, finishedOnDevice: 19 };
// extras: optional { <PENDING_COLS name>: value } written in the same flush.
function setPendingStatus(sheet, rowNumber, sessionCode, status, extras) {
  sheet.getRange(rowNumber, PENDING_STATUS_COL).setValue(status);
  if (extras) {
    for (var name in extras) {
      if (!Object.prototype.hasOwnProperty.call(extras, name) || !PENDING_COLS[name]) continue;
      sheet.getRange(rowNumber, PENDING_COLS[name]).setValue(extras[name]);
    }
  }
  SpreadsheetApp.flush();
  invalidatePendingSnapshot(sessionCode);
}

// Refresh current rows without repeatedly copying the whole growing sheet.
// Full snapshots retain old recovery rows; a changed row count/identity falls
// back to a full read so registration, retakes and maintenance remain visible.
function refreshExamineePendingRows(sheet, rows, sessionCode, idNumber) {
  if (!rows || sheet.getLastRow() !== rows.length) return sheet.getDataRange().getValues();
  var matchingRows = 0;
  for (var m = 1; m < rows.length; m++) {
    if (String(rows[m][0]) === String(sessionCode) && normalizeId(rows[m][1]) === normalizeId(idNumber)) matchingRows++;
  }
  if (matchingRows > 4) return sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) !== String(sessionCode) || normalizeId(rows[i][1]) !== normalizeId(idNumber)) continue;
    var live = sheet.getRange(i + 1, 1, 1, rows[i].length).getValues()[0];
    if (!live || String(live[0]) !== String(sessionCode) || normalizeId(live[1]) !== normalizeId(idNumber)) {
      return sheet.getDataRange().getValues();
    }
    rows[i] = live;
  }
  return rows;
}

// Helper: mark ALL active pending rows for this session+ID as completed.
// Closes EVERY in_exam/approved row (not just the latest) — a duplicate pending
// row otherwise leaves the soldier stuck on the board even though they finished
// and submitted (reported: "stuck in ממתינים/במבחן despite finishing").
function markPendingCompleted(sessionCode, idNumber, pendingSnapshot) {
  var pendSheet = pendingSnapshot ? pendingSnapshot.sheet : getSheet('ממתינים');
  var pendData = pendingSnapshot
    ? refreshExamineePendingRows(pendSheet, pendingSnapshot.rows, sessionCode, idNumber)
    : pendSheet.getDataRange().getValues();
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(sessionCode) && normalizeId(pendData[j][1]) === normalizeId(idNumber) && (String(pendData[j][5]).trim() === 'in_exam' || String(pendData[j][5]).trim() === 'approved')) {
      pendSheet.getRange(j + 1, 6).setValue('completed');
      pendData[j][5] = 'completed';
    }
  }
}

