function countAttempts(idNumber, license, resultRows, resultSheet) {
  // A submit already has the complete history. Other callers keep the full read.
  var data = resultRows || getSheet('תוצאות').getDataRange().getValues();
  if (resultRows && resultSheet && resultSheet.getLastRow() !== data.length) {
    return countAttempts(idNumber, license);
  }
  if (resultRows && resultSheet) {
    var matchingHistory = 0;
    for (var h = 1; h < data.length; h++) {
      if (normalizeId(data[h][1]) === normalizeId(idNumber) && String(data[h][4]) === String(license)) matchingHistory++;
    }
    // Many retakes are cheaper to refresh in one request than many tiny reads.
    if (matchingHistory > 4) return countAttempts(idNumber, license);
  }
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(idNumber) && String(data[i][4]) === String(license)) {
      var row = data[i];
      if (resultRows && resultSheet) {
        // An examiner may have overturned an old result without appending a
        // row. Refresh this examinee's history before assigning the attempt.
        var live = resultSheet.getRange(i + 1, 1, 1, 14).getValues()[0];
        if (!live || normalizeId(live[1]) !== normalizeId(idNumber) || String(live[4]) !== String(license) || String(live[13]) !== String(row[13])) {
          return countAttempts(idNumber, license);
        }
        row = live;
      }
      var status = String(row[7] || '').trim();
      if (status === 'בוטל') continue; // overturned DQ is not a real attempt
      count++;
    }
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

