// ========== handlers ==========

function handleLogin(p) {
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var row = i + 1; // sheet rows are 1-indexed
      // Rate limiting: column I (index 8) = failed attempts, column J (index 9) = lockout until
      var failedAttempts = Number(data[i][8]) || 0;
      var lockoutUntil = data[i][9];
      if (lockoutUntil) {
        var lockoutDate = lockoutUntil instanceof Date ? lockoutUntil : new Date(lockoutUntil);
        if (new Date() < lockoutDate) {
          var minsLeft = Math.ceil((lockoutDate - new Date()) / 60000);
          return jsonResponse({ status: 'error', message: 'החשבון נעול עקב ניסיונות כושלים. נסה שוב בעוד ' + minsLeft + ' דקות' });
        }
        // Lockout expired — reset counter
        failedAttempts = 0;
        sheet.getRange(row, 9).setValue(0);    // column I = failed attempts reset
        sheet.getRange(row, 10).setValue('');   // column J = lockout cleared
      }
      if (String(data[i][2]) === String(p.password)) {
        if (data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE') {
          // Successful login — reset failed attempts
          if (failedAttempts > 0) {
            sheet.getRange(row, 9).setValue(0);    // column I = failed attempts reset
            sheet.getRange(row, 10).setValue('');   // column J = lockout cleared
          }
          // Generate token and store in sheet (columns G=7, H=8 → indices 6,7)
          // Support multiple tokens (multi-device) separated by comma, max 5
          var token = generateToken();
          var expiry = new Date();
          expiry.setHours(expiry.getHours() + 12);
          var existingTokens = String(data[i][6] || '').trim();
          var tokenList = existingTokens ? existingTokens.split(',') : [];
          tokenList.push(token);
          if (tokenList.length > 5) tokenList = tokenList.slice(-5); // keep last 5
          sheet.getRange(row, 7).setValue(tokenList.join(','));   // column G = tokens
          sheet.getRange(row, 8).setValue(expiry);   // column H = expiry
          return jsonResponse({ status: 'ok', examiner: { name: data[i][0], id: normalizeId(data[i][1]), examinerNumber: String(data[i][4] || ''), role: String(data[i][5] || 'בוחן'), token: token } });
        } else {
          return jsonResponse({ status: 'error', message: 'החשבון אינו פעיל' });
        }
      } else {
        // Wrong password — increment failed attempts
        failedAttempts++;
        sheet.getRange(row, 9).setValue(failedAttempts);   // column I = failed attempts
        if (failedAttempts >= 5) {
          var lockout = new Date();
          lockout.setMinutes(lockout.getMinutes() + 15);
          sheet.getRange(row, 10).setValue(lockout);        // column J = lockout until
          return jsonResponse({ status: 'error', message: 'יותר מדי ניסיונות כושלים. החשבון ננעל ל-15 דקות' });
        }
        return jsonResponse({ status: 'error', message: 'סיסמה שגויה' });
      }
    }
  }
  return jsonResponse({ status: 'error', message: 'בוחן לא נמצא' });
}

function handleVerifyLogin(p) {
  if (!p.examinerId || !p.token) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטי אימות', tokenExpired: true });
  }
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.examinerId)) {
      var storedTokens = String(data[i][6] || '').split(',');
      var expiry = data[i][7];
      if (storedTokens.indexOf(p.token) === -1) {
        return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
      }
      if (!expiry) {
        return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
      }
      var expiryDate = expiry instanceof Date ? expiry : new Date(expiry);
      if (new Date() > expiryDate) {
        return jsonResponse({ status: 'error', message: 'פג תוקף ההתחברות', tokenExpired: true });
      }
      if (!(data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE')) {
        return jsonResponse({ status: 'error', message: 'החשבון אינו פעיל' });
      }
      return jsonResponse({ status: 'ok', examiner: { name: data[i][0], id: normalizeId(data[i][1]), examinerNumber: String(data[i][4] || ''), role: String(data[i][5] || 'בוחן'), token: p.token } });
    }
  }
  return jsonResponse({ status: 'error', message: 'בוחן לא נמצא', tokenExpired: true });
}

function handleGetSites() {
  var sheet = getSheet('אתרים');
  var data = sheet.getDataRange().getValues();
  var sites = [];
  for (var i = 1; i < data.length; i++) {
    var classrooms = String(data[i][3] || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    sites.push({
      name: data[i][0],
      id: data[i][1],
      managerPhone: data[i][2],
      classrooms: classrooms
    });
  }
  return jsonResponse({ status: 'ok', sites: sites });
}

function handleListSessions(p) {
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var examinerId = normalizeId(p.examinerId);
  var sitesSheet = getSheet('אתרים');
  var sitesData = sitesSheet.getDataRange().getValues();
  // Build sites lookup for manager phone
  var sitesMap = {};
  for (var s = 1; s < sitesData.length; s++) {
    sitesMap[String(sitesData[s][0]).trim()] = { managerPhone: sitesData[s][2] || '' };
  }
  var sessions = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (normalizeId(data[i][1]) === examinerId) {
      var siteName = String(data[i][3] || '').trim();
      sessions.push({
        code: String(data[i][0]),
        site: data[i][3] || '',
        classroom: data[i][4] || '',
        license: data[i][5] || '',
        language: data[i][6] || 'he',
        audioMode: data[i][7] || 'off',
        created: data[i][8] || '',
        validUntil: data[i][9] || '',
        active: data[i][10] === true || String(data[i][10]).toUpperCase() === 'TRUE',
        quotas: decodeSessionQuotas(data[i][11], data[i][12], data[i][5]),
        // Defensive read — see handleGetSessionInfo comment.
        responsibleExaminer: String((data[i].length > 13 ? data[i][13] : '') || ''),
        defaultPopulation: String((data[i].length > 14 ? data[i][14] : '') || ''),
        managerPhone: sitesMap[siteName] ? sitesMap[siteName].managerPhone : ''
      });
    }
  }
  // Return up to 20 most recent sessions
  return jsonResponse({ status: 'ok', sessions: sessions.slice(0, 20) });
}

