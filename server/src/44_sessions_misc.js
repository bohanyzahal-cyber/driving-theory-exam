// Returns a sorted list of active examiners' names, used by the session-create
// dropdown to pick the בוחן אחראי. Only sends `name` — IDs/roles/tokens
// don't belong on the client. Active = column D in 'בוחנים' is exactly 'כן'
// (the same truthiness check used elsewhere).
//
// RESPONSIBLE_EXAMINER_HIDE_LIST: names to omit from this dropdown even when
// they're marked active in the sheet. Use when an examiner is still active in
// the system (can log in, see their dashboard) but shouldn't be selectable as
// a "responsible examiner" on the work order. Comparison is normalized
// (trim + lowercase + collapsed whitespace) so casing/spacing variants match.
var RESPONSIBLE_EXAMINER_HIDE_LIST = [
  'תומר לוי',
  'אביאור שמעוני',
  'דוד בטיטו'
];
function _normalizeNameForHideList(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function handleListActiveExaminers(p) {
  diagMark('sheet:examiners-list');   // 17/09: 196s for this one small read
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  var hideSet = {};
  for (var hi = 0; hi < RESPONSIBLE_EXAMINER_HIDE_LIST.length; hi++) {
    hideSet[_normalizeNameForHideList(RESPONSIBLE_EXAMINER_HIDE_LIST[hi])] = true;
  }
  var names = [];
  for (var i = 1; i < data.length; i++) {
    var active = data[i][3];
    var isActive = (active === true) || (active === 'כן') || (String(active).toUpperCase() === 'TRUE');
    if (!isActive) continue;
    var name = String(data[i][0] || '').trim();
    if (!name) continue;
    if (hideSet[_normalizeNameForHideList(name)]) continue;
    names.push(name);
  }
  names.sort(function(a, b) { return a.localeCompare(b, 'he'); });
  return jsonResponse({ status: 'ok', examiners: names });
}

// Validates the JSON quotas payload sent from the examiner UI. Returns either
// { rows: [...] } on success or { error: 'msg' }. The same checks are mirrored
// in examiner.html createSessionBtn — kept in sync so a forged client still
// fails server-side.
var QUOTA_VALID_LICENSES = { B:1, '1':1, C1:1, C:1, D:1 };
function parseAndValidateQuotas(raw) {
  if (!raw) return { error: 'יש להזין כמויות נבחנים לפי דרגה' };
  var parsed;
  try { parsed = JSON.parse(raw); } catch(e) { return { error: 'מבנה כמויות לא תקין' }; }
  if (!Array.isArray(parsed) || parsed.length === 0) {
    return { error: 'יש להזין לפחות שורת כמויות אחת' };
  }
  var seen = {};
  var clean = [];
  for (var i = 0; i < parsed.length; i++) {
    var r = parsed[i] || {};
    var lic = String(r.license || '').trim();
    var site = String(r.site || '').trim();  // '' = host site (backward compatible)
    var req = parseInt(r.requested, 10);
    var appr = parseInt(r.approved, 10);
    if (!QUOTA_VALID_LICENSES[lic]) {
      return { error: 'דרגה לא חוקית בשורה ' + (i + 1) };
    }
    // Uniqueness is per (site, license): the same license may appear once per
    // site (host + guest) but not twice for the same site.
    var _qkey = site + '|' + lic;
    if (seen[_qkey]) {
      return { error: 'דרגה "' + lic + '" מופיעה יותר מפעם אחת' + (site ? ' לאתר "' + site + '"' : '') };
    }
    seen[_qkey] = true;
    // Quantity is OPTIONAL (mirrors the client redesign e0708f2, which removed the
    // requested/approved fields from session opening). A missing/blank/0 quantity
    // defaults to 0 instead of blocking session creation — the examiner no longer
    // has to type a number to open an exam.
    if (!isFinite(req) || req < 0) req = 0;
    if (!isFinite(appr) || appr < 0) appr = 0;
    if (appr > req) appr = req;
    clean.push({ site: site, license: lic, requested: req, approved: appr });
  }
  return { rows: clean };
}

// Decode column L into an array of quota rows. Handles three historical shapes:
//   1. Empty cell           → []
//   2. Plain number         → [{license: <session.license>, requested: <num>, approved: <colM>}]
//      (early prototype that stored requested/approved as separate columns L,M)
//   3. JSON array string    → parsed array
// Used by every session reader so backward-compat is centralised.
function decodeSessionQuotas(colL, colM, sessionLicense) {
  if (colL === '' || colL == null) return [];
  var s = String(colL).trim();
  if (s.charAt(0) === '[') {
    try {
      var arr = JSON.parse(s);
      if (Array.isArray(arr)) {
        var out = [];
        for (var i = 0; i < arr.length; i++) {
          var r = arr[i] || {};
          out.push({
            site: String(r.site || ''),
            license: String(r.license || ''),
            requested: Number(r.requested) || 0,
            approved: Number(r.approved) || 0
          });
        }
        return out;
      }
    } catch(e) {}
    return [];
  }
  // Legacy single-pair format: column L = requested, column M = approved
  var legacyReq = parseInt(s, 10);
  var legacyAppr = parseInt(colM, 10);
  if (isFinite(legacyReq) && legacyReq > 0) {
    return [{
      site: '',
      license: String(sessionLicense || 'B'),
      requested: legacyReq,
      approved: isFinite(legacyAppr) ? legacyAppr : 0
    }];
  }
  return [];
}

// Address of the polling Worker (DESIGN §3.4). Empty = examinees poll this
// script directly; setting/clearing the ScriptProperty switches the whole fleet
// within one getSessionInfo, without a Pages deploy.
function gatewayUrl() {
  try { return String(PropertiesService.getScriptProperties().getProperty('GATEWAY_URL') || '').trim(); }
  catch (e) { return ''; }
}

// ---- One 'סשנים' read per execution ----------------------------------------
// addExamTime and disqualify each read the whole sheet twice — once for the
// ownership check, once for the session's examiner name (review C R12). The
// memo lives for one request, which is far shorter than any state it caches.
var _sessionRowsMemo = null;
function sessionRows() {
  if (!_sessionRowsMemo) _sessionRowsMemo = getSheet('סשנים').getDataRange().getValues();
  return _sessionRowsMemo;
}
function sessionRowByCode(sessionCode) {
  var rows = sessionRows(), want = String(sessionCode || '').trim();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === want) return rows[i];
  }
  return null;
}
// Same rule as verifyExaminerForSession (20_auth.js): the session's own
// examiner, and nothing when the session does not exist. Served from the memo
// so the caller's later lookups are free. ⚠ The two must stay in step until the
// auth module is rewritten to take a row.
function examinerOwnsSession(sessionCode, examinerId) {
  if (!examinerId) return false;
  var row = sessionRowByCode(sessionCode);
  return !!row && normalizeId(row[1]) === normalizeId(examinerId);
}

function handleUpdateSession(p) {
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var searchCode = String(p.sessionCode).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchCode && (data[i][10] === true || String(data[i][10]).toUpperCase() === 'TRUE')) {
      if (normalizeId(data[i][1]) !== normalizeId(p.examinerId)) {
        return jsonResponse({ status: 'error', message: 'אין הרשאה לעדכן סשן זה' });
      }
      var row = i + 1;
      if (p.license) sheet.getRange(row, 6).setValue(p.license);
      if (p.language) sheet.getRange(row, 7).setValue(p.language);
      if (p.audioMode) sheet.getRange(row, 8).setValue(p.audioMode);
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'סשן לא נמצא' });
}

function handleCloseSession(p) {
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var searchCode = String(p.sessionCode).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchCode && normalizeId(data[i][1]) === normalizeId(p.examinerId)) {
      sheet.getRange(i + 1, 11).setValue(false);
      var cleanup = cleanupStuckDisqualified(searchCode);
      return jsonResponse({ status: 'ok', cleanup: cleanup });
    }
  }
  return jsonResponse({ status: 'error', message: 'סשן לא נמצא' });
}

// Triggered from handleCloseSession. Sweeps pending rows in this session whose
// status got stuck on 'disqualified' without a real result, and moves them to a
// terminal status so the session closes clean.
//   no result row    → 'cancelled' (DQ fired but nothing was ever recorded)
//   latest is 'בוטל' → 'completed' (a result existed but was overturned)
// Rows whose latest result is 'פסול'/'עבר'/'נכשל' are left as 'disqualified' —
// those are real outcomes awaiting examiner confirm/overturn.
function cleanupStuckDisqualified(sessionCode) {
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var stuck = [];
  for (var i = 1; i < pendData.length; i++) {
    if (String(pendData[i][0]) === String(sessionCode) &&
        String(pendData[i][5] || '').trim() === 'disqualified') {
      stuck.push({ rowIdx: i, idKey: normalizeId(pendData[i][1]) });
    }
  }
  if (stuck.length === 0) return { cancelled: 0, completed: 0, skipped: 0 };

  // A session closes at the end of its day; results older than the retention
  // window cannot belong to it, so the tail is enough.
  var resData = readResultsTail().rows;
  var latestByExaminee = {};
  for (var r = 1; r < resData.length; r++) {
    if (String(resData[r][13]) !== String(sessionCode)) continue;
    // resData is in append order; later row wins as "latest"
    latestByExaminee[normalizeId(resData[r][1])] = String(resData[r][7] || '').trim();
  }

  var cancelled = 0, completed = 0, skipped = 0;
  for (var k = 0; k < stuck.length; k++) {
    var latest = latestByExaminee[stuck[k].idKey];
    if (!latest) {
      setPendingStatus(pendSheet, stuck[k].rowIdx + 1, sessionCode, 'cancelled');
      cancelled++;
    } else if (latest === 'בוטל') {
      setPendingStatus(pendSheet, stuck[k].rowIdx + 1, sessionCode, 'completed');
      completed++;
    } else {
      skipped++;
    }
  }
  return { cancelled: cancelled, completed: completed, skipped: skipped };
}

function handleGetSessionInfo(p) {
  var sheet = getSheet('סשנים');
  var data = sessionRows();
  var searchCode = String(p.sessionCode).trim();
  for (var i = 1; i < data.length; i++) {
    var rowCode = String(data[i][0]).trim();
    if (rowCode === searchCode) {
      var active = data[i][10];
      if (active !== true && active !== 'TRUE' && String(active).toUpperCase() !== 'TRUE') {
        return jsonResponse({ status: 'error', message: 'הסשן הסתיים' });
      }
      var validUntil = new Date(data[i][9]);
      if (new Date() > validUntil) {
        sheet.getRange(i + 1, 11).setValue(false);
        return jsonResponse({ status: 'error', message: 'תוקף הסשן פג' });
      }
      var _siQuotas = decodeSessionQuotas(data[i][11], data[i][12], data[i][5]);
      // Build the distinct site list (host first, then guest sites declared in
      // the quotas). Quota rows with an empty site belong to the host (column D).
      // The examinee picks from this list when more than one site exists.
      var _hostSite = String(data[i][3] || '').trim();
      var _siteSeen = {};
      var _sites = [];
      if (_hostSite) { _sites.push(_hostSite); _siteSeen[_hostSite] = true; }
      for (var _sq = 0; _sq < _siQuotas.length; _sq++) {
        var _sName = String(_siQuotas[_sq].site || '').trim() || _hostSite;
        if (_sName && !_siteSeen[_sName]) { _siteSeen[_sName] = true; _sites.push(_sName); }
      }
      return jsonResponse({
        status: 'ok',
        session: {
          // The client checks `build` to notice an old server behind a new page,
          // and reads `gateway.url` to decide where the examinee polls. An empty
          // url (ScriptProperty GATEWAY_URL unset) means "poll me directly" —
          // that is the kill switch for the Worker, with no Pages push.
          build: THEORY_API_BUILD,
          gateway: { url: gatewayUrl() },
          site: data[i][3],
          sites: _sites,
          classroom: data[i][4],
          license: data[i][5],
          language: data[i][6],
          audioMode: data[i][7],
          examinerName: data[i][2],
          validUntil: data[i][9],
          quotas: _siQuotas,
          // Column N (13) may be missing on rows created before this feature
          // shipped — defensive read returns '' for those, treating them as
          // sessions without a designated responsible examiner.
          responsibleExaminer: String((data[i].length > 13 ? data[i][13] : '') || ''),
          // Default population set by the examiner at session open (col O, idx 14).
          // The examinee's form pre-selects it but can change it. '' on old rows.
          defaultPopulation: String((data[i].length > 14 ? data[i][14] : '') || '')
        }
      });
    }
  }
  return jsonResponse({ status: 'error', message: 'קוד סשן לא תקין' });
}

