function handleListAllSessions(p) {
  // Token already verified upstream (in examinerActions allowlist). Add a role
  // check here since the action isn't restricted by ownership.
  var role = getExaminerRole(p.examinerId);
  if (role !== 'מפקד') {
    return jsonResponse({ status: 'error', message: 'פעולה זו זמינה רק למפקדים' });
  }
  diagMark('sheet:sessions-list');
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var sitesSheet = getSheet('אתרים');
  var sitesData = sitesSheet.getDataRange().getValues();
  var sitesMap = {};
  for (var s = 1; s < sitesData.length; s++) {
    sitesMap[String(sitesData[s][0]).trim()] = { managerPhone: sitesData[s][2] || '' };
  }
  var now = new Date();
  var sessions = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var active = data[i][10] === true || String(data[i][10]).toUpperCase() === 'TRUE';
    if (!active) continue;
    var validUntil = data[i][9] ? new Date(data[i][9]) : null;
    if (validUntil && now > validUntil) continue;
    var siteName = String(data[i][3] || '').trim();
    sessions.push({
      code: String(data[i][0]),
      examinerId: normalizeId(data[i][1]),
      examinerName: data[i][2] || '',
      site: data[i][3] || '',
      classroom: data[i][4] || '',
      license: data[i][5] || '',
      language: data[i][6] || 'he',
      audioMode: data[i][7] || 'off',
      created: data[i][8] || '',
      validUntil: data[i][9] || '',
      active: true,
      quotas: decodeSessionQuotas(data[i][11], data[i][12], data[i][5]),
      managerPhone: sitesMap[siteName] ? sitesMap[siteName].managerPhone : ''
    });
  }
  // Cap response size — newest first (we already iterate in reverse)
  return jsonResponse({ status: 'ok', sessions: sessions.slice(0, 100) });
}

function handleCreateSession(p) {
  var sheet = getSheet('סשנים');
  var code = generateSessionCode();
  var now = new Date();
  var validUntil = new Date(now.getTime() + 8 * 60 * 60 * 1000);

  // Lookup examiner name (normalize ID to handle leading zeros)
  var exSheet = getSheet('בוחנים');
  var exData = exSheet.getDataRange().getValues();
  var exRow = -1;
  var examinerName = '';
  for (var ei = 1; ei < exData.length; ei++) {
    if (normalizeId(exData[ei][1]) === normalizeId(p.examinerId)) { exRow = ei + 1; examinerName = exData[ei][0]; break; }
  }

  // Column L: per-license quotas, stored as JSON. Array of rows like:
  //   [{license:'B', requested:20, approved:18}, {license:'C1', requested:5, approved:5}]
  // Mirrors the plan table in the examiner report — one quota row per license.
  // Column M is reserved (was approvedCount in the previous single-pair design;
  // kept blank now to leave room for future extension without renumbering).
  var quotas = parseAndValidateQuotas(p.quotas);
  if (quotas.error) {
    return jsonResponse({ status: 'error', message: quotas.error });
  }

  // r35 (KNOWN_ISSUES #44): a site that moved to the new system opens no new
  // session here — neither as the host site nor as a guest site in the quotas.
  // Checked before anything is written; open sessions are not affected.
  var sessionSites = [p.site || ''];
  for (var qs = 0; qs < quotas.rows.length; qs++) sessionSites.push(quotas.rows[qs].site || '');
  var movedErr = movedSiteRefusal(sessionSites);
  if (movedErr) return movedErr;

  // Column N (13): בוחן אחראי — name of the senior/responsible examiner when
  // multiple examiners work the same site/day per the פקודת עבודה. When the
  // session is opened by a solo examiner this can equal the examiner himself,
  // or be left blank if he's the responsible. The Rav-Bochen / commander
  // reports surface this field so the chain of responsibility matches the
  // physical staffing on the ground.
  var responsibleExaminer = String(p.responsibleExaminer || '').trim();

  // r35.2: every caller-sent cell as text (cellSafeRow). The site, classroom and
  // examiner name are copied from here into every result row of the session.
  sheet.appendRow(cellSafeRow([
    code,
    p.examinerId,
    examinerName,
    p.site || '',
    p.classroom || '',
    p.license || 'B',
    p.language || 'he',
    p.audioMode || 'off',
    now.toISOString(),
    validUntil.toISOString(),
    true,
    JSON.stringify(quotas.rows),
    '',
    responsibleExaminer,
    String(p.defaultPopulation || '').trim()   // O (idx 14) = default population for the session
  ]));

  return jsonResponse({
    status: 'ok',
    sessionCode: code,
    validUntil: validUntil.toISOString(),
    examinerName: examinerName,
    responsibleExaminer: responsibleExaminer
  });
}

