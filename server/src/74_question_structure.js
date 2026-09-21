var EXAM_STRUCTURE_SERVER = {
  'B':  { 'בטיחות': 7, 'הכרת הרכב': 7, 'חוק': 7, 'תמרורים': 9 },
  '1':  { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 6, 'תמרורים': 6, 'ספציפי': 8 },
  'C1': { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 5, 'תמרורים': 5, 'ספציפי': 10 },
  'C':  { 'בטיחות': 5, 'הכרת הרכב': 4, 'חוק': 3, 'תמרורים': 4, 'ספציפי': 14 },
  'D':  { 'בטיחות': 4, 'הכרת הרכב': 2, 'חוק': 5, 'תמרורים': 4, 'ספציפי': 15 }
};

function classifyCategoryServer(cat) {
  var c = String(cat || '').trim();
  if (/ספציפי/.test(c)) return 'ספציפי'; // ספציפי
  if (/בטיחות/.test(c)) return 'בטיחות'; // בטיחות
  if (/הכרת הרכב/.test(c)) return 'הכרת הרכב'; // הכרת הרכב
  if (/חוק/.test(c)) return 'חוק'; // חוק
  if (/תמרורים/.test(c)) return 'תמרורים'; // תמרורים
  if (/זכות קדימה/.test(c)) return 'חוק'; // זכות קדימה → חוק
  return '';
}

function filterByLicenseServer(pool, license) {
  return pool.filter(function(q) {
    var cat = String(q.category || '');
    // "מתן זכות קדימה" applies to all license types
    if (/זכות קדימה/.test(cat)) return true;
    if (license === '1') {
      var lt = String(q.licenseType || '').trim();
      if (lt !== '' && lt !== 'N/A') return false;
      if (cat.indexOf('1') === -1) return false;
      return true;
    }
    if (license === 'C') {
      var lic = String(q.licenseType || '').trim();
      return lic === 'C' || lic === 'C/E' || lic === 'C+E' || lic === 'CE';
    }
    var lic2 = String(q.licenseType || '').trim();
    return lic2 === license;
  });
}

function shuffleArrayServer(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

