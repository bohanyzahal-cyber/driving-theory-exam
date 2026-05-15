// © 2026 Vitaly Gitelman. All Rights Reserved.
// generate_questions_data.js
//
// Reads questions.js + questions_<lang>.js translation files, merges them
// into per-language JSON files for the server-side question delivery.
//
// Output: deployment/generated/questions_<lang>.json (gitignored)
//
// Usage:  node deployment/generate_questions_data.js
//
// User then uploads each generated file to a private Google Drive folder
// that Apps Script reads via getExamQuestions.

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const OUT_DIR = path.join(__dirname, 'generated');
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

// Languages without a translation file (use Hebrew text directly):
//   he — native
//   ru — no translations file yet, falls back to Hebrew on the client
const TRANSLATION_LANGS = [
  { code: 'en', file: 'questions_en.js', varName: 'TRANSLATIONS_EN' },
  { code: 'ar', file: 'questions_ar.js', varName: 'TRANSLATIONS_AR' },
  { code: 'fr', file: 'questions_fr.js', varName: 'TRANSLATIONS_FR' },
  { code: 'es', file: 'questions_es.js', varName: 'TRANSLATIONS_ES' },
  { code: 'am', file: 'questions_am.js', varName: 'TRANSLATIONS_AM' }
];

function extractLiteral(src, declToken, openChar, closeChar) {
  // Locate the first `openChar` after the var declaration, then take everything
  // up to the LAST matching closeChar in the file. Tolerates trailing code like
  // `window.QUESTIONS = QUESTIONS;` after the literal.
  const declIdx = src.indexOf(declToken);
  if (declIdx === -1) return null;
  const startIdx = src.indexOf(openChar, declIdx);
  if (startIdx === -1) return null;
  const endIdx = src.lastIndexOf(closeChar);
  if (endIdx < startIdx) return null;
  return src.substring(startIdx, endIdx + 1);
}

function loadHebrewQuestions() {
  const p = path.join(ROOT, 'questions.js');
  const src = fs.readFileSync(p, 'utf8');
  const literal = extractLiteral(src, 'var QUESTIONS', '[', ']');
  if (!literal) throw new Error('Could not extract QUESTIONS literal from questions.js');
  // eslint-disable-next-line no-eval
  return eval(literal);
}

function loadTranslations(file, varName) {
  const p = path.join(ROOT, file);
  if (!fs.existsSync(p)) {
    console.warn(`Translation file missing: ${file}`);
    return {};
  }
  const src = fs.readFileSync(p, 'utf8');
  const literal = extractLiteral(src, 'var ' + varName, '{', '}');
  if (!literal) {
    console.warn(`Could not extract ${varName} literal from ${file}`);
    return {};
  }
  // eslint-disable-next-line no-eval
  return eval('(' + literal + ')');
}

// Strip transient/useless fields from a question record before publishing it
// to the server-side dataset. `ci` is intentionally stripped: the real answer
// key lives only in answer_key.gs server-side, so we don't ship it.
function cleanQuestion(q, langCode) {
  return {
    id: q.id,
    text: q.text,
    answers: q.answers,
    category: q.category,
    licenseType: q.licenseType || '',
    imageUrl: q.imageUrl || null,
    language: langCode
  };
}

function writeJson(outFile, data) {
  const json = JSON.stringify(data);
  fs.writeFileSync(outFile, json);
  const kb = (Buffer.byteLength(json, 'utf8') / 1024).toFixed(1);
  console.log(`  wrote ${path.basename(outFile)}  —  ${data.length} questions  —  ${kb} KB`);
}

function main() {
  console.log('Loading questions.js (Hebrew + Russian)…');
  const all = loadHebrewQuestions();
  console.log(`  loaded ${all.length} total entries`);

  // questions.js holds both `he` and `ru` rows interleaved — split them.
  const hebrew = all.filter(q => (q.language || 'he') === 'he');
  const russian = all.filter(q => q.language === 'ru');
  console.log(`  ${hebrew.length} he  /  ${russian.length} ru`);

  writeJson(path.join(OUT_DIR, 'questions_he.json'), hebrew.map(q => cleanQuestion(q, 'he')));
  if (russian.length > 0) {
    writeJson(path.join(OUT_DIR, 'questions_ru.json'), russian.map(q => cleanQuestion(q, 'ru')));
  }

  // For each translation language, produce a merged dataset:
  //   take each Hebrew question that HAS a translation, replace text+answers
  //   with the translation, keep id/category/licenseType/imageUrl intact.
  for (const lang of TRANSLATION_LANGS) {
    console.log(`Building ${lang.code} dataset…`);
    const translations = loadTranslations(lang.file, lang.varName);
    const merged = [];
    for (const q of hebrew) {
      const t = translations[q.id];
      if (!t) continue; // Hebrew questions without this translation are skipped
      merged.push({
        id: q.id,
        text: t.t || q.text,
        answers: Array.isArray(t.a) && t.a.length === q.answers.length ? t.a : q.answers,
        category: q.category,
        licenseType: q.licenseType || '',
        imageUrl: q.imageUrl || null,
        language: lang.code
      });
    }
    writeJson(path.join(OUT_DIR, `questions_${lang.code}.json`), merged);
  }

  console.log('Done.');
}

main();
