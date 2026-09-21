#!/usr/bin/env node
/**
 * fix_bank_images.js — one-time data fix on the 7 source question dumps.
 *
 * WHY: id 120 carries `imageUrl:"realistic_visual_road_demo"` — a placeholder
 * that was never a URL. The old client asked gov.il for it and got a 404; the
 * new static bank stores an image BASENAME, so a non-URL value would become a
 * bogus file name and the build would (rightly) fail. Any imageUrl that is not
 * http(s) is therefore rewritten to "" = "this question has no image".
 *
 * The dumps are single-line compact JSON and round-trip byte-identically
 * through JSON.parse/stringify (verified 2026-09-21), so re-serialising them
 * changes nothing except the fixed values.
 *
 * Idempotent: running it twice reports 0 changes and leaves the files alone.
 * A .bak_preImgFix_2026-09-21 copy is written next to each changed file.
 *
 * Usage: node tools/fix_bank_images.js [--dry]
 */
'use strict';
const fs = require('fs');
const path = require('path');

const LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
const GENERATED = path.join(__dirname, '..', 'deployment', 'generated');
const BACKUP_SUFFIX = '.bak_preImgFix_2026-09-21';
const dry = process.argv.includes('--dry');

function fixOne(lang) {
  const file = path.join(GENERATED, 'questions_' + lang + '.json');
  const before = fs.readFileSync(file, 'utf8');
  const rows = JSON.parse(before);
  const changed = [];
  for (const row of rows) {
    const url = String(row.imageUrl == null ? '' : row.imageUrl);
    if (url === '' || /^https?:\/\//.test(url)) continue;
    changed.push({ id: row.id, licenseType: String(row.licenseType || ''), was: url });
    row.imageUrl = '';
  }
  const after = JSON.stringify(rows);
  const result = {
    lang,
    file,
    bytesBefore: Buffer.byteLength(before),
    bytesAfter: Buffer.byteLength(after),
    changed
  };
  if (!changed.length || dry) return result;
  fs.writeFileSync(file + BACKUP_SUFFIX, before, 'utf8');
  fs.writeFileSync(file, after, 'utf8');
  result.wrote = true;
  return result;
}

function main() {
  let total = 0;
  const ids = new Set();
  for (const lang of LANGS) {
    const r = fixOne(lang);
    total += r.changed.length;
    r.changed.forEach(c => ids.add(c.id));
    console.log(
      r.lang.padEnd(3),
      String(r.bytesBefore).padStart(8) + ' -> ' + String(r.bytesAfter).padStart(8) + ' bytes',
      '| fixed rows: ' + r.changed.length +
      (r.changed.length ? ' (ids ' + [...new Set(r.changed.map(c => c.id))].join(',') +
        ', licenses ' + r.changed.map(c => c.licenseType || '1').join('/') +
        ', was "' + r.changed[0].was + '")' : '') +
      (r.wrote ? ' [written]' : dry ? ' [dry-run]' : '')
    );
  }
  console.log('total rows fixed: ' + total + (total ? ', ids: ' + [...ids].join(',') : ' (already clean)'));
  if (total && !dry) console.log('backups: deployment/generated/questions_<lang>.json' + BACKUP_SUFFIX);
}

main();
