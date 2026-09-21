// Shared Google-services mock for the server tests.
//
// The built external_exam_apps_script.js is run inside a vm context whose
// SpreadsheetApp / CacheService / PropertiesService / Utilities are fakes that
// COUNT what they were asked to do. Cost is a correctness property here: the
// exam-day outages were reads, not logic, so a test that asserts "one tail read
// and one append" is as important as one that asserts the score.
//
//   const { createEnv } = require('./helpers/server_env.cjs');
//   const env = createEnv({ sheets: { 'ממתינים': rows }, properties: {}, sources: ['deployment/answer_key.gs'] });
//   env.ctx.handleStartExam({ ... });
//   env.sheet('מבחנים').rows        // what was written
//   env.counters()                  // { fullReads, cellsRead, appends, setValues }
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { randomUUID, createHmac, createHash } = require('node:crypto');

const ROOT = path.resolve(__dirname, '..', '..');
const SERVER_FILE = path.join(ROOT, 'external_exam_apps_script.js');

function createEnv(options) {
  const opts = options || {};
  const clock = { t: opts.now || Date.parse('2026-09-22T06:30:00Z') };
  const entries = new Map();
  const properties = new Map(Object.entries(opts.properties || {}));
  const logs = [];
  const spreadsheets = new Map();

  const RealDate = Date;
  function FakeDate(...args) { return args.length ? new RealDate(...args) : new RealDate(clock.t); }
  FakeDate.now = () => clock.t;
  FakeDate.parse = RealDate.parse;
  FakeDate.UTC = RealDate.UTC;
  FakeDate.prototype = RealDate.prototype;

  const cache = {
    get(key) { const e = entries.get(key); return e && e.expires > clock.t ? e.value : null; },
    getAll(keys) { return Object.fromEntries(keys.map(k => [k, this.get(k)]).filter(([, v]) => v !== null)); },
    put(key, value, ttl = 600) { entries.set(key, { value, expires: clock.t + ttl * 1000 }); },
    putAll(values, ttl) { for (const [k, v] of Object.entries(values)) this.put(k, v, ttl); },
    remove(key) { entries.delete(key); },
    removeAll(keys) { for (const key of keys) entries.delete(key); }
  };

  function makeSheet(ss, name) {
    if (ss.sheets.has(name)) return ss.sheets.get(name);
    const sheet = {
      name, rows: [], fullReads: 0, rangeReads: 0, cellsRead: 0, appends: 0, setValues: 0, deletedRows: 0,
      getName() { return this.name; },
      setName(n) { ss.sheets.delete(this.name); this.name = n; ss.sheets.set(n, this); return this; },
      appendRow(row) { this.appends++; this.rows.push(row.slice()); },
      getLastRow() { return this.rows.length; },
      getLastColumn() { return this.rows.reduce((w, r) => Math.max(w, r.length), 0); },
      getMaxRows() { return Math.max(1000, this.rows.length); },
      getMaxColumns() { return Math.max(30, this.getLastColumn()); },
      insertRowsAfter() {}, insertColumnsAfter() {},
      // Real deletion (1-based, like Sheets): the archive job deletes bottom-up
      // and its whole correctness argument is that the remaining row numbers
      // stay valid, which a no-op cannot test.
      deleteRows(start, count) {
        const n = count || 1;
        this.rows.splice(start - 1, n);
        this.deletedRows += n;
      },
      copyTo(target) { const c = makeSheet(target, 'Copy of ' + this.name); c.rows = this.rows.map(r => r.slice()); return c; },
      getRange(startRow, startCol, numRows, numCols) {
        const self = this;
        const height = numRows || 1, width = numCols || 1;
        return {
          setValue(value) {
            self.setValues++;
            const row = self.rows[startRow - 1];
            if (row) row[startCol - 1] = value;
            return this;
          },
          setValues(values) {
            self.setValues += values.length;
            for (let i = 0; i < values.length; i++) {
              const row = self.rows[startRow - 1 + i] || (self.rows[startRow - 1 + i] = []);
              for (let j = 0; j < values[i].length; j++) row[startCol - 1 + j] = values[i][j];
            }
            return this;
          },
          setFontWeight() { return this; },
          getValue() { self.rangeReads++; self.cellsRead++; const row = self.rows[startRow - 1]; return row ? row[startCol - 1] : ''; },
          getValues() {
            self.rangeReads++;
            self.cellsRead += height * width;
            const out = [];
            for (let i = 0; i < height; i++) {
              const row = self.rows[startRow - 1 + i] || [];
              const slice = [];
              for (let j = 0; j < width; j++) slice.push(row[startCol - 1 + j] === undefined ? '' : row[startCol - 1 + j]);
              out.push(slice);
            }
            return out;
          }
        };
      },
      getDataRange() {
        const self = this;
        return {
          getValues() { self.fullReads++; self.cellsRead += self.rows.length * self.getLastColumn(); return self.rows; },
          getFormulas() { return self.rows.map(r => r.map(() => '')); }
        };
      }
    };
    ss.sheets.set(name, sheet);
    return sheet;
  }
  function makeSpreadsheet(id, ssName) {
    const ss = {
      id, ssName, sheets: new Map(),
      getId: () => id, getName: () => ssName, getUrl: () => 'https://example.invalid/' + id,
      getSheetByName: n => ss.sheets.get(n) || null,
      insertSheet: n => makeSheet(ss, n),
      getSheets: () => [...ss.sheets.values()],
      deleteSheet: sh => { ss.sheets.delete(sh.name); }
    };
    spreadsheets.set(id, ss);
    return ss;
  }

  const active = makeSpreadsheet('active', 'exam');
  for (const [name, rows] of Object.entries(opts.sheets || {})) {
    makeSheet(active, name).rows = rows.map(r => r.slice());
  }

  const blob = data => {
    const bytes = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
    return { getBytes: () => [...bytes], getDataAsString: () => bytes.toString('utf8') };
  };
  const triggers = [];
  const ctx = {
    Date: FakeDate,
    Logger: { log: s => logs.push(String(s)) },
    console,
    CacheService: { getScriptCache: () => cache },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperties: () => Object.fromEntries(properties),
        getProperty: k => (properties.has(k) ? properties.get(k) : null),
        setProperty: (k, v) => { properties.set(k, String(v)); },
        deleteProperty: k => { properties.delete(k); }
      })
    },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Session: { getScriptTimeZone: () => 'Asia/Jerusalem' },
    Utilities: {
      getUuid: randomUUID,
      newBlob: blob,
      sleep: ms => { clock.t += ms; },
      formatDate: d => new RealDate(d ? d.getTime() : clock.t).toISOString(),
      base64Encode: b => Buffer.from(b).toString('base64'),
      base64EncodeWebSafe: b => Buffer.from(typeof b === 'string' ? b : Buffer.from(b)).toString('base64url'),
      base64Decode: s => [...Buffer.from(s, 'base64')],
      computeHmacSha256Signature: (value, key) => [...createHmac('sha256', key).update(String(value)).digest()],
      // Apps Script returns SIGNED bytes; sessionSnapshot's hex encoder has to
      // undo that, so the mock must be signed too or the test would pass on a
      // hash the real runtime never produces.
      DigestAlgorithm: { SHA_256: 'SHA_256', MD5: 'MD5' },
      Charset: { UTF_8: 'UTF_8' },
      computeDigest: (algorithm, value) => [...createHash(algorithm === 'MD5' ? 'md5' : 'sha256')
        .update(String(value), 'utf8').digest()].map(b => (b > 127 ? b - 256 : b))
    },
    ScriptApp: {
      newTrigger: fn => ({
        timeBased: () => ({
          after: () => ({ create: () => { const t = { fn, getHandlerFunction: () => fn }; triggers.push(t); return t; } }),
          everyHours: n => ({ create: () => { const t = { fn, everyHours: n, getHandlerFunction: () => fn }; triggers.push(t); return t; } }),
          everyMinutes: n => ({ create: () => { const t = { fn, everyMinutes: n, getHandlerFunction: () => fn }; triggers.push(t); return t; } }),
          // atHour(h).everyDays(1)[.inTimezone(tz)].create() — the nightly jobs
          // pin the time zone, so inTimezone has to be chainable both ways.
          atHour: hour => ({
            everyDays: days => {
              const spec = {
                create: () => { const t = { fn, hour, days, tz: spec.tz, getHandlerFunction: () => fn }; triggers.push(t); return t; },
                inTimezone: tz => { spec.tz = tz; return spec; }
              };
              return spec;
            }
          })
        })
      }),
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: t => { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); },
      getService: () => ({ getUrl: () => 'https://script.invalid/exec' })
    },
    SpreadsheetApp: {
      flush() {},
      getActiveSpreadsheet: () => active,
      openById: id => { const ss = spreadsheets.get(id); if (!ss) throw new Error('no spreadsheet ' + id); return ss; },
      create: ssName => makeSpreadsheet('ss' + (spreadsheets.size + 1), ssName)
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: s => ({ _s: s, setMimeType() { return this; }, getContent() { return this._s; } })
    },
    HtmlService: { createHtmlOutput: s => ({ _s: s, getContent() { return this._s; } }) },
    MimeType: { JSON: 'application/json' }
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(SERVER_FILE, 'utf8'), ctx, { filename: 'external_exam_apps_script.js' });
  for (const extra of opts.sources || []) {
    vm.runInContext(fs.readFileSync(path.isAbsolute(extra) ? extra : path.join(ROOT, extra), 'utf8'), ctx, { filename: String(extra) });
  }

  const sheet = name => makeSheet(active, name);
  const counters = () => {
    const total = { fullReads: 0, rangeReads: 0, cellsRead: 0, appends: 0, setValues: 0, perSheet: {} };
    for (const s of active.sheets.values()) {
      total.fullReads += s.fullReads; total.rangeReads += s.rangeReads; total.cellsRead += s.cellsRead;
      total.appends += s.appends; total.setValues += s.setValues;
      total.perSheet[s.name] = { fullReads: s.fullReads, rangeReads: s.rangeReads, cellsRead: s.cellsRead, appends: s.appends, setValues: s.setValues };
    }
    // reads = every call that crosses the Sheets service to FETCH data, which is
    // the unit the exam-day incidents were counted in (a tail read is two: the
    // header and the tail).
    total.reads = total.fullReads + total.rangeReads;
    return total;
  };
  const resetCounters = () => {
    for (const s of active.sheets.values()) { s.fullReads = 0; s.rangeReads = 0; s.cellsRead = 0; s.appends = 0; s.setValues = 0; }
  };

  return {
    ctx, cache, entries, properties, logs, clock, triggers, spreadsheets, active,
    sheet, sheets: active.sheets, counters, resetCounters,
    rows: name => sheet(name).rows,
    json: out => (out && typeof out.getContent === 'function' ? JSON.parse(out.getContent()) : out)
  };
}

module.exports = { createEnv, SERVER_FILE, ROOT };
