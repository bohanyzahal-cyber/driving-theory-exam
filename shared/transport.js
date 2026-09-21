// shared/transport.js — the ONE client transport layer, loaded by examinee.html,
// examiner.html, teacher.html, student.html and exam.html before their inline
// script. ES5 only (classroom PCs run old Edge/Chrome; iPads run old Safari).
//
// What it owns, so that no page re-implements it:
//   1. fetchJsonWithTimeout — a bounded fetch that classifies every failure
//      (timeout / network / http / nonjson) and feeds the transport-health state.
//   2. Transport health — our server always answers HTTP 200 + JSON. Anything else
//      (Google's HTML error page, 404/429/5xx from its front door) means the
//      BACKEND is degraded and the whole fleet must slow down; two consecutive
//      timeouts count too (a ~93 s Sheets hang looks like no network from here).
//   3. Poll pacing — jitter, x1.5 backoff on slow/failed answers, one step down per
//      healthy answer, a 30–60 s floor while degraded.
//   4. createPollLoop — answer → wait → ask again, with an in-flight guard (a
//      restart while a request is still out must never start a second chain),
//      a generation counter, a 60 s poll deadline, and the long-poll gap: when
//      the answer says the server HELD the request, the wait already happened
//      on the wire, so the next request goes out at once. A single-chain page
//      (the examiner dashboard) can opt out of the fleet ladder with steady:true.
//   5. createApi — GET/POST helpers with an always-forwarded timeout and a page
//      supplied decorator that attaches credentials.
//   7. createUpdateCheck — version.json based (per-page content hash), never an
//      ETag: a Pages push that did not change this page shows nothing, and a
//      proxy that rewrites headers cannot fake a new version.
//   8. A small client log ring that a page attaches to its next important POST,
//      so an exam-day incident finally has a client-side trace.
//
// Nothing here touches the DOM or page state. Pages decide what to show.
(function (global) {
  'use strict';

  var API_TIMEOUT_MS = 30000;            // what a person is waiting for on screen
  var POLL_TIMEOUT_MS = 60000;           // unattended polls: keep a slow answer, abandon a hung one
  var CRITICAL_POST_TIMEOUT_MS = 60000;  // startExam / submitResult on a cold container
  var SLOW_ANSWER_MS = 6000;             // above this a poll answer counts as "slow" for pacing
  var POLL_JITTER = 0.3;
  var POLL_BACKOFF_FACTOR = 1.5;
  var POLL_DEGRADED_MIN_MS = 30000;
  var POLL_DEGRADED_MAX_MS = 60000;
  var TIMEOUTS_UNTIL_DEGRADED = 2;
  var LOG_MAX_ENTRIES = 50;
  var LOG_MAX_BYTES = 2048;

  // ---------- long polling ----------
  // The session gateway can HOLD a poll until the answer for this examinee
  // actually changes (the examiner's decision), up to LONGPOLL_WAIT_SEC, and
  // answers unchanged when the hold expires. One request per ~25 s instead of
  // one per 2-6 s is a ~5x cut in Worker requests, and the decision reaches the
  // examinee in about a second instead of in a poll interval.
  // The page asks for the hold (wait + fp in the URL); the loop only needs to
  // know that the answer came back HELD, because then the wait already happened
  // on the wire and the next request should go out almost at once.
  var LONGPOLL_WAIT_SEC = 25;               // what the page asks the server to hold for
  var LONGPOLL_GAP_MS = 250;                // ...and how long we wait before asking again
  var LONGPOLL_TIMEOUT_MS = 25000 + 15000;  // request deadline while holding (POLL_TIMEOUT_MS stays the ceiling)

  // ---------- 1. bounded fetch ----------
  // noteHealth=false is for endpoints that are NOT the exam backend (the
  // version.json probe on GitHub Pages): a 404 there says nothing about Apps
  // Script, and letting it mark the backend degraded would slow every poll in
  // the page to the 30-60 s floor for no reason.
  function boundedFetch(url, opts, timeoutMs, noteHealth) {
    opts = opts || {};
    return new Promise(function (resolve, reject) {
      var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
      if (controller) opts.signal = controller.signal;
      var settled = false;
      var timer = setTimeout(function () {
        var err = new Error('Request timed out');
        err.name = 'TimeoutError'; err.transport = 'timeout';
        finish(err);
        if (controller) { try { controller.abort(); } catch (e) {} }
      }, timeoutMs || API_TIMEOUT_MS);
      function finish(err, data) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (noteHealth !== false) noteTransport(err || null);
        if (err) reject(err); else resolve(data);
      }
      Promise.resolve().then(function () { return fetch(url, opts); }).then(function (response) {
        return response.text().then(function (body) {
          if (!response.ok) {
            var httpErr = new Error('HTTP ' + response.status);
            httpErr.name = 'HttpError'; httpErr.transport = 'http'; httpErr.status = response.status;
            throw httpErr;
          }
          try { return JSON.parse(body); }
          catch (parseErr) {
            var jsonErr = new Error('Non-JSON response');
            jsonErr.name = 'SyntaxError'; jsonErr.transport = 'nonjson'; jsonErr.bodyHead = String(body || '').slice(0, 120);
            throw jsonErr;
          }
        });
      }).then(function (data) { finish(null, data); }, function (err) {
        if (err && !err.transport) { err.transport = (err.name === 'AbortError') ? 'timeout' : 'network'; }
        finish(err);
      });
    });
  }
  function fetchJsonWithTimeout(url, opts, timeoutMs) { return boundedFetch(url, opts, timeoutMs, true); }
  // Same bounded fetch for endpoints that are NOT Apps Script — the version.json
  // probe on Pages and the question-bank Worker. A slow Worker is not a slow
  // backend, and letting it degrade the health state would floor every poll in
  // the page at 30-60 s while Apps Script is perfectly healthy.
  function fetchJsonQuiet(url, opts, timeoutMs) { return boundedFetch(url, opts, timeoutMs, false); }

  // ---------- 2. transport health ----------
  var transportState = 'ok';   // 'ok' | 'degraded'
  var transportDegradedSince = 0;
  var consecutiveTimeouts = 0;
  function noteTransport(err) {
    if (!err) { transportState = 'ok'; transportDegradedSince = 0; consecutiveTimeouts = 0; return; }
    var degrade = false;
    if (err.transport === 'http' || err.transport === 'nonjson') degrade = true;
    if (err.transport === 'timeout') { consecutiveTimeouts++; if (consecutiveTimeouts >= TIMEOUTS_UNTIL_DEGRADED) degrade = true; }
    else consecutiveTimeouts = 0;   // a network error is this device, not the backend
    if (degrade) {
      if (transportState !== 'degraded') transportDegradedSince = Date.now();
      transportState = 'degraded';
    }
  }
  function isBackendDegraded() { return transportState === 'degraded'; }
  function degradedSince() { return transportDegradedSince; }

  // ---------- 3. pacing ----------
  function jitterMs(ms) {
    return Math.round(ms * (1 - POLL_JITTER + 2 * POLL_JITTER * Math.random()));
  }
  function nextPollDelay(currentMs, baseMs, maxMs, worse) {
    if (worse) return Math.min(Math.round(currentMs * POLL_BACKOFF_FACTOR), maxMs);
    return Math.max(baseMs, Math.round(currentMs / POLL_BACKOFF_FACTOR));
  }
  function pacePoll(currentMs, baseMs, maxMs, worse) {
    if (isBackendDegraded()) {
      var slowed = nextPollDelay(currentMs, baseMs, POLL_DEGRADED_MAX_MS, worse);
      return Math.min(Math.max(slowed, POLL_DEGRADED_MIN_MS), POLL_DEGRADED_MAX_MS);
    }
    return nextPollDelay(currentMs, baseMs, maxMs, worse);
  }

  // ---------- 4. poll loop ----------
  // opts: { name, baseMs, maxMs, tick: function() -> Promise|any,
  //         steady: optional true — see "steady mode" below,
  //         nextDelay: optional function(info) -> ms override (e.g. a 2 s sync cadence),
  //         onSettled: optional function(info),
  //         onRestart: optional function() — called just before restartIfStuck
  //           revives the chain, so the page can invalidate its own generation
  //           counter before the fresh tick captures it }
  // info: { ok, slow, elapsedMs, failed, held, error, result }
  // tick() rejecting or returning {status:'error'} counts as failed; the loop never dies.
  // A resolved answer carrying held > 0 is a long-poll answer: see LONGPOLL_GAP_MS.
  //
  // ---- steady mode (opt-in, 2026-09-21) ----
  // WHY: the x1.5 ladder and the 30-60 s degraded floor exist to stop a FLEET
  // from stampeding — hundreds of examinee devices polling Apps Script directly,
  // released together, retrying harder exactly when the backend is already on
  // its knees. A page that runs ONE chain (answer -> wait -> ask again, never
  // two requests in flight) cannot stampede: when the server takes 8 s to
  // answer, the answer time IS the throttle, and slowing down on top of it only
  // delays the person watching the screen — which is how an examiner ends up
  // pressing F5 to see a result that already landed. Such a loop opts in with
  // steady:true and every outcome (ok, slow, failed, error) schedules the next
  // tick at baseMs, jittered as always. A held answer still uses LONGPOLL_GAP_MS,
  // nextDelay still gets the last word, info.slow / info.failed are still
  // reported so the page can SAY the server is slow, and the in-flight guard and
  // generation counter are untouched. Loops WITHOUT steady behave exactly as
  // before — the ladder stays the default for anything a fleet runs.
  function createPollLoop(opts) {
    var timer = null, running = false, gen = 0, inFlight = false, inFlightSince = 0;
    var delayMs = opts.baseMs;
    function schedule(ms) {
      if (!running) return;
      timer = setTimeout(run, jitterMs(ms));
    }
    function run() {
      if (!running) return;
      var myGen = gen;
      timer = null;
      inFlight = true; inFlightSince = Date.now();
      var t0 = inFlightSince;
      var result;
      try { result = opts.tick(); } catch (e) { result = Promise.reject(e); }
      Promise.resolve(result).then(function (data) { return { ok: !(data && data.status === 'error'), result: data, error: null }; },
                                   function (err) { return { ok: false, result: null, error: err }; })
        .then(function (outcome) {
          inFlight = false;
          if (!running || myGen !== gen) return;   // stopped or restarted meanwhile: this chain is dead
          var elapsed = Date.now() - t0;
          // A HELD answer (the server kept the request open until this client's
          // answer changed, or until the hold expired) is not a slow answer: the
          // wait is what we asked for. It is also the whole interval — the next
          // request goes out after LONGPOLL_GAP_MS instead of baseMs, so the
          // next hold starts immediately and a decision is never more than a
          // gap away. Failures are untouched: a failed long poll backs off
          // exactly like any other failed poll, degraded pacing included.
          var held = (outcome.ok && outcome.result && typeof outcome.result.held === 'number' && outcome.result.held > 0) ? outcome.result.held : 0;
          var info = { ok: outcome.ok, failed: !outcome.ok, slow: !held && elapsed > SLOW_ANSWER_MS, elapsedMs: elapsed, held: held, error: outcome.error, result: outcome.result };
          // steady mode: one chain per page, so the answer time is the throttle
          // and the ladder/floor would only add a wait on top of a wait.
          delayMs = opts.steady ? opts.baseMs : pacePoll(delayMs, opts.baseMs, opts.maxMs, info.slow || info.failed);
          var next = held ? LONGPOLL_GAP_MS : delayMs;
          if (opts.nextDelay) { var override = opts.nextDelay(info, next); if (typeof override === 'number' && override > 0) next = override; }
          if (opts.onSettled) { try { opts.onSettled(info); } catch (e) {} }
          schedule(next);
        });
    }
    return {
      name: opts.name,
      start: function () { if (running) this.stop(); running = true; gen++; delayMs = opts.baseMs; run(); },
      stop: function () { running = false; gen++; if (timer) { clearTimeout(timer); timer = null; } },
      isRunning: function () { return running; },
      isInFlight: function () { return inFlight; },
      // The iOS rescue: after a page freeze the chain may be dead (its promise never
      // settles). Restart only when nothing is in flight, or the in-flight request is
      // older than the poll deadline — never pile a second request on a live one.
      restartIfStuck: function () {
        if (!running) return false;
        if (inFlight && Date.now() - inFlightSince < POLL_TIMEOUT_MS) return false;
        if (opts.onRestart) { try { opts.onRestart(); } catch (e) {} }
        this.start();
        return true;
      },
      currentDelayMs: function () { return delayMs; }
    };
  }

  // ---------- 5. api helpers ----------
  // decorate(params, method) may add credentials/origin; must return the params.
  function createApi(config) {
    var apiUrl = config.apiUrl, origin = config.origin;
    var decorate = config.decorate || function (p) { return p; };
    function query(params) {
      var qs = [];
      for (var k in params) if (Object.prototype.hasOwnProperty.call(params, k) && params[k] !== undefined) qs.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
      qs.push('_t=' + Date.now());
      return qs.join('&');
    }
    return {
      get: function (params, timeoutMs) {
        var p = decorate(params || {}, 'GET');
        if (!p.origin) p.origin = origin;
        return fetchJsonWithTimeout(apiUrl + '?' + query(p), { cache: 'no-store' }, timeoutMs || API_TIMEOUT_MS);
      },
      post: function (payload, timeoutMs) {
        var p = decorate(payload || {}, 'POST');
        if (!p.origin) p.origin = origin;
        return fetchJsonWithTimeout(apiUrl, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify(p) }, timeoutMs || API_TIMEOUT_MS);
      },
      // fire-and-forget; no answer is read
      postNoWait: function (payload) {
        var p = decorate(payload || {}, 'POST');
        if (!p.origin) p.origin = origin;
        try { fetch(apiUrl, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify(p) }).catch(function () {}); } catch (e) {}
      },
      url: function (params) { var p = decorate(params || {}, 'GET'); if (!p.origin) p.origin = origin; return apiUrl + '?' + query(p); }
    };
  }

  // ---------- 7. update check ----------
  // opts: { page: 'examinee.html', versionUrl: 'version.json', intervalMs, onNewVersion(build) }
  // A change must be seen on two consecutive polls; anything that is not a 200 JSON
  // with a hash for this page is ignored. The first good answer is the base.
  function createUpdateCheck(opts) {
    var baseHash = null, pendingHash = null, notified = false, timer = null;
    function probe() {
      // noteHealth=false: version.json is served by GitHub Pages, not by the
      // exam backend, so its failures must not degrade the poll pacing.
      return boundedFetch(opts.versionUrl || 'version.json', { cache: 'no-store' }, 15000, false)
        .then(function (v) { return (v && v.pages && typeof v.pages[opts.page] === 'string') ? { hash: v.pages[opts.page], build: v.build || '' } : null; },
              function () { return null; });
    }
    function tick() {
      if (notified) return;
      probe().then(function (v) {
        if (!v) return;
        if (!baseHash) { baseHash = v.hash; return; }
        if (v.hash === baseHash) { pendingHash = null; return; }
        if (pendingHash !== v.hash) { pendingHash = v.hash; return; }
        notified = true;
        try { opts.onNewVersion(v.build); } catch (e) {}
      });
    }
    return {
      start: function () { probe().then(function (v) { if (v) baseHash = v.hash; }); timer = setInterval(tick, opts.intervalMs || 120000); },
      stop: function () { if (timer) clearInterval(timer); timer = null; },
      // for tests
      _tick: tick
    };
  }

  // ---------- 8. client log ring ----------
  var logRing = [];
  function log(event, data) {
    try {
      var entry = { t: Date.now(), e: String(event) };
      if (data !== undefined) entry.d = typeof data === 'string' ? data.slice(0, 120) : data;
      logRing.push(entry);
      if (logRing.length > LOG_MAX_ENTRIES) logRing.splice(0, logRing.length - LOG_MAX_ENTRIES);
    } catch (e) {}
  }
  // Returns a JSON-safe array bounded to LOG_MAX_BYTES (oldest entries dropped first).
  function drainLog() {
    var out = logRing.slice();
    while (out.length && JSON.stringify(out).length > LOG_MAX_BYTES) out.shift();
    logRing = [];
    return out;
  }

  global.ExamTransport = {
    API_TIMEOUT_MS: API_TIMEOUT_MS,
    POLL_TIMEOUT_MS: POLL_TIMEOUT_MS,
    CRITICAL_POST_TIMEOUT_MS: CRITICAL_POST_TIMEOUT_MS,
    SLOW_ANSWER_MS: SLOW_ANSWER_MS,
    LONGPOLL_WAIT_SEC: LONGPOLL_WAIT_SEC,
    LONGPOLL_GAP_MS: LONGPOLL_GAP_MS,
    LONGPOLL_TIMEOUT_MS: LONGPOLL_TIMEOUT_MS,
    fetchJsonWithTimeout: fetchJsonWithTimeout,
    fetchJsonQuiet: fetchJsonQuiet,
    noteTransport: noteTransport,
    isBackendDegraded: isBackendDegraded,
    degradedSince: degradedSince,
    jitterMs: jitterMs,
    nextPollDelay: nextPollDelay,
    pacePoll: pacePoll,
    createPollLoop: createPollLoop,
    createApi: createApi,
    createUpdateCheck: createUpdateCheck,
    log: log,
    drainLog: drainLog,
    // test hooks
    _resetHealth: function () { transportState = 'ok'; transportDegradedSince = 0; consecutiveTimeouts = 0; },
    _setJitter: function (fn) { jitterMs = fn || jitterMs; global.ExamTransport.jitterMs = jitterMs; }
  };
})(typeof window !== 'undefined' ? window : this);
