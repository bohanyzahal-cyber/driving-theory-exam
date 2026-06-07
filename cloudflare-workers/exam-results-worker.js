/**
 * Cloudflare Worker: exam-results
 * מאחסן ומגיש דפי תוצאות מבחן דרך KV
 * Deploy at: steep-night-dd06.bohanyzahal.workers.dev
 *
 * KV Binding: EXAM_RESULTS
 *
 * Endpoints:
 *   POST /       — receives { html: "..." }, stores in KV, returns { status: "ok", link: "..." }
 *   GET  /:id    — serves the stored HTML directly (Content-Type: text/html)
 *   GET  /       — health check
 */

const ALLOWED_ORIGINS = [
  'https://bohanyzahal-cyber.github.io',
  'http://localhost',
  'http://127.0.0.1'
];

const MAX_HTML_SIZE = 2 * 1024 * 1024; // 2MB
// Report links are PERMANENT — the chief examiner can print/save the combined
// report any time, not only within 24h. The KV put() below omits expirationTtl,
// so entries persist until manually deleted. (The underlying results also live
// in the Google Sheet, so this is just the rendered HTML snapshot.)
// To re-introduce an expiry later, set a value here and pass it as
// `expirationTtl` in the put() call.
// const TTL_SECONDS = 86400; // 24 hours (disabled — links no longer expire)

function getCorsHeaders(request) {
  var origin = request.headers.get('Origin') || '';
  var isAllowed = ALLOWED_ORIGINS.some(function(allowed) {
    return origin === allowed || origin.startsWith(allowed + ':') || origin.startsWith(allowed + '/');
  });
  return {
    'Access-Control-Allow-Origin': isAllowed ? origin : ALLOWED_ORIGINS[0],
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, X-Auth-Token',
    'Access-Control-Max-Age': '86400'
  };
}

function generateId() {
  return crypto.randomUUID();
}

// --- HMAC token verification (matches handleGetResultUploadToken in Apps Script) ---
// Token format: "<payloadB64>.<sigB64>" where payload is JSON({"exp": <ms>})
// and signature = HMAC-SHA256(payloadB64, UPLOAD_SECRET), both base64url-encoded.
function base64UrlDecode(s) {
  s = s.replace(/-/g, '+').replace(/_/g, '/');
  var pad = (4 - (s.length % 4)) % 4;
  if (pad) s += '='.repeat(pad);
  var bin = atob(s);
  var bytes = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function verifyUploadToken(tokenStr, secret) {
  if (!tokenStr || typeof tokenStr !== 'string' || tokenStr.indexOf('.') < 0) {
    return { valid: false, reason: 'malformed' };
  }
  var parts = tokenStr.split('.');
  if (parts.length !== 2) return { valid: false, reason: 'malformed' };
  var payloadB64 = parts[0];
  var sigB64 = parts[1];
  try {
    var key = await crypto.subtle.importKey(
      'raw',
      new TextEncoder().encode(secret),
      { name: 'HMAC', hash: 'SHA-256' },
      false,
      ['verify']
    );
    var sigBytes = base64UrlDecode(sigB64);
    var dataBytes = new TextEncoder().encode(payloadB64);
    var valid = await crypto.subtle.verify('HMAC', key, sigBytes, dataBytes);
    if (!valid) return { valid: false, reason: 'bad_signature' };
    var payloadJson = new TextDecoder().decode(base64UrlDecode(payloadB64));
    var payload = JSON.parse(payloadJson);
    if (typeof payload.exp !== 'number' || Date.now() > payload.exp) {
      return { valid: false, reason: 'expired' };
    }
    return { valid: true };
  } catch (e) {
    return { valid: false, reason: 'verify_error' };
  }
}

export default {
  async fetch(request, env) {
    var url = new URL(request.url);
    var cors = getCorsHeaders(request);

    // --- CORS preflight ---
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: cors });
    }

    // --- POST / — store result HTML ---
    if (request.method === 'POST' && url.pathname === '/') {
      // Auth: require valid HMAC token issued by Apps Script.
      // The UPLOAD_SECRET worker binding must match the RESULT_UPLOAD_SECRET
      // ScriptProperty in Apps Script (both sides hold the same secret string).
      if (!env.UPLOAD_SECRET) {
        return Response.json(
          { status: 'error', message: 'UPLOAD_SECRET binding not configured on worker' },
          { status: 500, headers: cors }
        );
      }
      var authToken = request.headers.get('X-Auth-Token') || '';
      var authResult = await verifyUploadToken(authToken, env.UPLOAD_SECRET);
      if (!authResult.valid) {
        return Response.json(
          { status: 'error', message: 'Unauthorized', reason: authResult.reason },
          { status: 401, headers: cors }
        );
      }
      try {
        var body = await request.json();
        var html = body.html;

        if (!html || typeof html !== 'string') {
          return Response.json(
            { status: 'error', message: 'Missing or invalid html field' },
            { status: 400, headers: cors }
          );
        }

        if (html.length > MAX_HTML_SIZE) {
          return Response.json(
            { status: 'error', message: 'HTML too large (max 2MB)' },
            { status: 413, headers: cors }
          );
        }

        var id = generateId();

        // No expirationTtl → the entry never expires (permanent link).
        await env.EXAM_RESULTS.put(id, html, {
          metadata: {
            created: new Date().toISOString(),
            size: html.length
          }
        });

        var link = url.origin + '/' + id;

        return Response.json(
          { status: 'ok', link: link },
          { status: 200, headers: cors }
        );
      } catch (err) {
        return Response.json(
          { status: 'error', message: err.message || 'Server error' },
          { status: 500, headers: cors }
        );
      }
    }

    // --- GET /:id — serve stored HTML ---
    if (request.method === 'GET' && url.pathname.length > 1) {
      var id = url.pathname.slice(1);

      // UUID format validation
      if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) {
        return new Response(
          '<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
          '<body style="font-family:Arial,sans-serif;text-align:center;padding:60px 20px;">' +
          '<h1 style="color:#ef4444;">\u05E7\u05D9\u05E9\u05D5\u05E8 \u05DC\u05D0 \u05EA\u05E7\u05D9\u05DF</h1>' +
          '<p style="color:#64748b;">\u05D4\u05DB\u05EA\u05D5\u05D1\u05EA \u05D0\u05D9\u05E0\u05D4 \u05EA\u05E7\u05D9\u05E0\u05D4</p>' +
          '</body></html>',
          { status: 400, headers: { 'Content-Type': 'text/html; charset=utf-8' } }
        );
      }

      var html = await env.EXAM_RESULTS.get(id);

      if (!html) {
        return new Response(
          '<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
          '<body style="font-family:Arial,sans-serif;text-align:center;padding:60px 20px;">' +
          '<h1 style="color:#f59e0b;">\u05D4\u05D3\u05D5"\u05D7 \u05DC\u05D0 \u05E0\u05DE\u05E6\u05D0</h1>' +
          '<p style="color:#64748b;">\u05D9\u05D9\u05EA\u05DB\u05DF \u05E9\u05D4\u05E7\u05D9\u05E9\u05D5\u05E8 \u05E9\u05D2\u05D5\u05D9 \u05D0\u05D5 \u05E9\u05D4\u05D3\u05D5"\u05D7 \u05D4\u05D5\u05E1\u05E8</p>' +
          '</body></html>',
          { status: 404, headers: { 'Content-Type': 'text/html; charset=utf-8' } }
        );
      }

      return new Response(html, {
        status: 200,
        headers: {
          'Content-Type': 'text/html; charset=utf-8',
          'Cache-Control': 'public, max-age=3600'
        }
      });
    }

    // --- GET / — health check ---
    if (request.method === 'GET' && url.pathname === '/') {
      return Response.json({ status: 'ok', service: 'exam-results' });
    }

    return new Response('Not Found', { status: 404 });
  }
};
