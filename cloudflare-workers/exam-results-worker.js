/**
 * Cloudflare Worker: exam-results  — ES Modules format (deploy with `wrangler deploy`)
 * מאחסן ומגיש דפי תוצאות מבחן דרך KV
 * Deploy at: steep-night-dd06.bohanyzahal.workers.dev
 *
 * Deployed via wrangler (see wrangler.toml next to this file). The dashboard
 * "Edit code → Deploy" was returning 403 on the legacy services/environments
 * endpoint, so wrangler (modern /workers/scripts/ endpoint) is the deploy path.
 *
 * Bindings (accessed via env):
 *   KV namespace: EXAM_RESULTS
 *   Secret:       UPLOAD_SECRET   (must match RESULT_UPLOAD_SECRET in Apps Script)
 *
 * Endpoints:
 *   POST /     — { html } -> store in KV (PERMANENT, no TTL) -> { status:"ok", link }
 *   GET  /:id  — serve the stored HTML
 *   GET  /     — health check (shows links:"permanent" so a deploy is verifiable)
 */

const ALLOWED_ORIGINS = [
  'https://bohanyzahal-cyber.github.io',
  'http://localhost',
  'http://127.0.0.1'
];

const MAX_HTML_SIZE = 2 * 1024 * 1024; // 2MB
// Report links are PERMANENT — the KV put() below omits expirationTtl, so entries
// persist until manually deleted. (The underlying results also live in the Google
// Sheet, so this is just the rendered HTML snapshot.)

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
      // Auth: require valid HMAC token issued by Apps Script. UPLOAD_SECRET
      // binding must match RESULT_UPLOAD_SECRET ScriptProperty in Apps Script.
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
          '<h1 style="color:#ef4444;">קישור לא תקין</h1>' +
          '<p style="color:#64748b;">הכתובת אינה תקינה</p>' +
          '</body></html>',
          { status: 400, headers: { 'Content-Type': 'text/html; charset=utf-8' } }
        );
      }

      var stored = await env.EXAM_RESULTS.get(id);

      if (!stored) {
        return new Response(
          '<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
          '<body style="font-family:Arial,sans-serif;text-align:center;padding:60px 20px;">' +
          '<h1 style="color:#f59e0b;">הדו"ח לא נמצא</h1>' +
          '<p style="color:#64748b;">ייתכן שהקישור שגוי או שהדו"ח הוסר</p>' +
          '</body></html>',
          { status: 404, headers: { 'Content-Type': 'text/html; charset=utf-8' } }
        );
      }

      return new Response(stored, {
        status: 200,
        headers: {
          'Content-Type': 'text/html; charset=utf-8',
          'Cache-Control': 'public, max-age=3600'
        }
      });
    }

    // --- GET / — health check ---
    // 'links' makes the DEPLOYED behaviour verifiable: a quick
    // `curl https://steep-night-dd06.bohanyzahal.workers.dev/` must show
    // links:"permanent". If it still shows just {"status":"ok","service":"exam-results"}
    // the new version is NOT live and links will still expire.
    if (request.method === 'GET' && url.pathname === '/') {
      return Response.json({ status: 'ok', service: 'exam-results', version: '2026-06-09', links: 'permanent' });
    }

    return new Response('Not Found', { status: 404 });
  }
};
