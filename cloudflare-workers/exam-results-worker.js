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
const TTL_SECONDS = 86400; // 24 hours

function getCorsHeaders(request) {
  var origin = request.headers.get('Origin') || '';
  var isAllowed = ALLOWED_ORIGINS.some(function(allowed) {
    return origin === allowed || origin.startsWith(allowed + ':') || origin.startsWith(allowed + '/');
  });
  return {
    'Access-Control-Allow-Origin': isAllowed ? origin : ALLOWED_ORIGINS[0],
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400'
  };
}

function generateId() {
  return crypto.randomUUID();
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

        await env.EXAM_RESULTS.put(id, html, {
          expirationTtl: TTL_SECONDS,
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
          '<h1 style="color:#f59e0b;">\u05D4\u05E7\u05D9\u05E9\u05D5\u05E8 \u05E4\u05D2 \u05EA\u05D5\u05E7\u05E3</h1>' +
          '<p style="color:#64748b;">\u05E7\u05D9\u05E9\u05D5\u05E8\u05D9\u05DD \u05EA\u05E7\u05E4\u05D9\u05DD \u05DC-24 \u05E9\u05E2\u05D5\u05EA</p>' +
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
