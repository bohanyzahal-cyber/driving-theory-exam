/**
 * Cloudflare Worker — Driving Theory Image Proxy
 *
 * Proxies and caches gov.il images to bypass:
 *  - Mobile carrier image compression (HOT/Cellcom/Pelephone)
 *  - Chrome Data Saver / Lite Mode
 *  - Samsung Internet / Brave third-party blocking
 *  - In-app WebView TLS issues (WhatsApp/Telegram preview)
 *  - gov.il referer/hotlink restrictions
 *  - Slow DNS to gov.il from certain ISPs
 *
 * Usage from the page:
 *   <img src="https://your-worker.workers.dev/img/BlobFolder/generalpage/tq_pic_01/he/TQ_PIC_31276.jpg">
 *
 * Caching:
 *  - Edge cache (Cloudflare): 30 days for successful images
 *  - Browser cache: 7 days
 *  - 404 / errors are also cached briefly so we don't hammer gov.il
 *
 * Free tier impact:
 *  - 100k requests/day free
 *  - At 300 examinees × 20 images = 6,000 reqs/day → ~6% of free tier
 *  - First-time fetch: gov.il is hit; all subsequent requests serve from edge cache (free)
 */

const ORIGIN = 'https://www.gov.il';
const ALLOWED_PATH_PREFIX = '/BlobFolder/'; // only allow gov.il blob storage paths
const SUCCESS_CACHE_SECONDS = 60 * 60 * 24 * 30; // 30 days
const ERROR_CACHE_SECONDS = 60 * 5;              // 5 minutes
const ALLOWED_EXTENSIONS = /\.(jpg|jpeg|png|gif|webp|svg)$/i;

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // ---- CORS preflight ----
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET, OPTIONS',
          'Access-Control-Max-Age': '86400'
        }
      });
    }

    if (request.method !== 'GET' && request.method !== 'HEAD') {
      return new Response('Method Not Allowed', { status: 405 });
    }

    // ---- Health check ----
    if (url.pathname === '/' || url.pathname === '/health') {
      return new Response('Image Proxy OK\nUsage: /img/BlobFolder/...\n', {
        headers: { 'Content-Type': 'text/plain; charset=utf-8' }
      });
    }

    // ---- Validate path ----
    // We expect: /img/BlobFolder/generalpage/tq_pic_XX/he/TQ_PIC_XXXX.jpg
    // Strip leading "/img" prefix.
    let originPath = url.pathname;
    if (originPath.startsWith('/img/')) {
      originPath = originPath.slice('/img'.length);
    }

    if (!originPath.startsWith(ALLOWED_PATH_PREFIX)) {
      return jsonError(400, 'Invalid path. Expected /img' + ALLOWED_PATH_PREFIX + '...');
    }
    if (!ALLOWED_EXTENSIONS.test(originPath)) {
      return jsonError(400, 'Invalid file extension');
    }
    // Block path traversal
    if (originPath.includes('..') || originPath.includes('//')) {
      return jsonError(400, 'Invalid path characters');
    }

    const originUrl = ORIGIN + originPath;
    const cacheKey = new Request(url.toString(), { method: 'GET' });
    const cache = caches.default;

    // ---- Try edge cache first ----
    let response = await cache.match(cacheKey);
    if (response) {
      // Add a header so we can debug cache hits in DevTools Network tab
      response = new Response(response.body, response);
      response.headers.set('X-Proxy-Cache', 'HIT');
      return response;
    }

    // ---- Fetch from gov.il (without sending Referer) ----
    let originResp;
    try {
      originResp = await fetch(originUrl, {
        method: 'GET',
        redirect: 'follow',
        cf: {
          // Tell Cloudflare to also cache at network layer (free, automatic)
          cacheTtl: SUCCESS_CACHE_SECONDS,
          cacheEverything: true
        },
        headers: {
          // Hide our origin from gov.il — they sometimes block requests with a referer
          'User-Agent': 'Mozilla/5.0 (compatible; DrivingTheoryProxy/1.0)',
          'Accept': 'image/jpeg,image/png,image/webp,image/*,*/*;q=0.8',
          'Accept-Language': 'he-IL,he;q=0.9,en;q=0.8'
        }
      });
    } catch (e) {
      return jsonError(502, 'Origin fetch failed: ' + (e.message || 'unknown'));
    }

    // ---- Build response with our own caching headers ----
    const isSuccess = originResp.ok;
    const cacheSeconds = isSuccess ? SUCCESS_CACHE_SECONDS : ERROR_CACHE_SECONDS;
    const browserCacheSeconds = isSuccess ? 60 * 60 * 24 * 7 : 60; // browser: 7 days success / 1 min error

    // Stream the body (don't buffer — saves CPU/memory)
    const newHeaders = new Headers();
    const contentType = originResp.headers.get('Content-Type') || 'image/jpeg';
    newHeaders.set('Content-Type', contentType);
    newHeaders.set('Cache-Control', `public, max-age=${browserCacheSeconds}, s-maxage=${cacheSeconds}, immutable`);
    newHeaders.set('Access-Control-Allow-Origin', '*');
    newHeaders.set('X-Proxy-Cache', 'MISS');
    newHeaders.set('X-Proxy-Origin-Status', String(originResp.status));

    const proxied = new Response(originResp.body, {
      status: originResp.status,
      statusText: originResp.statusText,
      headers: newHeaders
    });

    // ---- Save to edge cache (background, doesn't delay response) ----
    if (isSuccess) {
      ctx.waitUntil(cache.put(cacheKey, proxied.clone()));
    }

    return proxied;
  }
};

function jsonError(status, message) {
  return new Response(JSON.stringify({ error: message, status }), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'public, max-age=60'
    }
  });
}
