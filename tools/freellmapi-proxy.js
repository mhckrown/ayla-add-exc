/* Puente local entre el complemento y FreeLLMAPI.
 * FreeLLMAPI solo acepta peticiones de navegador desde los orígenes de su lista
 * DASHBOARD_ORIGINS, y la app de escritorio no permite configurarla. Este puente
 * (sin dependencias) reenvía las peticiones y añade las cabeceras CORS necesarias.
 *
 * Uso:   node tools/freellmapi-proxy.js        (o: npm run freellmapi)
 * En Ayla: endpoint http://localhost:3939/v1, modelo "auto", API key = clave unificada.
 * Variables: FREELLMAPI_URL (por defecto detecta :3001 y :31415), PROXY_PORT (3939). */
const http = require('http');

const PORT = parseInt(process.env.PROXY_PORT || '3939', 10);
const CANDIDATES = process.env.FREELLMAPI_URL
  ? [process.env.FREELLMAPI_URL]
  : ['http://127.0.0.1:3001', 'http://127.0.0.1:31415'];
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || 'https://mhckrown.github.io,https://localhost:3000,http://localhost:3000')
  .split(',').map((s) => s.trim()).filter(Boolean);

let target = null;

async function findTarget() {
  for (const base of CANDIDATES) {
    try {
      const r = await fetch(new URL('/v1/models', base), { signal: AbortSignal.timeout(1500) });
      if (r.status < 500) return base.replace(/\/+$/, '');
    } catch { /* siguiente */ }
  }
  return null;
}

function corsHeaders(origin) {
  const allowed = origin && (ALLOWED_ORIGINS.includes('*') || ALLOWED_ORIGINS.includes(origin));
  return allowed ? {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Headers': 'Authorization, Content-Type, x-api-key, anthropic-version, anthropic-dangerous-direct-browser-access',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Private-Network': 'true',
    'Access-Control-Max-Age': '600',
    Vary: 'Origin',
  } : {};
}

http.createServer(async (req, res) => {
  const cors = corsHeaders(req.headers.origin);
  if (req.method === 'OPTIONS') { res.writeHead(204, cors); return res.end(); }
  if (!req.url.startsWith('/v1/')) { res.writeHead(404, cors); return res.end('Solo se reenvía /v1/*'); }
  if (!target) target = await findTarget();
  if (!target) {
    res.writeHead(502, { ...cors, 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ error: { message: `FreeLLMAPI no responde en ${CANDIDATES.join(' ni ')}. ¿Está abierto?` } }));
  }
  const upstream = new URL(req.url, target);
  const headers = { ...req.headers, host: upstream.host };
  delete headers.origin; delete headers.referer;
  const proxyReq = http.request(upstream, { method: req.method, headers }, (up) => {
    const h = { ...up.headers };
    for (const k of Object.keys(h)) {
      if (/^(access-control-|cross-origin-|content-security-policy|x-frame-options)/i.test(k)) delete h[k];
    }
    Object.assign(h, cors);
    res.writeHead(up.statusCode, h);
    up.pipe(res);
  });
  proxyReq.on('error', (e) => {
    target = null;
    if (!res.headersSent) res.writeHead(502, { ...cors, 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: { message: `Error conectando con FreeLLMAPI: ${e.message}` } }));
  });
  req.on('aborted', () => proxyReq.destroy());
  res.on('close', () => { if (!res.writableEnded) proxyReq.destroy(); });
  req.pipe(proxyReq);
}).listen(PORT, '127.0.0.1', async () => {
  target = await findTarget();
  console.log(`Puente FreeLLMAPI -> http://127.0.0.1:${PORT}/v1  (destino: ${target || 'no encontrado todavía'})`);
  console.log(`Orígenes permitidos: ${ALLOWED_ORIGINS.join(', ')}`);
});
