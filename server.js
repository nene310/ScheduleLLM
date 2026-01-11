const http = require('http');
const crypto = require('crypto');
const { URL } = require('url');

const env = (k, d = undefined) => (process.env[k] === undefined ? d : process.env[k]);

const PORT = Number(env('PORT', '3000'));
const REQUIRE_AUTH = String(env('REQUIRE_AUTH', 'true')).toLowerCase() !== 'false';

const LLM_BASE_URL = String(env('LLM_BASE_URL', 'https://dashscope.aliyuncs.com/compatible-mode/v1')).replace(/\/$/, '');
const LLM_API_KEY = String(env('LLM_API_KEY', ''));
const DEFAULT_MODEL = String(env('LLM_MODEL', 'qwen-flash'));

const AUTH_USER = String(env('AUTH_USER', 'admin'));
const AUTH_PASS = String(env('AUTH_PASS', ''));
const JWT_SECRET = String(env('JWT_SECRET', ''));
const JWT_TTL_SECONDS = Number(env('JWT_TTL_SECONDS', '43200'));

const RATE_LIMIT_RPM = Number(env('RATE_LIMIT_RPM', '120'));
const MAX_BODY_BYTES = Number(env('MAX_BODY_BYTES', '1048576'));
const ALLOWED_ORIGINS = String(env('ALLOWED_ORIGINS', '')).split(',').map(s => s.trim()).filter(Boolean);

const COOKIE_NAME = String(env('AUTH_COOKIE_NAME', 'schedulellm_token'));
const COOKIE_SECURE = String(env('COOKIE_SECURE', 'true')).toLowerCase() !== 'false';

function nowMs() { return Date.now(); }

function base64urlEncode(buf) {
  return Buffer.from(buf).toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
}
function base64urlDecode(str) {
  const s = String(str).replace(/-/g, '+').replace(/_/g, '/');
  const pad = s.length % 4 === 0 ? '' : '='.repeat(4 - (s.length % 4));
  return Buffer.from(s + pad, 'base64');
}

function timingSafeEqualStr(a, b) {
  const aa = Buffer.from(String(a));
  const bb = Buffer.from(String(b));
  if (aa.length !== bb.length) return false;
  return crypto.timingSafeEqual(aa, bb);
}

function jwtSignHS256(payload, secret) {
  const header = { alg: 'HS256', typ: 'JWT' };
  const h = base64urlEncode(JSON.stringify(header));
  const p = base64urlEncode(JSON.stringify(payload));
  const data = `${h}.${p}`;
  const sig = crypto.createHmac('sha256', secret).update(data).digest();
  return `${data}.${base64urlEncode(sig)}`;
}

function jwtVerifyHS256(token, secret) {
  const parts = String(token || '').split('.');
  if (parts.length !== 3) return { ok: false, err: 'bad_token' };
  const [h, p, s] = parts;
  const data = `${h}.${p}`;
  const expected = crypto.createHmac('sha256', secret).update(data).digest();
  const got = base64urlDecode(s);
  if (expected.length !== got.length || !crypto.timingSafeEqual(expected, got)) return { ok: false, err: 'bad_sig' };
  let payload;
  try {
    payload = JSON.parse(base64urlDecode(p).toString('utf8'));
  } catch {
    return { ok: false, err: 'bad_payload' };
  }
  const exp = Number(payload.exp || 0) * 1000;
  if (exp && nowMs() > exp) return { ok: false, err: 'expired' };
  return { ok: true, payload };
}

function parseCookies(req) {
  const header = req.headers.cookie || '';
  const out = {};
  header.split(';').forEach(part => {
    const idx = part.indexOf('=');
    if (idx < 0) return;
    const k = part.slice(0, idx).trim();
    const v = part.slice(idx + 1).trim();
    if (k) out[k] = decodeURIComponent(v);
  });
  return out;
}

function sendJson(res, status, obj, headers = {}) {
  const body = Buffer.from(JSON.stringify(obj));
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': String(body.length),
    ...headers
  });
  res.end(body);
}

function sendText(res, status, text, headers = {}) {
  const body = Buffer.from(String(text || ''), 'utf8');
  res.writeHead(status, {
    'Content-Type': 'text/plain; charset=utf-8',
    'Content-Length': String(body.length),
    ...headers
  });
  res.end(body);
}

function readJson(req, maxBytes) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > maxBytes) {
        reject(Object.assign(new Error('body_too_large'), { code: 'body_too_large' }));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8');
      if (!raw) return resolve({ raw: '', json: null });
      try {
        return resolve({ raw, json: JSON.parse(raw) });
      } catch (e) {
        reject(Object.assign(new Error('bad_json'), { code: 'bad_json', detail: e.message }));
      }
    });
    req.on('error', reject);
  });
}

const rate = new Map();
function rateLimitKey(req) {
  const xf = String(req.headers['x-forwarded-for'] || '');
  const ip = xf.split(',')[0].trim() || req.socket.remoteAddress || 'unknown';
  return ip;
}
function rateLimitAllow(key, limitPerMin) {
  const t = nowMs();
  const w = 60_000;
  const rec = rate.get(key) || { win: Math.floor(t / w), count: 0 };
  const win = Math.floor(t / w);
  if (rec.win !== win) {
    rec.win = win;
    rec.count = 0;
  }
  rec.count += 1;
  rate.set(key, rec);
  return rec.count <= limitPerMin;
}

const nonceStore = new Map();
function rememberNonce(subject, nonce, ttlMs = 5 * 60_000) {
  const t = nowMs();
  const exp = t + ttlMs;
  const key = `${subject}::${nonce}`;
  nonceStore.set(key, exp);
}
function seenNonce(subject, nonce) {
  const t = nowMs();
  const key = `${subject}::${nonce}`;
  const exp = nonceStore.get(key);
  if (!exp) return false;
  if (t > exp) {
    nonceStore.delete(key);
    return false;
  }
  return true;
}
setInterval(() => {
  const t = nowMs();
  for (const [k, exp] of nonceStore.entries()) {
    if (t > exp) nonceStore.delete(k);
  }
  for (const [k, v] of rate.entries()) {
    if (!v || typeof v.win !== 'number') rate.delete(k);
  }
}, 30_000).unref();

function corsHeaders(req) {
  const origin = String(req.headers.origin || '');
  if (!origin) return {};
  if (ALLOWED_ORIGINS.length === 0) return {};
  if (!ALLOWED_ORIGINS.includes(origin)) return {};
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Credentials': 'true',
    'Vary': 'Origin'
  };
}

function setAuthCookie(res, token) {
  const parts = [
    `${COOKIE_NAME}=${encodeURIComponent(token)}`,
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    `Max-Age=${JWT_TTL_SECONDS}`
  ];
  if (COOKIE_SECURE) parts.push('Secure');
  res.setHeader('Set-Cookie', parts.join('; '));
}

function clearAuthCookie(res) {
  const parts = [
    `${COOKIE_NAME}=`,
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    'Max-Age=0'
  ];
  if (COOKIE_SECURE) parts.push('Secure');
  res.setHeader('Set-Cookie', parts.join('; '));
}

function getAuthSubject(req) {
  const authz = String(req.headers.authorization || '');
  if (authz.startsWith('Bearer ')) {
    const token = authz.slice('Bearer '.length).trim();
    if (!JWT_SECRET) return { ok: false, err: 'no_jwt_secret' };
    const v = jwtVerifyHS256(token, JWT_SECRET);
    if (!v.ok) return { ok: false, err: v.err };
    return { ok: true, sub: String(v.payload.sub || ''), payload: v.payload };
  }

  const cookies = parseCookies(req);
  const token = cookies[COOKIE_NAME];
  if (!token) return { ok: false, err: 'no_token' };
  if (!JWT_SECRET) return { ok: false, err: 'no_jwt_secret' };
  const v = jwtVerifyHS256(token, JWT_SECRET);
  if (!v.ok) return { ok: false, err: v.err };
  return { ok: true, sub: String(v.payload.sub || ''), payload: v.payload };
}

async function forwardToLLM(body) {
  if (!LLM_API_KEY) {
    return { ok: false, status: 500, data: { error: 'server_missing_llm_api_key' } };
  }

  const model = String(body?.model || DEFAULT_MODEL);
  const messages = Array.isArray(body?.messages) ? body.messages : null;
  const temperature = body?.temperature === undefined ? 0.1 : Number(body.temperature);

  if (!messages || messages.length === 0) {
    return { ok: false, status: 400, data: { error: 'missing_messages' } };
  }

  const upstream = `${LLM_BASE_URL}/chat/completions`;
  const resp = await fetch(upstream, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${LLM_API_KEY}`
    },
    body: JSON.stringify({ model, messages, temperature })
  });

  const text = await resp.text();
  let json = null;
  try { json = text ? JSON.parse(text) : null; } catch { json = null; }

  if (!resp.ok) {
    return { ok: false, status: resp.status, data: json || { error: 'upstream_error', raw: text.slice(0, 2000) } };
  }
  return { ok: true, status: 200, data: json };
}

const server = http.createServer(async (req, res) => {
  const t0 = nowMs();
  const rid = String(req.headers['x-request-id'] || crypto.randomUUID());
  res.setHeader('X-Request-Id', rid);

  const c = corsHeaders(req);
  for (const [k, v] of Object.entries(c)) res.setHeader(k, v);
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type,Authorization,X-Request-Id,X-Timestamp,X-Nonce');
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url || '/', `http://${req.headers.host || 'localhost'}`);
  const path = url.pathname;

  const key = rateLimitKey(req);
  if (!rateLimitAllow(key, RATE_LIMIT_RPM)) {
    sendJson(res, 429, { error: 'rate_limited', requestId: rid });
    return;
  }

  try {
    if (req.method === 'GET' && path === '/healthz') {
      sendJson(res, 200, { ok: true });
      return;
    }

    if (req.method === 'POST' && path === '/api/auth/logout') {
      clearAuthCookie(res);
      sendJson(res, 200, { ok: true });
      return;
    }

    if (req.method === 'POST' && path === '/api/auth/login') {
      const { json } = await readJson(req, MAX_BODY_BYTES);
      const user = String(json?.username || '');
      const pass = String(json?.password || '');

      if (!JWT_SECRET || !AUTH_PASS) {
        sendJson(res, 500, { error: 'server_auth_not_configured', requestId: rid });
        return;
      }
      if (user !== AUTH_USER || !timingSafeEqualStr(pass, AUTH_PASS)) {
        sendJson(res, 401, { error: 'invalid_credentials', requestId: rid });
        return;
      }

      const iat = Math.floor(nowMs() / 1000);
      const exp = iat + JWT_TTL_SECONDS;
      const token = jwtSignHS256({ sub: user, iat, exp, jti: crypto.randomUUID() }, JWT_SECRET);
      setAuthCookie(res, token);
      sendJson(res, 200, { ok: true, exp, requestId: rid });
      return;
    }

    if (req.method === 'POST' && path === '/api/llm') {
      if (REQUIRE_AUTH) {
        const auth = getAuthSubject(req);
        if (!auth.ok || !auth.sub) {
          sendJson(res, 401, { error: 'unauthorized', reason: auth.err, requestId: rid });
          return;
        }

        const ts = Number(req.headers['x-timestamp'] || 0);
        const nonce = String(req.headers['x-nonce'] || '');
        if (!ts || !nonce) {
          sendJson(res, 400, { error: 'missing_timestamp_or_nonce', requestId: rid });
          return;
        }
        const skew = Math.abs(nowMs() - ts);
        if (skew > 2 * 60_000) {
          sendJson(res, 400, { error: 'timestamp_skew_too_large', requestId: rid });
          return;
        }
        if (seenNonce(auth.sub, nonce)) {
          sendJson(res, 409, { error: 'replay_detected', requestId: rid });
          return;
        }
        rememberNonce(auth.sub, nonce);
      }

      const { json } = await readJson(req, MAX_BODY_BYTES);
      const out = await forwardToLLM(json);
      sendJson(res, out.status, out.ok ? out.data : { ...out.data, requestId: rid });
      return;
    }

    sendJson(res, 404, { error: 'not_found', requestId: rid });
  } catch (e) {
    sendJson(res, 500, { error: 'server_error', requestId: rid });
  } finally {
    const ms = nowMs() - t0;
    const line = JSON.stringify({
      t: new Date().toISOString(),
      rid,
      ip: rateLimitKey(req),
      method: req.method,
      path,
      status: res.statusCode,
      ms
    });
    process.stdout.write(line + '\n');
  }
});

server.listen(PORT, '0.0.0.0', () => {
  process.stdout.write(JSON.stringify({
    t: new Date().toISOString(),
    msg: 'server_listening',
    port: PORT,
    llmBaseUrl: LLM_BASE_URL,
    requireAuth: REQUIRE_AUTH
  }) + '\n');
});