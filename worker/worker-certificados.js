/**
 * Cloudflare Worker — Previfuego (certificados + locales)
 *
 * Guarda y sirve los certificados PDF y la lista de locales usando un token
 * de GitHub que vive AQUÍ como secret del Worker y NUNCA viaja al navegador
 * ni a ningún celular. Así:
 *   - Cualquier celular (técnico o admin), la primera vez que abre la app,
 *     ya ve los 198 locales actualizados — sin token, sin configuración.
 *   - Los técnicos respaldan certificados sin ingresar ninguna clave.
 *   - El admin lista/descarga certificados y edita locales sin tener el token.
 *
 * Solo puede tocar certificados/*.pdf y datos/libreta_locales.json — nada
 * más del repo.
 *
 * ── Despliegue (una sola vez) ────────────────────────────────────────────────
 * 1. Cloudflare → Workers & Pages → Create → Worker → nombre: "previfuego-cert"
 * 2. Pega este código y Deploy.
 * 3. Settings → Variables and Secrets → agrega 1 SECRET:
 *      GH_TOKEN = (un GitHub token con permiso de escritura al repo)
 * 4. Copia la URL del Worker y pásala para embeberla en la app.
 */

const GH_REPO = 'alejosl0801/previfuego-facturacion';
const GH_BRANCH = 'main';
const LOCALES_PATH = 'datos/libreta_locales.json';
const LOCALES_SEMILLA = 'datos/locales_iniciales.json'; // datos base, si aún no existe la libreta

export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };
    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });

    const auth = { Authorization: `token ${env.GH_TOKEN}`, Accept: 'application/vnd.github+json', 'User-Agent': 'previfuego-worker' };
    const url = new URL(request.url);
    const action = url.searchParams.get('action');
    const json = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { ...cors, 'Content-Type': 'application/json' } });

    // Solo se permite operar dentro de certificados/*.pdf
    const pathOk = p => typeof p === 'string' && /^certificados\/[\w.\-\/]+\.pdf$/.test(p) && !p.includes('..');

    const ghGet = (path) => fetch(`https://api.github.com/repos/${GH_REPO}/contents/${path}?ref=${GH_BRANCH}`, { headers: auth });
    const ghPut = (path, contentB64, message, sha) => fetch(`https://api.github.com/repos/${GH_REPO}/contents/${path}`, {
      method: 'PUT', headers: { ...auth, 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, content: contentB64, branch: GH_BRANCH, ...(sha ? { sha } : {}) }),
    });

    try {
      // ── LOCALES: leer la lista (GET ?action=locales) ──────────────────────
      // Sin token, sin login: cualquier celular la llama al abrir la app.
      if (action === 'locales' && request.method === 'GET') {
        let r = await ghGet(LOCALES_PATH);
        if (!r.ok) r = await ghGet(LOCALES_SEMILLA); // primera vez: aún no hay libreta guardada
        if (!r.ok) return json({ locales: [] });
        const d = await r.json();
        const texto = decodeURIComponent(escape(atob(d.content.replace(/\n/g, ''))));
        return json({ locales: JSON.parse(texto) });
      }

      // ── LOCALES: guardar la lista completa (POST { action:'locales', locales:[...] }) ──
      if (action === 'locales' && request.method === 'POST') {
        const body = await request.json();
        if (!Array.isArray(body.locales)) return json({ error: 'Formato inválido' }, 400);
        const contenido = btoa(unescape(encodeURIComponent(JSON.stringify(body.locales, null, 2))));
        let sha = null;
        const chk = await ghGet(LOCALES_PATH);
        if (chk.ok) sha = (await chk.json()).sha;
        const put = await ghPut(LOCALES_PATH, contenido, 'libreta: actualización ' + new Date().toISOString(), sha);
        if (!put.ok) return json({ error: 'GitHub: ' + (await put.text()).slice(0, 200) }, 502);
        return json({ ok: true, total: body.locales.length });
      }

      // ── CERTIFICADOS: subir uno (POST { path, content }) ──────────────────
      if (request.method === 'POST' && !action) {
        const body = await request.json();
        if (!pathOk(body.path)) return json({ error: 'Ruta no permitida' }, 400);
        if (!body.content) return json({ error: 'Falta el contenido' }, 400);
        let sha = null;
        const chk = await ghGet(body.path);
        if (chk.ok) sha = (await chk.json()).sha;
        const put = await ghPut(body.path, body.content, `cert: ${body.path}`, sha);
        if (!put.ok) return json({ error: 'GitHub: ' + (await put.text()).slice(0, 200) }, 502);
        return json({ ok: true });
      }

      // ── CERTIFICADOS: listar carpeta (GET ?action=list&dir=certificados/MES-ANIO) ──
      if (action === 'list') {
        const dir = url.searchParams.get('dir') || '';
        if (!/^certificados\/[\w.\-]+$/.test(dir)) return json({ error: 'Carpeta no permitida' }, 400);
        const r = await ghGet(dir);
        if (!r.ok) return json({ files: [] });
        const files = await r.json();
        return json({ files: (Array.isArray(files) ? files : []).filter(f => f.name.endsWith('.pdf')).map(f => ({ name: f.name, path: f.path })) });
      }

      // ── CERTIFICADOS: descargar uno (GET ?action=get&path=...) → base64 ───
      if (action === 'get') {
        const p = url.searchParams.get('path') || '';
        if (!pathOk(p)) return json({ error: 'Ruta no permitida' }, 400);
        const r = await ghGet(p);
        if (!r.ok) return json({ error: 'No encontrado' }, 404);
        const d = await r.json();
        return json({ content: d.content });
      }

      return json({ error: 'Acción no reconocida' }, 400);
    } catch (e) {
      return json({ error: String(e && e.message || e) }, 500);
    }
  },
};
