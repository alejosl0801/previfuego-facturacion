/**
 * Cloudflare Worker — Certificados Previfuego
 *
 * Guarda y sirve los PDF de certificados en el repo de GitHub, usando un
 * token que vive AQUÍ como secret del Worker y NUNCA viaja al navegador.
 * Así los técnicos respaldan certificados sin ingresar ninguna clave, y el
 * admin los lista/descarga desde la app sin tener el token.
 *
 * Solo puede tocar la carpeta certificados/*.pdf — no el resto del repo.
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

export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };
    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });

    const auth = { Authorization: `token ${env.GH_TOKEN}`, Accept: 'application/vnd.github+json', 'User-Agent': 'previfuego-cert' };
    const url = new URL(request.url);
    const json = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { ...cors, 'Content-Type': 'application/json' } });

    // Solo se permite operar dentro de certificados/*.pdf
    const pathOk = p => typeof p === 'string' && /^certificados\/[\w.\-\/]+\.pdf$/.test(p) && !p.includes('..');

    try {
      // ── Subir un certificado (POST { path, content }) ──
      if (request.method === 'POST') {
        const body = await request.json();
        if (!pathOk(body.path)) return json({ error: 'Ruta no permitida' }, 400);
        if (!body.content) return json({ error: 'Falta el contenido' }, 400);
        // ¿Existe ya? (para actualizar con sha)
        let sha = null;
        const chk = await fetch(`https://api.github.com/repos/${GH_REPO}/contents/${body.path}?ref=${GH_BRANCH}`, { headers: auth });
        if (chk.ok) sha = (await chk.json()).sha;
        const put = await fetch(`https://api.github.com/repos/${GH_REPO}/contents/${body.path}`, {
          method: 'PUT', headers: { ...auth, 'Content-Type': 'application/json' },
          body: JSON.stringify({ message: `cert: ${body.path}`, content: body.content, branch: GH_BRANCH, ...(sha ? { sha } : {}) }),
        });
        if (!put.ok) return json({ error: 'GitHub: ' + (await put.text()).slice(0, 200) }, 502);
        return json({ ok: true });
      }

      // ── Listar certificados de una carpeta (GET ?action=list&dir=certificados/MES-ANIO) ──
      if (url.searchParams.get('action') === 'list') {
        const dir = url.searchParams.get('dir') || '';
        if (!/^certificados\/[\w.\-]+$/.test(dir)) return json({ error: 'Carpeta no permitida' }, 400);
        const r = await fetch(`https://api.github.com/repos/${GH_REPO}/contents/${dir}?ref=${GH_BRANCH}`, { headers: auth });
        if (!r.ok) return json({ files: [] });
        const files = await r.json();
        return json({ files: (Array.isArray(files) ? files : []).filter(f => f.name.endsWith('.pdf')).map(f => ({ name: f.name, path: f.path })) });
      }

      // ── Descargar un certificado (GET ?action=get&path=...) → base64 ──
      if (url.searchParams.get('action') === 'get') {
        const p = url.searchParams.get('path') || '';
        if (!pathOk(p)) return json({ error: 'Ruta no permitida' }, 400);
        const r = await fetch(`https://api.github.com/repos/${GH_REPO}/contents/${p}?ref=${GH_BRANCH}`, { headers: auth });
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
