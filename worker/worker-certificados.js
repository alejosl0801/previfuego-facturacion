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
 * Los certificados con fotos suelen pesar varios MB, y la API de "Contents"
 * de GitHub (PUT /contents/{path}) tiene un límite práctico de ~1MB antes de
 * empezar a fallar. Por eso subirCertGitHub usa la Git Data API (blob + tree
 * + commit), que soporta archivos de hasta 100MB.
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

    // Sube un archivo de cualquier tamaño razonable (hasta 100MB) vía la Git
    // Data API: blob -> tree -> commit -> mover la rama al nuevo commit.
    // Necesario para PDFs con fotos, que fácilmente pasan de 1MB.
    async function ghPutBlob(path, base64Content, message) {
      const paso = async (nombre, resp) => {
        if (!resp.ok) throw new Error(nombre + ': HTTP ' + resp.status + ' ' + (await resp.text()).slice(0, 300));
        return resp.json();
      };
      const refData = await paso('ref', await fetch(
        `https://api.github.com/repos/${GH_REPO}/git/refs/heads/${GH_BRANCH}`, { headers: auth }));
      const commitSha = refData.object.sha;

      const commitData = await paso('commit', await fetch(
        `https://api.github.com/repos/${GH_REPO}/git/commits/${commitSha}`, { headers: auth }));
      const baseTreeSha = commitData.tree.sha;

      const blobData = await paso('blob', await fetch(`https://api.github.com/repos/${GH_REPO}/git/blobs`, {
        method: 'POST', headers: { ...auth, 'Content-Type': 'application/json' },
        body: JSON.stringify({ content: base64Content, encoding: 'base64' }),
      }));

      const treeData = await paso('tree', await fetch(`https://api.github.com/repos/${GH_REPO}/git/trees`, {
        method: 'POST', headers: { ...auth, 'Content-Type': 'application/json' },
        body: JSON.stringify({ base_tree: baseTreeSha, tree: [{ path, mode: '100644', type: 'blob', sha: blobData.sha }] }),
      }));

      const newCommitData = await paso('newCommit', await fetch(`https://api.github.com/repos/${GH_REPO}/git/commits`, {
        method: 'POST', headers: { ...auth, 'Content-Type': 'application/json' },
        body: JSON.stringify({ message, tree: treeData.sha, parents: [commitSha] }),
      }));

      await paso('updateRef', await fetch(`https://api.github.com/repos/${GH_REPO}/git/refs/heads/${GH_BRANCH}`, {
        method: 'PATCH', headers: { ...auth, 'Content-Type': 'application/json' },
        body: JSON.stringify({ sha: newCommitData.sha }),
      }));
    }

    // Lee un archivo devolviendo su base64, usando la Git Data API si la
    // Contents API no incluyó el contenido (pasa con archivos > ~1MB).
    async function ghGetContent(path) {
      const r = await ghGet(path);
      if (!r.ok) return null;
      const d = await r.json();
      if (d.content) return d.content;
      if (!d.sha) return null;
      const blobR = await fetch(`https://api.github.com/repos/${GH_REPO}/git/blobs/${d.sha}`, { headers: auth });
      if (!blobR.ok) return null;
      return (await blobR.json()).content;
    }

    try {
      // ── DIAGNÓSTICO: probar que el token puede escribir (GET ?action=selftest) ──
      // Solo para depurar: escribe/borra un archivo de prueba (con el mismo
      // método que usan los certificados reales) y devuelve el error real de
      // GitHub si algo falla.
      if (action === 'selftest') {
        if (!env.GH_TOKEN) return json({ ok: false, paso: 'secret', error: 'GH_TOKEN no está configurado en el Worker' });
        const testPath = 'certificados/_selftest.txt';
        try {
          await ghPutBlob(testPath, btoa('selftest ' + new Date().toISOString()), 'selftest: prueba de escritura');
        } catch (e) {
          return json({ ok: false, paso: 'escritura', error: String(e && e.message || e) });
        }
        // Limpieza: borrar el archivo de prueba
        const chk = await ghGet(testPath);
        if (chk.ok) {
          const sha = (await chk.json()).sha;
          await fetch(`https://api.github.com/repos/${GH_REPO}/contents/${testPath}`, {
            method: 'DELETE', headers: { ...auth, 'Content-Type': 'application/json' },
            body: JSON.stringify({ message: 'selftest: limpieza', sha, branch: GH_BRANCH }),
          });
        }
        return json({ ok: true, mensaje: 'El Worker puede leer y escribir en el repo correctamente.' });
      }

      // ── LOCALES: leer la lista (GET ?action=locales) ──────────────────────
      // Sin token, sin login: cualquier celular la llama al abrir la app.
      if (action === 'locales' && request.method === 'GET') {
        let content = await ghGetContent(LOCALES_PATH);
        if (!content) content = await ghGetContent(LOCALES_SEMILLA); // primera vez: aún no hay libreta guardada
        if (!content) return json({ locales: [] });
        const texto = decodeURIComponent(escape(atob(content.replace(/\n/g, ''))));
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
        try {
          await ghPutBlob(body.path, body.content, `cert: ${body.path}`);
        } catch (e) {
          return json({ error: 'GitHub: ' + String(e && e.message || e) }, 502);
        }
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
        const content = await ghGetContent(p);
        if (!content) return json({ error: 'No encontrado' }, 404);
        return json({ content });
      }

      return json({ error: 'Acción no reconocida' }, 400);
    } catch (e) {
      return json({ error: String(e && e.message || e) }, 500);
    }
  },
};
