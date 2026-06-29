/**
 * Cloudflare Worker — Sincronización Previfuego → Supabase
 *
 * Recibe los datos (locales + extintores) desde la app y los inserta en
 * Supabase usando la SERVICE KEY, que vive aquí como secret del Worker y
 * NUNCA viaja al navegador.
 *
 * ── Despliegue ───────────────────────────────────────────────────────────────
 * 1. Cloudflare → Workers & Pages → Create → Worker → nombre: "previfuego-sync"
 * 2. Pega este código y haz Deploy.
 * 3. En Settings → Variables and Secrets, agrega estos 3 SECRETS (tipo "Secret"):
 *      SUPABASE_URL          = https://ewypvxglljanfkwyeank.supabase.co
 *      SUPABASE_SERVICE_KEY  = (la NUEVA service key que generaste en Supabase)
 *      SYNC_TOKEN            = (una contraseña que inventes, ej. 40 caracteres aleatorios)
 * 4. Copia la URL del Worker (ej. https://previfuego-sync.TU-CUENTA.workers.dev)
 *    y pégala en la app (Configuración → Sincronización), junto con el SYNC_TOKEN.
 */

export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, X-Sync-Token',
    };
    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });
    if (request.method !== 'POST') {
      return json({ error: 'Método no permitido' }, 405, cors);
    }

    // Autenticación: token compartido (de bajo privilegio, rotable sin tocar Supabase)
    if (request.headers.get('X-Sync-Token') !== env.SYNC_TOKEN) {
      return json({ error: 'No autorizado' }, 401, cors);
    }

    let body;
    try { body = await request.json(); } catch { return json({ error: 'JSON inválido' }, 400, cors); }
    const { locales, extintores } = body || {};
    if (!Array.isArray(locales) || !Array.isArray(extintores) || !locales.length) {
      return json({ error: 'Faltan locales o extintores' }, 400, cors);
    }

    const SB = env.SUPABASE_URL;
    const h = {
      'apikey': env.SUPABASE_SERVICE_KEY,
      'Authorization': `Bearer ${env.SUPABASE_SERVICE_KEY}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal',
    };

    // Reemplazo total: borra y reinserta (extintores primero por la FK)
    await fetch(`${SB}/rest/v1/extintores?id=gt.0`, { method: 'DELETE', headers: h });
    await fetch(`${SB}/rest/v1/locales?codigo=not.is.null`, { method: 'DELETE', headers: h });

    const rLoc = await fetch(`${SB}/rest/v1/locales`, { method: 'POST', headers: h, body: JSON.stringify(locales) });
    if (!rLoc.ok) return json({ error: 'Locales: ' + (await rLoc.text()).slice(0, 300) }, 502, cors);

    let extOk = 0;
    for (let i = 0; i < extintores.length; i += 200) {
      const batch = extintores.slice(i, i + 200);
      const rExt = await fetch(`${SB}/rest/v1/extintores`, { method: 'POST', headers: h, body: JSON.stringify(batch) });
      if (!rExt.ok) return json({ error: `Extintores lote ${Math.floor(i / 200) + 1}: ` + (await rExt.text()).slice(0, 300) }, 502, cors);
      extOk += batch.length;
    }

    return json({ ok: true, locales: locales.length, extintores: extOk }, 200, cors);
  },
};

function json(obj, status, cors) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...cors, 'Content-Type': 'application/json' },
  });
}
