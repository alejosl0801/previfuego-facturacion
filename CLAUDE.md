# Previfuego — Certificados de Extintores

Guía de contexto para cualquier sesión de Claude que trabaje en este repositorio.
Léela completa antes de tocar código o de asumir cómo funciona algo — este proyecto
tiene varias piezas no obvias (un Worker externo, una carpeta de Dropbox aparte,
convenciones de nombres) que son fáciles de romper si se asume mal.

## Qué es esto

Previfuego es una empresa ecuatoriana de mantenimiento y recarga de extintores.
Este repo es una app web de una sola página (`index.html`, sin build ni framework)
que sus técnicos de campo y su administrador usan para emitir certificados de
mantenimiento en PDF, directamente desde el celular.

- **Sitio en vivo**: https://alejosl0801.github.io/previfuego-facturacion/
  (GitHub Pages, se despliega solo al hacer push a `main` vía
  `.github/workflows/pages.yml` / `JamesIves/github-pages-deploy-action`).
- **Repo**: `alejosl0801/previfuego-facturacion` (público).
- El dueño (usuario) es **no técnico** — prefiere instrucciones exactas y accionables
  a explicaciones largas. Responde en español, directo y a veces en mayúsculas
  cuando está frustrado; eso no es agresión, es su estilo normal de escritura rápida
  desde el celular.

## Arquitectura

```
index.html               ← toda la app: HTML + CSS + JS en un solo archivo
worker/worker-certificados.js  ← código del Cloudflare Worker (se copia/pega a mano
                                  en el dashboard de Cloudflare — este repo NO lo
                                  despliega automáticamente)
datos/locales_iniciales.json   ← semilla inicial de la libreta de locales (público)
scripts/sync_certificados_local.py  ← script para bajar certificados a una compu local
certificados/             ← EXISTE en el repo (main), pero está excluido del deploy
                             de GitHub Pages. Aquí caen los PDF que sube el Worker.
```

### Perfiles: Admin vs Técnico

La app tiene dos perfiles (selector al abrir, o botón arriba a la izquierda):

- **Técnico de campo**: solo ve "Emitir Certificado". Adjunta fotos y genera
  certificados. Nunca ve el PDF que genera, ni lo descarga, ni tiene acceso a
  Configuración — todo eso es a propósito (cero-configuración para técnicos,
  y el admin es quien revisa/audita, no quien genera).
- **Admin**: ve "Emitir Certificado" (solo para revisar/descargar lo ya generado,
  nunca genera), "Locales" (editar la libreta) y "Configuración".

### El flujo de un certificado, paso a paso

1. Técnico abre la app → elige un local pendiente en "Emitir Certificado".
2. Adjunta fotos **directamente desde la tarjeta del local en la lista** (botones
   "📷 Antes" / "📷 Cilindro" (solo si el local tiene un CO2 ≥50 LBS, ver
   `tieneCilindroGrande()`) / "📷 Después") — esto es deliberado: el mantenimiento
   real toma 20-40 minutos entre fotos, así que NO es un asistente/wizard dentro
   de un modal; cada foto se guarda al toque, sin abrir nada. Las fotos se
   comprimen a ~350KB c/u (`certComprimirImagen`) y se guardan en un borrador en
   `localStorage` (`certDraft_<codigo>`, expira a los 14 días).
3. Cuando las fotos ya están, toca "📜 Certificado" → un modal simple (tipo de
   novedades/accesorios, nada de fotos ahí) → "Generar Certificado".
4. `generarCertLocal()` arma el PDF con jsPDF **en el propio celular** (no hay
   backend que genere el PDF).
5. El PDF se sube a GitHub **a través de un Cloudflare Worker**
   (`https://previfuego-cert.alejosl0801.workers.dev`), NUNCA directo a la API de
   GitHub con un token del cliente — los técnicos no tienen (ni deben tener)
   ningún token. El Worker guarda el `GH_TOKEN` como secret del lado del servidor.
   Ruta en el repo: `certificados/{MES}-{AÑO}/CERT-{CODIGO}-{AÑO}.pdf` (mes
   calendario en que se generó, en mayúsculas sin tildes: ENERO...DICIEMBRE).
6. Solo se marca el local como "✅ visitado" **después de confirmar** que la
   subida al Worker fue exitosa (con reintentos y backoff). Si falla (sin señal,
   Worker caído), el certificado queda en una cola local
   (`certPendientesSubir`) y se reintenta solo — nunca se le dice al técnico
   "listo" antes de que el PDF esté de verdad respaldado.
7. El correo (EmailJS) se envía también solo después de confirmar la subida.
8. El admin, en su propio dispositivo, sincroniza qué está "visitado" consultando
   el Worker (`sincronizarVisitadosDesdeGitHub()`), porque el estado de
   "visitado" vive en `localStorage` (por dispositivo) y por sí solo no se entera
   de lo que hizo el técnico en su celular.

### El Cloudflare Worker (pieza crítica, no está en Cloudflare vía este repo)

- Código fuente versionado en `worker/worker-certificados.js`, pero el despliegue
  real es **manual**: el usuario copia/pega el contenido en el dashboard de
  Cloudflare (Workers & Pages → `previfuego-cert` → Edit code → Deploy) cada vez
  que este archivo cambia. Si editas este archivo, avisa al usuario que debe
  volver a pegarlo y darle Deploy — un cambio aquí NO se aplica solo.
- El secret `GH_TOKEN` (un GitHub token con permiso de escritura al repo) vive
  **solo** en Cloudflare (Settings → Variables and Secrets), nunca en el código
  ni en el navegador de nadie.
- Usa la **Git Data API** de GitHub (blob → tree → commit → update ref) para
  subir archivos, no la API de "Contents" (`PUT /contents/{path}`) — esa tiene un
  límite práctico de ~1MB y los certificados con 2-3 fotos fácilmente lo superan.
  Si tocas el Worker, no vuelvas a la Contents API para escribir archivos grandes.
- Endpoints públicos (sin token, cualquiera puede llamarlos — el repo es público
  y de bajo riesgo, esto fue una decisión consciente, no un descuido):
  - `GET ?action=locales` — la libreta de locales completa.
  - `POST ?action=locales` con `{locales:[...]}` — guarda la libreta completa.
  - `POST` (sin `action`) con `{path, content}` (base64) — sube un certificado.
  - `GET ?action=list&dir=certificados/{MES}-{AÑO}` — lista los PDF de un mes.
  - `GET ?action=get&path=certificados/{MES}-{AÑO}/CERT-{CODIGO}-{AÑO}.pdf` —
    devuelve `{content: base64}`.
  - `GET ?action=selftest` — diagnóstico: prueba escribir/borrar un archivo y
    devuelve el error real de GitHub si algo falla.

### Los certificados de prueba (importante para cualquier tarea de datos)

**Todo lo generado en la app entre enero y julio de 2026 fue de prueba**, mientras
se armaba y depuraba el sistema (ver historial de commits: certificados de
A007, A020, V005, DI001 en `certificados/JULIO-2026/`). El uso real por los
técnicos empieza en **agosto de 2026**. Cualquier script o proceso que sincronice
o cuente certificados "reales" debe ignorar todo lo anterior a agosto-2026.

### La carpeta de Dropbox (independiente del repo/app)

Aparte de todo esto, la empresa ya tenía y sigue teniendo su propio archivo
histórico de certificados en Dropbox, en:

```
Previfuego / 2026 / 0 GRUPO KFC / CERTIFICADOS / [carpeta por mes]
```

Con carpetas de ENERO a JULIO ya existentes ahí (certificados reales de otra
fuente/proceso manual, sin relación con las pruebas de la app). El objetivo es
que, desde agosto en adelante, los certificados reales que la app genere
también terminen en esa misma carpeta de Dropbox — para eso existe
`scripts/sync_certificados_local.py`, que corre en la computadora del usuario
(esta sesión no tiene acceso a su filesystem local ni a Dropbox) y baja los PDF
del Worker hacia una carpeta local. **La convención exacta de nombres de mes
dentro de esa carpeta de Dropbox (mayúsculas, tildes, prefijos numéricos) hay
que verificarla mirando las carpetas ENERO-JULIO ya existentes** — no asumirla.

## Convenciones de trabajo en este repo

- **Rama de desarrollo**: `claude/youthful-fermi-a1719q`. Se hacen PRs contra
  `main` y se fusionan (squash) — el usuario pidió fusionar cada cambio
  automáticamente, sin esperar confirmación.
- **Versión visible**: `<meta name="app-version" content="X.Y">` en el `<head>`
  de `index.html`. Se sube en cada cambio que toque `index.html`, para que el
  usuario pueda confirmar visualmente (badge "vX.Y" junto al logo) que un
  dispositivo ya tomó la actualización. La app además revisa esta meta cada 45s
  y se autorecarga si cambió — no hace falta pedirle a nadie que borre caché.
- Como el historial en `main` tiene muchos commits directos (de otras sesiones o
  del propio usuario) además de merges por PR, `main` no siempre es
  fast-forward desde la rama de trabajo. Antes de abrir un PR nuevo: `git fetch
  origin main`, y si diverge, rebasar los commits propios sobre
  `origin/main` (cherry-pick, saltando los que queden vacíos porque ya están
  fusionados) antes de hacer push con `--force-with-lease`.
- No hay suite de tests. Verificación = revisar sintaxis JS (`new Function()`
  sobre el contenido de cada `<script>`) y, cuando es posible, un smoke test con
  Playwright headless. El entorno de esta sesión **no tiene salida de red hacia
  dominios como `workers.dev`** (política del sandbox) — para probar el Worker
  de verdad hay que pedirle al usuario que abra la URL él mismo, o usar la otra
  sesión local/Claude in Chrome que sí tiene acceso a su navegador real.
