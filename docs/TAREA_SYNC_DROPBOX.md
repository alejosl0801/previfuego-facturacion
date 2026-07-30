# Tarea: sincronizar certificados nuevos a la carpeta de Dropbox

> Antes de hacer nada, lee `CLAUDE.md` en la raíz del repo — tiene todo el
> contexto del proyecto (cómo funciona la app, el Worker, por qué los
> certificados de enero-julio 2026 son de prueba, etc.). Este documento asume
> que ya lo leíste y solo cubre esta tarea puntual.

## Objetivo

Que los certificados **reales** que la app genere desde agosto de 2026 en
adelante terminen automáticamente en la carpeta de Dropbox de la empresa, junto
a los certificados históricos de enero-julio que ya están ahí (de otra fuente,
sin relación con la app).

La carpeta de Dropbox (ruta lógica, dentro de la cuenta de Dropbox del usuario):

```
Previfuego / 2026 / 0 GRUPO KFC / CERTIFICADOS / [carpeta por mes]
```

Ya existen ahí las carpetas de ENERO a JULIO.

## Pasos

1. **Ubica la carpeta real en esta computadora.** Debería estar sincronizada
   localmente por la app de Dropbox, típicamente bajo `~/Dropbox/...` (Mac/Linux)
   o `C:\Users\<usuario>\Dropbox\...` (Windows). Búscala si no está en la
   ubicación por defecto.

2. **Verifica la convención exacta de nombres** mirando las carpetas ENERO a
   JULIO que ya existen ahí: ¿mayúsculas o no?, ¿con tildes?, ¿algún prefijo
   numérico (ej. "01 Enero")? Usa exactamente ese mismo patrón para todo lo que
   crees a continuación — no inventes un formato nuevo.

3. **Crea la carpeta de AGOSTO** dentro de `CERTIFICADOS`, siguiendo el patrón
   detectado en el paso 2.

4. **Adapta el script de sincronización.** Ya existe en el repo:
   `scripts/sync_certificados_local.py`
   (fuente: https://raw.githubusercontent.com/alejosl0801/previfuego-facturacion/main/scripts/sync_certificados_local.py)

   Este script baja de `certificados/AGOSTO-2026/` en adelante (vía el Worker
   público, sin token — ver `CLAUDE.md`) hacia una carpeta local, sin repetir
   descargas ya hechas. Actualmente:
   - Guarda todo bajo `CARPETA_DESTINO/{AÑO}/{MES}/`, con `{MES}` en MAYÚSCULAS
     SIN TILDES (así los nombra la app internamente — ENERO, FEBRERO, ...).
   - `CARPETA_DESTINO` apunta a una carpeta genérica de ejemplo
     (`~/PrevifuegoCertificados`).

   Ajusta:
   - `CARPETA_DESTINO` → la ruta real de `.../CERTIFICADOS` que encontraste en
     el paso 1.
   - Si el patrón de nombres de Dropbox (paso 2) es distinto al que usa el
     script por defecto (MAYÚSCULAS sin tildes), agrega un mapeo que traduzca
     el nombre interno de la app al nombre real de la carpeta de Dropbox, para
     que el PDF cargue en la carpeta que ya creaste y no en una nueva.
   - Si la estructura real de Dropbox no tiene un nivel de carpeta "por año"
     dentro de `CERTIFICADOS` (porque el año ya está más arriba, en
     `Previfuego/2026/...`), ajusta el script para que los PDF caigan
     directamente en `CERTIFICADOS/AGOSTO/`, sin crear un subnivel de año
     redundante.

5. **Prueba el script una vez a mano** (`python3 sync_certificados_local.py`).
   Es normal que no baje nada todavía si ningún técnico ha generado un
   certificado real desde agosto — eso no es un error.

6. **Si el usuario quiere que sea automático**, prográmalo (Programador de
   Tareas de Windows, o `cron`/`launchd` en Mac) para correr cada cierto tiempo
   (por ejemplo cada hora).

## Al terminar

Reporta en español, corto: la ruta exacta de Dropbox que usaste, qué
convención de nombres encontraste, y si quedó la sincronización automática
configurada o no.
