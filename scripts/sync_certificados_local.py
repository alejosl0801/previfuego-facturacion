#!/usr/bin/env python3
"""
Sincroniza los certificados PDF de Previfuego a una carpeta local.

Descarga cada certificado generado desde AGOSTO-2026 en adelante (todo lo
anterior fueron pruebas y se ignora a propósito) a una carpeta en esta
computadora, organizada por año y mes. Ya descargados no se vuelven a bajar,
así que se puede correr una y otra vez sin duplicar trabajo — pensado para
ejecutarse periódicamente (Tarea programada de Windows, cron/launchd, o a
mano) desde tu propia computadora.

No necesita ningún token: usa el mismo Worker público que usa la app.

Uso:
    python3 sync_certificados_local.py

Configuración: ver las constantes de abajo (carpeta de destino y el
mes/año a partir del cual se considera "real", no prueba).
"""
import base64
import datetime
import json
import os
import urllib.error
import urllib.parse
import urllib.request

WORKER_URL = 'https://previfuego-cert.alejosl0801.workers.dev'
CARPETA_DESTINO = os.path.join(os.path.expanduser('~'), 'PrevifuegoCertificados')

# Todo lo generado ANTES de este mes/año fueron certificados de prueba —
# se ignoran a propósito, no se descargan.
CUTOFF_ANIO = 2026
CUTOFF_MES = 'AGOSTO'

MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
         'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']


def carpetas_a_revisar():
    """Genera pares (mes, año) desde el cutoff hasta el año que viene."""
    anio_actual = datetime.date.today().year
    idx_cutoff = MESES.index(CUTOFF_MES)
    for anio in range(CUTOFF_ANIO, anio_actual + 2):
        inicio = idx_cutoff if anio == CUTOFF_ANIO else 0
        for mes in MESES[inicio:]:
            yield mes, anio


def get_json(url):
    with urllib.request.urlopen(url, timeout=20) as r:
        return json.loads(r.read().decode('utf-8'))


def main():
    os.makedirs(CARPETA_DESTINO, exist_ok=True)
    nuevos = 0
    revisados = 0

    for mes, anio in carpetas_a_revisar():
        dir_repo = f'certificados/{mes}-{anio}'
        url_list = f'{WORKER_URL}/?action=list&dir={urllib.parse.quote(dir_repo)}'
        try:
            data = get_json(url_list)
        except (urllib.error.URLError, json.JSONDecodeError) as e:
            print(f'  [aviso] no se pudo revisar {dir_repo}: {e}')
            continue

        archivos = data.get('files', [])
        if not archivos:
            continue

        carpeta_local = os.path.join(CARPETA_DESTINO, str(anio), mes)
        for f in archivos:
            revisados += 1
            destino = os.path.join(carpeta_local, f['name'])
            if os.path.exists(destino):
                continue
            url_get = f'{WORKER_URL}/?action=get&path={urllib.parse.quote(f["path"])}'
            try:
                d = get_json(url_get)
                contenido = d.get('content')
                if not contenido:
                    print(f'  [aviso] sin contenido para {f["path"]}: {d.get("error")}')
                    continue
                os.makedirs(carpeta_local, exist_ok=True)
                with open(destino, 'wb') as fh:
                    fh.write(base64.b64decode(contenido))
                nuevos += 1
                print(f'  descargado: {anio}/{mes}/{f["name"]}')
            except (urllib.error.URLError, json.JSONDecodeError) as e:
                print(f'  [error] {f["path"]}: {e}')

    print(f'\nListo. {nuevos} certificado(s) nuevo(s) descargado(s) '
          f'({revisados} revisado(s) en total).')
    print(f'Carpeta: {CARPETA_DESTINO}')


if __name__ == '__main__':
    main()
