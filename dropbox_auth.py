"""
dropbox_auth.py  –  Autenticación Dropbox por REFRESH TOKEN (permanente).

Uso INTERNO del backend (scripts de Python). NO se usa en la app web (index.html),
que trabaja solo con Gemini + GitHub.

🔒 SEGURIDAD: este archivo NO contiene credenciales (el repo es público).
   Las credenciales se leen de, en este orden:
     1) Variables de entorno:  DBX_APP_KEY, DBX_APP_SECRET, DBX_REFRESH_TOKEN
     2) Archivo local gitignored:  .dropbox_creds.json
        {"app_key": "...", "app_secret": "...", "refresh_token": "..."}

   Para que persista entre sesiones de forma segura, configura las 3 variables
   de entorno como SECRETS del entorno (no en el repo) o haz el repo privado.
"""

import os
import json
import requests

_CREDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".dropbox_creds.json")


def _load_creds() -> dict:
    key    = os.environ.get("DBX_APP_KEY")
    secret = os.environ.get("DBX_APP_SECRET")
    refresh = os.environ.get("DBX_REFRESH_TOKEN")
    if key and secret and refresh:
        return {"app_key": key, "app_secret": secret, "refresh_token": refresh}
    if os.path.exists(_CREDS_FILE):
        with open(_CREDS_FILE) as f:
            return json.load(f)
    raise RuntimeError(
        "Faltan credenciales de Dropbox. Define DBX_APP_KEY / DBX_APP_SECRET / "
        "DBX_REFRESH_TOKEN como variables de entorno, o crea .dropbox_creds.json "
        "(gitignored)."
    )


_cached_token = None


def get_access_token(force_refresh: bool = False) -> str:
    """Devuelve un access token válido, renovándolo con el refresh token."""
    global _cached_token
    if _cached_token and not force_refresh:
        return _cached_token
    c = _load_creds()
    r = requests.post(
        "https://api.dropboxapi.com/oauth2/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": c["refresh_token"],
            "client_id": c["app_key"],
            "client_secret": c["app_secret"],
        },
        timeout=60,
    )
    r.raise_for_status()
    _cached_token = r.json()["access_token"]
    return _cached_token


if __name__ == "__main__":
    tok = get_access_token()
    print("Access token OK:", tok[:35], "…")
    r = requests.post(
        "https://api.dropboxapi.com/2/files/list_folder",
        headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json"},
        data=json.dumps({"path": "/Previfuego/2026", "limit": 5}),
        timeout=60,
    )
    print("Test list:", r.status_code)
