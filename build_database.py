#!/usr/bin/env python3
"""
build_database.py  –  Builds BASE_DATOS_KFC.xlsx (198 locales)

Sources (Dropbox):
  1. PRESUPUESTO PROVEEDORES 2026  → brand names (source of truth)
  2. MATRIZ EXTINTORES GRUPO KFC   → extintor data (primary)
  3. PROYECCIÓN INGRESOS MENSUAL 2026 → dates + data for missing locals

Logic:
  - 200 PRESUPUESTO PREVIFUEGO rows
  - Remove G042 (closed) → 199
  - H068 merges into K172 (not a separate local) → 198 unique locals
  - V091 not yet in PRESUPUESTO but IS in PROYECCIÓN → counted as the 198th
  - ALL locals are in OCT2026-SEP2027 recarga cycle
  - COBRO = RECARGA price (which already includes mantt)
  - Cycle sheet months: OCT/NOV/DEC → recarga 2026; JAN-SEP → recarga 2027
"""

import requests, io, re, time, json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from collections import defaultdict, Counter

TOKEN = (
    "sl.u.AGi_BG-2pjHUsNGNwXqO3-9F2zw0Ho_5by7OvUjvSERga-7WOZCiThedAHHGsQvRU9ngFrB_BDlbbft_eXoaGsNbI1OOmSXWBlCyYVWoF_F-ACSEh91E61VYOjtgKOc9f8FKYj2HyT4TLx4UTB-djpNthqbpCHLhQZMVUv_70laskwf2__wRkcRRPVgXifzMtv1-2Le00eVhlbiykP324HCcb73dMzvT4jQJucswMi2ZzZBiabYD_B-I1m0ugl6DXki56kdyza2YlMUnz6ihJuzPO66SRjnvq9Cf80q59zXWXcEhACPttZl3fmwA7TjliQh9bBkQ3qrjonaRiPRRRqxybaV_4v_5qZ6v431aSJRT2O3z_zPOCpl7eCyk0spR8fNV1ExXwvFKN1N3N4pMDcdAlWkZjg0g74WidzCQRf4fLvhJuf3-eFBBLnQeZHuxvfTQHJZbWrW3zGPu8fctyow14V4ejg37DELUHYXFU5KEOYS92dZzHXFq7WY0C-Drz563sVj1Mo2_kTQ68-HJheGLdWu_4Dsi1PNU4UCn0NePzsFWLKyWYN9o-byG8cpUNbCLbvKRXROecemUDPI1Hu1a2bT0lyR3JVFPtLtAuBUHHGiZIO2KKpmWKgUsaNkzYtOHM51_8EpFa_LYw59Cg3b_Q28GU3fYycG6fl0lDeLNzlcl8w7YamQtyMpXB_Miln3FGzIb6wsktmtkSsdAmswhQPrWTTYEkYEXjstzxdG7qvTIzNMhotEtn-RyzsJ6xAzDLEeLojvoufcaZx0JBz1i64Zv2D1o_BEKQcbFt5euKpa1V8JSdL_mcuFq8QNaDT1lJanha1BazKogV9KRDP1-SgmGsrFXbH0nH-PnOKQn3-Ntie2mYwbydRR45QvAIJBm2XhWEwJ6lWzg9TgOE_hixkDJGCrtJ7sQgHYd6mF4qmEPKNL-JBxEQbb0ZAFXaKl8S9r4lhjbeZys9J4FrGTuAn1hoyIePGNz7eqY4OUVYLcWeUTIitSgsoK0BbRbowY_KYybXLYZpNjauzEasRmFHWh3KnNfXvxHws-nVReHtj0Ofd-1E8I51ony8li6vOxcsMhF1CTyw4N7fuMJHge_QsVcEreUcMVjlXhBcmUKgmlnPqyzjs4ia44HYRsyAUs9XuI_un6EzYnZLEKTnHzP4-AYrlrn5y8PfIBEXFE9on9cAAiFz70uOjcPiWFd_XJhLuYHQG6QCy8ueG7Y09ErglEztMbZLxpkZk1QNbxOCpvBL_eUWDmRR1x1q-dFdeYg0U4fFuR_NlMM7fwNF1--"
)

FILES = {
    "PRESUPUESTO": "/Previfuego/2026/PRESUPUESTO PROVEEDORES AÑ0 2026.xlsx",
    "MATRIZ":      "/Previfuego/MATRIZ LOCALES/MATRIZ EXTINTORES GRUPO KFC.xlsx",
    "PROYECCION":  "/Previfuego/PRESUPUESTOS/PROYECCION INGRESOS MENSUAL 2026.xlsx",
}

MONTHS = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
          "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
MONTH_IDX = {m: i for i, m in enumerate(MONTHS)}
OCT_TO_DEC = {"OCTUBRE", "NOVIEMBRE", "DICIEMBRE"}

# ── Pricing ───────────────────────────────────────────────────────────────────

MANTT = {
    ("PQS",    "5"):   2.00,
    ("PQS",    "10"):  4.00,
    ("PQS",    "20"):  8.00,
    ("CO2",    "5"):   2.00,
    ("CO2",    "10"):  4.00,
    ("CO2",    "20"):  8.00,
    ("CO2",    "50"):  20.00,
    ("CO2",    "75"):  30.00,
    ("TIPO K", "2.5"): 8.00,
}
RECARGA = {
    ("PQS",    "5"):   4.90,
    ("PQS",    "10"):  9.80,
    ("PQS",    "20"):  19.60,
    ("CO2",    "5"):   4.50,
    ("CO2",    "10"):  9.00,
    ("CO2",    "20"):  18.00,
    ("CO2",    "50"):  45.00,
    ("CO2",    "75"):  67.50,
    ("TIPO K", "2.5"): 98.80,
}
CAP_DISPLAY = {
    ("PQS",    "5"):   "5 LBS",
    ("PQS",    "10"):  "10 LBS",
    ("PQS",    "20"):  "20 LBS",
    ("CO2",    "5"):   "5 LBS",
    ("CO2",    "10"):  "10 LBS",
    ("CO2",    "20"):  "20 LBS",
    ("CO2",    "50"):  "50 LBS",
    ("CO2",    "75"):  "75 LBS",
    ("TIPO K", "2.5"): "2.5 GLS",
}

# Brand fallback when not in PRESUPUESTO (for new locals K192, K194, V091, etc.)
BRAND_FALLBACK = {
    "K":  "KFC",
    "G":  "GUS",
    "M":  "MENESTRAS DEL NEGRO",
    "T":  "TROPIBURGER",
    "V":  "JUAN VALDEZ",
    "I":  "ILCAPPO",
    "J":  "CAJUN",
    "BS": "BASKIN ROBBINS",
    "CN": "CINNABON",
    "DI": "DOLCE INCONTRO",
    "CA": "CAFA",
    "A":  "AMERICAN DELI",
    "B":  "AMERICAN DELI",
    "F":  "AMERICAN DELI",
    "E":  "ESPAÑOL",
    "R":  "CASA RES",
    "H":  "HELADERIAS",
}

# Normalise brand names from PRESUPUESTO to consistent forms
MARCA_NORM = {
    "BASKIN":             "BASKIN ROBBINS",
    "DOLCE":              "DOLCE INCONTRO",
    "CAFA":               "CAFA",
    "GUS":                "GUS",
    "ILCAPPO":            "IL CAPPO",
    "HELADERIAS":         "HELADERIAS",
    "CASA RES":           "CASA RES",
    "MENESTRAS DEL NEGRO":"MENESTRAS DEL NEGRO",
}


def marca_norm(raw):
    if not raw: return ""
    s = str(raw).strip()
    return MARCA_NORM.get(s, s)


# ── Dropbox download ──────────────────────────────────────────────────────────

def dropbox_download(path, retries=4):
    delay = 2
    for attempt in range(retries + 1):
        try:
            r = requests.post(
                "https://content.dropboxapi.com/2/files/download",
                headers={
                    "Authorization": f"Bearer {TOKEN}",
                    "Dropbox-API-Arg": json.dumps({"path": path}),
                },
                timeout=90,
            )
            if r.status_code in (503, 429) and attempt < retries:
                print(f"  HTTP {r.status_code} – retry in {delay}s …")
                time.sleep(delay); delay *= 2; continue
            r.raise_for_status()
            return r.content
        except requests.RequestException as e:
            if attempt < retries:
                print(f"  Error: {e} – retry in {delay}s …")
                time.sleep(delay); delay *= 2
            else:
                raise


# ── Canonical CC key ──────────────────────────────────────────────────────────

def canonical(raw):
    """Return (prefix, number) tuple, or None if not parseable.
    Prefix is the canonical brand prefix (K, G, V, J, CN, BS, …).
    """
    if not raw:
        return None
    s = str(raw).strip().upper()
    s = re.sub(r'EC$', '', s).strip()

    for pat, pfx in [
        (r'^KFC(?:K+)?\s*(\d+)',  'K'),   # KFC02, KFCK192
        (r'^GUS\s*(\d+)',          'G'),   # GUS49
        (r'^JV\s*(\d+)',           'V'),   # JV15 → Juan Valdez
        (r'^JB\s*(\d+)',           'V'),   # JB = alternate Juan Valdez notation
        (r'^CJNC\s*(\d+)',         'J'),   # CJNC31 = CAJUN in MATRIZ
        (r'^CN\s*(\d+)',           'CN'),  # CN04 = CINNABON
        (r'^BS\s*(\d+)',           'BS'),  # BS04 = BASKIN ROBBINS
        (r'^DI\s*(\d+)',           'DI'),  # DI01 = DOLCE INCONTRO
        (r'^CA\s*(\d+)',           'CA'),  # CA01 = CAFA
        (r'^F\s*0*(\d+)',          'F'),   # F003 = AMERICAN DELI (Aeropuerto)
    ]:
        m = re.match(pat, s)
        if m:
            return (pfx, int(m.group(1)))

    # Single letter + optional space + digits
    m = re.match(r'^([A-Z])\s*(\d+)', s)
    if m:
        pfx, num = m.group(1), int(m.group(2))
        # J alone (not JV/JB) = CAJUN per user
        return (pfx, num)

    # Multi-letter + digits (handles ILCI02, AMCA19, ESPE07, TROT54 etc.)
    m = re.match(r'^([A-Z]{2,4})\s*(\d+)', s)
    if m:
        pfx, num = m.group(1), int(m.group(2))
        return (pfx, num)

    return None


def fmt_code(pfx, num):
    """Display code from canonical pair."""
    return f"{pfx}{num:03d}"


# ── Tipo / capacidad normalisation ────────────────────────────────────────────

def norm_tipo(raw):
    if raw is None: return ""
    s = str(raw).strip().upper()
    if re.match(r'^(K|TIPO\s*K|TYPE\s*K)$', s): return "TIPO K"
    if s.startswith("CO2"): return "CO2"
    if s.startswith("PQS"): return "PQS"
    return s


def norm_cap(raw):
    if raw is None: return ""
    if isinstance(raw, (int, float)):
        v = float(raw)
        return str(int(v)) if v == int(v) else str(round(v, 1))
    s = str(raw).strip().upper()
    s = re.sub(r'[A-Z\s]', '', s)
    s = s.replace(",,", ".").replace(",", ".")
    s = re.sub(r'\.{2,}', '.', s)
    m = re.match(r'(\d+(?:\.\d+)?)', s)
    if m:
        v = float(m.group(1))
        return str(int(v)) if v == int(v) else str(round(v, 1))
    return ""


def norm_tipo_cap(tipo_r, cap_r):
    tipo = norm_tipo(tipo_r)
    cap  = norm_cap(cap_r)
    # PQS 2.5 doesn't exist → TIPO K 2.5 GLS
    if tipo == "PQS" and cap == "2.5":
        tipo = "TIPO K"
    # Cap raw has 'G' (gallons) and value is 2.5 → TIPO K
    if cap == "2.5" and cap_r and re.search(r'G', str(cap_r).upper()):
        tipo = "TIPO K"
    return tipo, cap


def to_year(v):
    if v is None: return ""
    try: return str(int(float(str(v))))
    except: return str(v).strip()


# ── Download ──────────────────────────────────────────────────────────────────

print("Downloading PRESUPUESTO …")
data_pres = dropbox_download(FILES["PRESUPUESTO"])
print(f"  {len(data_pres):,} bytes")

print("Downloading MATRIZ …")
data_mat = dropbox_download(FILES["MATRIZ"])
print(f"  {len(data_mat):,} bytes")

print("Downloading PROYECCION …")
data_proy = dropbox_download(FILES["PROYECCION"])
print(f"  {len(data_proy):,} bytes")

wb_pres = openpyxl.load_workbook(io.BytesIO(data_pres), data_only=True)
wb_mat  = openpyxl.load_workbook(io.BytesIO(data_mat),  data_only=True)
wb_proy = openpyxl.load_workbook(io.BytesIO(data_proy), data_only=True)


# ── 1. Parse PRESUPUESTO (sheet SSO) ──────────────────────────────────────────
# Structure: header at row 4 (0-indexed)
#   col 1=EMPRESA, col 2=LOCAL (CC), col 3=C.COSTO, col 4=MARCA,
#   col 5=NOMBRE LOCAL, col 6=RESPONSABLE, col 7=Proveedor

print("\n=== Parsing PRESUPUESTO ===")
presupuesto_map = {}  # canonical_key -> {marca, nombre, cc_pres}

ws_sso = wb_pres["SSO"]
pres_rows = list(ws_sso.iter_rows(values_only=True))

for row in pres_rows[5:]:  # skip title + blank + header rows
    if not row or len(row) < 8:
        continue
    proveedor = str(row[7]).strip() if row[7] else ""
    if proveedor.upper() != "PREVIFUEGO":
        continue
    cc_raw = row[2]
    marca  = str(row[4]).strip() if row[4] else ""
    nombre = str(row[5]).strip() if row[5] else ""
    ckey = canonical(cc_raw)
    if not ckey:
        continue
    presupuesto_map[ckey] = {
        "marca":   marca_norm(marca),
        "nombre":  nombre,
        "cc_pres": str(cc_raw).strip() if cc_raw else "",
    }

print(f"PRESUPUESTO entries: {len(presupuesto_map)}")

# Confirm G042 and H068 are there
G042_KEY = ('G', 42)
H068_KEY = ('H', 68)
K172_KEY = ('K', 172)
print(f"  G042 in PRESUPUESTO: {G042_KEY in presupuesto_map}")
print(f"  H068 in PRESUPUESTO: {H068_KEY in presupuesto_map}")


# ── 2. Parse MATRIZ and PROYECCIÓN ────────────────────────────────────────────

SKIP_TIPO = re.compile(r'^(MANTENIMIENTO|SIST|SISTEMA|TOTAL|SUBTOTAL)', re.I)

def parse_extintor_sheet(ws, mes):
    """Return list of extintor dicts from a monthly sheet."""
    all_rows = list(ws.iter_rows(values_only=True))

    # Find header row (first row with ≥2 header-like keywords)
    header_idx = 0
    for i, row in enumerate(all_rows[:10]):
        cells = [str(c).strip().upper() if c else "" for c in row]
        hits = sum(1 for c in cells if c in ("CC","TIPO","CAPACIDAD","UBICACION","UBICACIÓN","LOCAL","NOMBRE"))
        if hits >= 2:
            header_idx = i
            break

    hrow = [str(c).strip().upper() if c else "" for c in all_rows[header_idx]]

    def find_col(*names):
        for name in names:
            for j, h in enumerate(hrow):
                if name in h:
                    return j
        return None

    c_cc   = find_col("CC") or 0
    c_ubic = find_col("UBICACION", "UBICACIÓN", "NOMBRE", "LOCAL") or 1
    c_tipo = find_col("TIPO") or 2
    c_cap  = find_col("CAPACIDAD", "CAP") or 3
    c_ult  = find_col("ULT", "ÚLTIMA", "ULTIMA", "ANTERIOR") or 9
    c_prox = find_col("PRÓX", "PROX", "SIGUIENTE") or 10

    rows_out = []
    cur_ckey   = None
    cur_cc_raw = None

    for row in all_rows[header_idx + 1:]:
        if not row or all(c is None for c in row):
            continue

        def gc(idx):
            if idx is not None and idx < len(row) and row[idx] is not None:
                return str(row[idx]).strip()
            return ""

        cc_raw  = row[c_cc] if c_cc < len(row) else None
        ubic    = gc(c_ubic)
        tipo_r  = row[c_tipo] if c_tipo < len(row) else None
        cap_r   = row[c_cap]  if c_cap  < len(row) else None
        ult_r   = row[c_ult]  if c_ult  < len(row) else None
        prox_r  = row[c_prox] if c_prox < len(row) else None

        # Forward-fill CC
        if cc_raw is not None and str(cc_raw).strip():
            ckey = canonical(str(cc_raw).strip())
            if ckey:
                cur_ckey   = ckey
                cur_cc_raw = str(cc_raw).strip()

        if not cur_ckey:
            continue

        tipo, cap = norm_tipo_cap(tipo_r, cap_r)

        if not tipo or tipo not in ("PQS", "CO2", "TIPO K"):
            continue
        if tipo_r and SKIP_TIPO.match(str(tipo_r)):
            continue

        rows_out.append({
            "ckey":    cur_ckey,
            "cc_raw":  cur_cc_raw,
            "mes":     mes,
            "ubic":    ubic,
            "tipo":    tipo,
            "cap":     cap,
            "año_ult": to_year(ult_r),
            "año_prox": to_year(prox_r),
        })

    return rows_out


print("\n=== Parsing MATRIZ ===")
matriz_rows = []
for sname in wb_mat.sheetnames:
    mes = sname.strip().upper()
    if mes not in MONTHS:
        continue
    rows = parse_extintor_sheet(wb_mat[sname], mes)
    print(f"  {sname}: {len(rows)} extintor rows")
    matriz_rows.extend(rows)

print(f"\nTotal MATRIZ extintor rows: {len(matriz_rows)}")
mat_keys = set(r["ckey"] for r in matriz_rows)
print(f"Unique MATRIZ locals:       {len(mat_keys)}")


print("\n=== Parsing PROYECCIÓN ===")
proy_rows = []
for sname in wb_proy.sheetnames:
    mes = sname.strip().upper()
    if mes not in MONTHS:
        continue
    rows = parse_extintor_sheet(wb_proy[sname], mes)
    print(f"  {sname}: {len(rows)} extintor rows")
    proy_rows.extend(rows)

proy_keys = set(r["ckey"] for r in proy_rows)
print(f"\nTotal PROYECCIÓN extintor rows: {len(proy_rows)}")
print(f"Unique PROYECCIÓN locals:       {len(proy_keys)}")

only_proy = proy_keys - mat_keys
print(f"\nIn PROYECCIÓN but NOT MATRIZ: {sorted(only_proy)}")


# ── 3. Apply corrections ──────────────────────────────────────────────────────

print("\n=== Applying corrections ===")

# (a) Remove G042 (closed)
before = len(matriz_rows)
matriz_rows = [r for r in matriz_rows if r["ckey"] != G042_KEY]
print(f"Removed G042: {before - len(matriz_rows)} rows removed")

# (b) H068 → already absent from MATRIZ (its extintores were merged into K172
#     which already includes 'heladería' extintor; no additional rows needed)
h068_in_mat = [r for r in matriz_rows if r["ckey"] == H068_KEY]
print(f"H068 rows in MATRIZ: {len(h068_in_mat)} (expected 0 – already part of K172)")

# (c) Add V091 from PROYECCIÓN if not in MATRIZ
V091_KEY = ('V', 91)
if V091_KEY not in mat_keys:
    proy_v091 = [r for r in proy_rows if r["ckey"] == V091_KEY]
    # Deduplicate (PROYECCIÓN sometimes lists each extintor twice)
    seen = set()
    deduped = []
    for r in proy_v091:
        sig = (r["tipo"], r["cap"], r["ubic"])
        if sig not in seen:
            seen.add(sig)
            deduped.append(r)
    print(f"Adding V091 from PROYECCIÓN: {len(deduped)} extintores (from {len(proy_v091)} raw rows)")
    for r in deduped:
        print(f"  {r['tipo']} {r['cap']}")
    matriz_rows.extend(deduped)
else:
    print("V091 already in MATRIZ")

# (d) Override MES from PROYECCIÓN where PROYECCIÓN is source of truth
proy_mes = {}
for r in proy_rows:
    if r["ckey"] not in proy_mes:
        proy_mes[r["ckey"]] = r["mes"]

overrides = 0
for r in matriz_rows:
    pm = proy_mes.get(r["ckey"])
    if pm and pm != r["mes"]:
        r["mes"] = pm
        overrides += 1
print(f"\nMES overridden from PROYECCIÓN: {overrides} rows")

# Final unique locals
final_keys = set(r["ckey"] for r in matriz_rows)
print(f"\nFinal unique locals: {len(final_keys)}")


# ── 4. Brand lookup ───────────────────────────────────────────────────────────

def get_marca(ckey):
    pinfo = presupuesto_map.get(ckey)
    if pinfo and pinfo["marca"]:
        return pinfo["marca"]
    pfx = ckey[0]
    return BRAND_FALLBACK.get(pfx, BRAND_FALLBACK.get(pfx[:1], "DESCONOCIDO"))


def get_nombre_local(ckey):
    pinfo = presupuesto_map.get(ckey)
    if pinfo and pinfo["nombre"]:
        return pinfo["nombre"]
    return ""


# ── 5. Build final rows ───────────────────────────────────────────────────────

print("\n=== Building final rows ===")

final_rows = []
for r in matriz_rows:
    ckey  = r["ckey"]
    mes   = r["mes"]
    tipo  = r["tipo"]
    cap   = r["cap"]

    pk  = (tipo, cap)
    cm  = MANTT.get(pk)
    cr  = RECARGA.get(pk)
    cd  = CAP_DISPLAY.get(pk, f"{cap}")

    # ALL locals are in OCT2026–SEP2027 recarga cycle
    # COBRO = RECARGA (includes mantt). If price unknown, fall back to MANTT.
    cobro      = cr if cr is not None else (cm if cm is not None else None)
    cobro_tipo = "RECARGA (incl. mantt)" if cr is not None else ("MANTT" if cm is not None else "SIN PRECIO")

    ano_rec     = 2026 if mes in OCT_TO_DEC else 2027
    ano_ult_rec = ano_rec - 3

    marca  = get_marca(ckey)
    nombre = get_nombre_local(ckey)
    codigo = fmt_code(*ckey)

    notas = []
    if cm is None:
        notas.append(f"precio no hallado ({tipo}/{cap})")

    final_rows.append({
        "CÓDIGO":          codigo,
        "CC_ORIGINAL":     r.get("cc_raw", ""),
        "MARCA":           marca,
        "NOMBRE_LOCAL":    nombre,
        "MES_SERVICIO":    mes,
        "UBICACIÓN":       r.get("ubic", ""),
        "TIPO":            tipo,
        "CAPACIDAD":       cd,
        "COSTO_MANTT":     round(cm, 2) if cm is not None else None,
        "PRECIO_RECARGA":  round(cr, 2) if cr is not None else None,
        "COBRO_ANUAL_EXT": round(cobro, 2) if cobro is not None else None,
        "TIPO_COBRO":      cobro_tipo,
        "AÑO_RECARGA":     ano_rec,
        "AÑO_ULT_RECARGA": ano_ult_rec,
        "NOTAS":           "; ".join(notas),
        "_cm":   cm   or 0.0,
        "_cr":   cr   or 0.0,
        "_cobro": cobro or 0.0,
        "_ckey": ckey,
    })

print(f"Total extintor rows: {len(final_rows)}")
print(f"Unique locals:       {len(set(r['CÓDIGO'] for r in final_rows))}")


# ── 6. Per-local summary ──────────────────────────────────────────────────────

local_sum = {}
for r in final_rows:
    k = r["CÓDIGO"]
    if k not in local_sum:
        local_sum[k] = {
            "CÓDIGO":          k,
            "CC_ORIGINAL":     r["CC_ORIGINAL"],
            "MARCA":           r["MARCA"],
            "NOMBRE_LOCAL":    r["NOMBRE_LOCAL"],
            "MES_SERVICIO":    r["MES_SERVICIO"],
            "N_EXTINTORES":    0,
            "TOTAL_MANTT":     0.0,
            "TOTAL_RECARGA":   0.0,
            "COBRO_ANUAL":     0.0,
            "AÑO_RECARGA":     r["AÑO_RECARGA"],
            "AÑO_ULT_RECARGA": r["AÑO_ULT_RECARGA"],
            "_desglose":       [],
        }
    s = local_sum[k]
    s["N_EXTINTORES"]  += 1
    s["TOTAL_MANTT"]   += r["_cm"]
    s["TOTAL_RECARGA"] += r["_cr"]
    s["COBRO_ANUAL"]   += r["_cobro"]
    s["_desglose"].append(
        f"{r['TIPO']} {r['CAPACIDAD']}=${r['_cobro']:.2f}"
    )

for s in local_sum.values():
    s["TOTAL_MANTT"]  = round(s["TOTAL_MANTT"],  2)
    s["TOTAL_RECARGA"]= round(s["TOTAL_RECARGA"], 2)
    s["COBRO_ANUAL"]  = round(s["COBRO_ANUAL"],  2)
    s["DESGLOSE"]     = " | ".join(s.pop("_desglose"))

n_locales   = len(local_sum)
total_anual = sum(s["COBRO_ANUAL"] for s in local_sum.values())

print(f"\n{'='*50}")
print(f"LOCALES TOTAL: {n_locales}  (target: 198)")
print(f"COBRO ANUAL TOTAL: ${total_anual:,.2f}")
print(f"{'='*50}")

print("\nLocales por mes:")
mc = Counter(s["MES_SERVICIO"] for s in local_sum.values())
for m in MONTHS:
    if mc.get(m): print(f"  {m:15s}: {mc[m]}")

print("\nLocales por marca:")
bc = Counter(s["MARCA"] for s in local_sum.values())
for b, c in sorted(bc.items(), key=lambda x: -x[1]):
    print(f"  {b:25s}: {c}")

# Audit: in PRESUPUESTO (minus G042, H068) but not in final
pres_expected = set(k for k in presupuesto_map if k not in (G042_KEY, H068_KEY))
final_codes   = set(r["_ckey"] for r in final_rows)
missing_from_final = pres_expected - final_codes
extra_in_final     = final_codes - pres_expected - {H068_KEY}

print(f"\nIn PRESUPUESTO but missing from DB ({len(missing_from_final)}):")
for k in sorted(missing_from_final):
    info = presupuesto_map[k]
    print(f"  {fmt_code(*k)}: {info['cc_pres']} – {info['marca']} – {info['nombre']}")

print(f"\nIn DB but not in PRESUPUESTO ({len(extra_in_final)}):")
for k in sorted(extra_in_final):
    print(f"  {fmt_code(*k)}")


# ── 7. Write Excel ────────────────────────────────────────────────────────────

OUT = "/home/user/previfuego-facturacion/BASE_DATOS_KFC.xlsx"
wb = openpyxl.Workbook()

H_FILL  = PatternFill("solid", fgColor="1F4E79")
H2_FILL = PatternFill("solid", fgColor="375623")
H3_FILL = PatternFill("solid", fgColor="7B0000")
ALT1    = PatternFill("solid", fgColor="D6E4F0")
ALT2    = PatternFill("solid", fgColor="E2EFDA")
H_FONT  = Font(color="FFFFFF", bold=True)
CENTER  = Alignment(horizontal="center")

def style_header(ws, fill):
    for cell in ws[1]:
        cell.fill = fill; cell.font = H_FONT; cell.alignment = CENTER

def auto_width(ws):
    for col in ws.columns:
        mx = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(mx + 2, 55)

month_sort = {m: i for i, m in enumerate(MONTHS)}


# Sheet 1: DETALLE
ws_det = wb.active
ws_det.title = "DETALLE"
COLS_DET = [
    "CÓDIGO","CC_ORIGINAL","MARCA","NOMBRE_LOCAL","MES_SERVICIO","UBICACIÓN",
    "TIPO","CAPACIDAD","COSTO_MANTT","PRECIO_RECARGA","COBRO_ANUAL_EXT",
    "TIPO_COBRO","AÑO_RECARGA","AÑO_ULT_RECARGA","NOTAS"
]
ws_det.append(COLS_DET)
style_header(ws_det, H_FILL)

for idx, r in enumerate(
    sorted(final_rows, key=lambda x: (month_sort.get(x["MES_SERVICIO"], 99), x["CÓDIGO"])),
    2
):
    ws_det.append([r.get(c) for c in COLS_DET])
    if idx % 2 == 0:
        for cell in ws_det[idx]: cell.fill = ALT1

auto_width(ws_det)


# Sheet 2: RESUMEN_LOCALES
ws_res = wb.create_sheet("RESUMEN_LOCALES")
COLS_RES = [
    "CÓDIGO","CC_ORIGINAL","MARCA","NOMBRE_LOCAL","MES_SERVICIO",
    "N_EXTINTORES","TOTAL_MANTT","TOTAL_RECARGA","COBRO_ANUAL",
    "AÑO_RECARGA","AÑO_ULT_RECARGA","DESGLOSE"
]
ws_res.append(COLS_RES)
style_header(ws_res, H2_FILL)

for idx, s in enumerate(
    sorted(local_sum.values(), key=lambda x: (month_sort.get(x["MES_SERVICIO"], 99), x["CÓDIGO"])),
    2
):
    ws_res.append([s.get(c) for c in COLS_RES])
    if idx % 2 == 0:
        for cell in ws_res[idx]: cell.fill = ALT2

auto_width(ws_res)


# Sheet 3: AUDITORÍA
ws_aud = wb.create_sheet("AUDITORÍA")
ws_aud.append(["TIPO", "CÓDIGO", "DETALLE"])
style_header(ws_aud, H3_FILL)

for k in sorted(missing_from_final):
    info = presupuesto_map[k]
    ws_aud.append(["EN_PRESUP_SIN_EXTINTOR", fmt_code(*k),
                   f"{info['cc_pres']} – {info['marca']} – {info['nombre']}"])

for k in sorted(extra_in_final):
    ws_aud.append(["EN_MATRIZ_SIN_PRESUP", fmt_code(*k), "Nuevo local no en PRESUPUESTO 2026"])

for r in final_rows:
    if r["NOTAS"]:
        ws_aud.append(["PRECIO_NO_HALLADO", r["CÓDIGO"],
                       f"{r['TIPO']} {r['CAPACIDAD']} – {r['NOTAS']}"])

ws_aud.append(["FUSIONADO", "H068",
               "Heladerías Mall del Norte fusionado en K172 (extintor 'heladería' incluido en K172)"])
ws_aud.append(["ELIMINADO", "G042",
               "GUS Quito Aguirre – local cerrado, excluido de la base"])

auto_width(ws_aud)


wb.save(OUT)
print(f"\n✓ Saved: {OUT}")
print(f"  DETALLE rows:  {len(final_rows)}")
print(f"  Locales:       {n_locales}")
print(f"  Cobro anual:   ${total_anual:,.2f}")
