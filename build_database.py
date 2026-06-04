#!/usr/bin/env python3
"""
build_database.py  –  Builds BASE_DATOS_KFC.xlsx (198 locales)

Sources (Dropbox):
  1. PRESUPUESTO PROVEEDORES 2026  → brand names (source of truth)
  2. MATRIZ EXTINTORES GRUPO KFC   → extintor data (primary)
  3. PROYECCIÓN INGRESOS MENSUAL 2026 → dates + data for K079/K192/K194
"""

import requests, io, re, time, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from collections import defaultdict

TOKEN = (
    "sl.u.AGgF5XKR5BwOCmo95YjLwdRE-buSMDIWiOPYXcWcE7SUTTaHr9AIpXz-9MVh42IhLVZsHJCrKnkxOwzu8i5hgLrIuJC8XlRGaCprnDHwr1Czw7hxc92BXwjNMEAMvR0iXS88gLeyV_v4Wbj4KCTCMX7qlmtztJWZSS1oN0mfnHX2dj2SP2A_4Ari0pVrlMqbUbnWlRGYRNgHrMA7K6bbUCfeBq-3IO8VaTTkZEnyYNMKPHUHu1CHyxypSrC44D5jlBEyPLguvyKppj_qwNQkfJSHHhLv-VvqohTPLRMewH0inx-UpIcq7ku28sGrd04g2-R2YJUOcEbAWOgu5Uwm2VCXXwTxTlbBPV2NQe--n_ZT_uzACDQHLV0E9c0FIA9Z8_v56DGJch7mMmR7TbavjM9v2yFDzf8Jdw4ip3LCYS1yhkHwV-_CTZ8fFdTcJJ5gm4CbjCGhI-k9eGvmp4xFeId6Wra8yWLA_DZd-Q11gHbb-FY7vCD7hWvNi0B_3N86FG4S1ha_0loOmJkav9dk6L81RhIcAzLiMsN64rO0hTS4gugUGT8MuXm9UbAwNL_7untY5spa_8ecfefeC894o-Ae3TVeAji3dvcrUYsgClYgK1fkylqt1SaeOj1K0PZFcjmxin-K9BUvfGTjUclowh-HxEDl7n6-QofS1DqczB0loetZV27LHZFoBGl4FDRpSDhydxE9XkwyoIXg6LSfUUNCwi9fFlBnISrO5eZaGrCjpqpftgEug_jPjRLSILFra6uakJOVT8gS-aZoWBZv9YacSnrJ6iuRf0mRo6X7nV9SOm2cmdtyUZvdxBxtjTfaQfFWD8VkHHmTHLR-Fmt50on5DDaocHhfzkJt8wyafU4YH7NSbz2mhfFmxcb1s5WEQrgslgW5dKmmi6VGBIcPDMeIPxoRZe43aodjcWYzZmOT7czMjPpI5J6CNalK8HZMI43oRVwKAml3TXMEg0CLhtlMRdsafEo3rIdoB81ZK3RbnnGTCXkXuXKDdzd1qHRyE3Z2DCbcHFYvwzYcT6gr4VnIc4WmsTt7f4s4bGoaHoJxD8K1zOn85J_VLHJBOXdknyDSyHxhNPSKebKo3_pAFXJe4Q4HoaCrGyxHH09xGS_1TteDVPnHqRUFK5fIa1EezABLZfDKnvKOBxLeaTe9AEO89ddAiYvpJaGkpCKKG0GIOtivpsNRwi4tczW06bmZZAUPLQFjkWhTcxvh_7QHYG_Jm63xC1B7st_7Zkez4HnAzUsTCEucTpIpTWvwAcEuu-5cgRUdbe2XHjeSeeFUUmPB"
)

FILES = {
    "PRESUPUESTO": "/Previfuego/2026/PRESUPUESTO PROVEEDORES AÑ0 2026.xlsx",
    "MATRIZ":      "/Previfuego/MATRIZ LOCALES/MATRIZ EXTINTORES GRUPO KFC.xlsx",
    "PROYECCION":  "/Previfuego/PRESUPUESTOS/PROYECCION INGRESOS MENSUAL 2026.xlsx",
}


def dropbox_download(path, retries=4):
    delay = 2
    for attempt in range(retries + 1):
        try:
            r = requests.post(
                "https://content.dropboxapi.com/2/files/download",
                headers={
                    "Authorization": f"Bearer {TOKEN}",
                    "Dropbox-API-Arg": f'{{"path": "{path}"}}',
                },
                timeout=90,
            )
            if r.status_code in (503, 429) and attempt < retries:
                print(f"  {r.status_code} – retry in {delay}s …")
                time.sleep(delay); delay *= 2; continue
            r.raise_for_status()
            return r.content
        except requests.RequestException as e:
            if attempt < retries:
                print(f"  Error: {e} – retry in {delay}s …")
                time.sleep(delay); delay *= 2
            else:
                raise
    return None


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

MONTHS = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
          "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
MONTH_IDX = {m: i for i, m in enumerate(MONTHS)}


# ── Canonical CC ──────────────────────────────────────────────────────────────

def canonical(raw):
    """Return (prefix, number) canonical pair. prefix ∈ {K,G,V,M,T,I,J,BS,CN,DI,CA,F,…}"""
    if not raw:
        return None
    s = str(raw).strip().upper()
    s = re.sub(r'EC$', '', s).strip()

    patterns = [
        (r'^KFC(?:K+)?\s*(\d+)',   'K'),
        (r'^GUS\s*(\d+)',           'G'),
        (r'^JV\s*(\d+)',            'V'),   # Juan Valdez
        (r'^(?:CJNC)\s*(\d+)',     'J'),   # CJNC = CAJUN
        (r'^CN\s*(\d+)',            'CN'),  # CINNABON
        (r'^BS\s*(\d+)',            'BS'),  # BASKIN ROBBINS
        (r'^DI\s*(\d+)',            'DI'),
        (r'^CA\s*(\d+)',            'CA'),
        (r'^[JB]{1,2}\s*(\d+)',     'J'),   # J, JB, B – but B/JB are Juan Valdez?
    ]
    # Actually:  B and JB = Juan Valdez per user; J = CAJUN
    # So separate them:
    for pat, pfx in [
        (r'^KFC(?:K+)?\s*(\d+)',   'K'),
        (r'^GUS\s*(\d+)',           'G'),
        (r'^JV\s*(\d+)',            'V'),
        (r'^JB\s*(\d+)',            'V'),   # JB = Juan Valdez
        (r'^CJNC\s*(\d+)',          'J'),   # CAJUN (wrong notation in MATRIZ)
        (r'^CN\s*(\d+)',            'CN'),
        (r'^BS\s*(\d+)',            'BS'),
        (r'^DI\s*(\d+)',            'DI'),
        (r'^CA\s*(\d+)',            'CA'),
        (r'^F\s*(\d+)',             'F'),
    ]:
        m = re.match(pat, s)
        if m:
            return (pfx, int(m.group(1)))

    # Single-letter + optional space + number
    m = re.match(r'^([A-Z])\s+(\d+)', s)
    if m:
        ltr, num = m.group(1), int(m.group(2))
        if ltr == 'B':
            ltr = 'V'   # B = Juan Valdez
        return (ltr, num)

    # Letter(s) + number (no space)
    m = re.match(r'^([A-Z]{1,3})(\d+)', s)
    if m:
        pfx, num = m.group(1), int(m.group(2))
        if pfx == 'B':
            pfx = 'V'
        return (pfx, num)

    return None


def fmt_code(pfx, num):
    """Format canonical pair back to display code."""
    if pfx in ('K','G','M','T','V','I','J','F'):
        return f"{pfx}{num:03d}"
    if pfx in ('BS','CN','DI','CA'):
        return f"{pfx}{num:03d}"
    return f"{pfx}{num:03d}"


def norm_tipo(raw):
    if raw is None: return ""
    s = str(raw).strip().upper()
    if re.match(r'^(K|TIPO\s*K|TYPE\s*K)$', s): return "TIPO K"
    if s.startswith("CO2"): return "CO2"
    if s.startswith("PQS"): return "PQS"
    return s


def norm_cap(raw):
    """Return numeric string: '5','10','20','50','75','2.5'"""
    if raw is None: return ""
    if isinstance(raw, (int, float)):
        v = float(raw)
        return str(int(v)) if v == int(v) else str(round(v, 1))
    s = str(raw).strip().upper()
    # Remove letter suffixes, fix double-comma/period
    s = re.sub(r'[A-Z\s]', '', s)
    s = s.replace(",,", ".").replace(",", ".")
    s = re.sub(r'\.+', '.', s)
    m = re.match(r'(\d+(?:\.\d+)?)', s)
    if m:
        v = float(m.group(1))
        return str(int(v)) if v == int(v) else str(round(v, 1))
    return ""


def norm_tipo_cap(tipo_r, cap_r):
    """Normalize tipo and cap, applying corrections."""
    tipo = norm_tipo(tipo_r)
    cap  = norm_cap(cap_r)
    # PQS 2.5 doesn't exist → must be TIPO K 2.5 GLS
    if tipo == "PQS" and cap == "2.5":
        tipo = "TIPO K"
    # If cap raw contains "G" and value is 2.5 → TIPO K
    if cap == "2.5" and cap_r and re.search(r'G', str(cap_r).upper()):
        tipo = "TIPO K"
    return tipo, cap


def to_year(v):
    if v is None: return ""
    try: return str(int(float(str(v))))
    except: return str(v).strip()


# ── Download files ────────────────────────────────────────────────────────────

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

print(f"\nPRESUPUESTO sheets: {wb_pres.sheetnames}")
print(f"MATRIZ sheets:      {wb_mat.sheetnames}")
print(f"PROYECCION sheets:  {wb_proy.sheetnames}")


# ── 1. Parse PRESUPUESTO – build {canonical_key: {empresa, marca, nombre}} ────
# The PRESUPUESTO has one row per local with columns including CC code, empresa, marca.
# We'll scan for "PREVIFUEGO" in each row to keep only our locals.

print("\n=== Parsing PRESUPUESTO ===")
presupuesto_map = {}   # canonical_key -> {empresa, marca, nombre_local, cc_pres}
presupuesto_raw = []   # for debugging

for sheet_name in wb_pres.sheetnames:
    ws = wb_pres[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    print(f"  Sheet '{sheet_name}': {len(rows)} rows")

    # Find header row: look for row containing "CC" and "EMPRESA" or "MARCA"
    header_idx = 0
    col_cc = col_empresa = col_marca = col_nombre = col_proveedor = None

    for i, row in enumerate(rows[:10]):
        cells = [str(c).strip().upper() if c else "" for c in row]
        if any("CC" in c for c in cells):
            # Try to identify columns
            for j, c in enumerate(cells):
                if c == "CC" or c == "CÓDIGO" or c == "CODIGO":
                    col_cc = j
                elif "EMPRESA" in c:
                    col_empresa = j
                elif "MARCA" in c:
                    col_marca = j
                elif "LOCAL" in c or "NOMBRE" in c:
                    col_nombre = j
                elif "PROVEEDOR" in c or "SERVICIO" in c:
                    col_proveedor = j
            if col_cc is not None:
                header_idx = i
                break

    print(f"    Header at row {header_idx}: cc={col_cc} empresa={col_empresa} marca={col_marca} nombre={col_nombre} proveedor={col_proveedor}")

    for row in rows[header_idx + 1:]:
        if not row or all(c is None for c in row):
            continue
        row_str = " ".join(str(c).upper() for c in row if c)

        # Check if this is a PREVIFUEGO row
        if "PREVIFUEGO" not in row_str:
            continue

        # Get CC value
        cc_val = None
        if col_cc is not None and col_cc < len(row):
            cc_val = row[col_cc]
        else:
            # Try first non-None cell
            for c in row:
                if c and str(c).strip():
                    cc_val = c
                    break

        if not cc_val:
            continue

        ckey = canonical(cc_val)
        if not ckey:
            continue

        # Get empresa and marca
        empresa = ""
        marca   = ""
        nombre  = str(cc_val).strip().upper()

        if col_empresa is not None and col_empresa < len(row) and row[col_empresa]:
            empresa = str(row[col_empresa]).strip()
        if col_marca is not None and col_marca < len(row) and row[col_marca]:
            marca = str(row[col_marca]).strip()
        # Also try extracting from CC name (e.g. "KFC 05 ALBORADA" contains brand implicitly)
        if not marca:
            # Detect brand from prefix
            pfx = ckey[0]
            marca = pfx  # will be overridden below if found in presupuesto text

        presupuesto_map[ckey] = {
            "empresa":    empresa,
            "marca":      marca,
            "cc_pres":    str(cc_val).strip(),
        }
        presupuesto_raw.append((ckey, str(cc_val).strip(), empresa, marca))

    print(f"    Found {sum(1 for k in presupuesto_map)} PREVIFUEGO entries so far")

print(f"\nTotal PRESUPUESTO entries: {len(presupuesto_map)}")

# Show a sample
print("Sample entries:")
for k, v in list(presupuesto_map.items())[:10]:
    print(f"  {k} → {v}")

# Show what brands appear
from collections import Counter
brand_counter = Counter(v['marca'] for v in presupuesto_map.values())
print("\nBrands in PRESUPUESTO:")
for b, c in sorted(brand_counter.items(), key=lambda x: -x[1]):
    print(f"  {b!r}: {c}")

empresa_counter = Counter(v['empresa'] for v in presupuesto_map.values())
print("\nEmpresas in PRESUPUESTO:")
for e, c in sorted(empresa_counter.items(), key=lambda x: -x[1]):
    print(f"  {e!r}: {c}")


# ── 2. Parse MATRIZ – main extintor rows ──────────────────────────────────────
# Sheets are month names. Columns: CC | UBICACION | TIPO | CAPACIDAD | ... | AÑO_ULT | AÑO_PROX

print("\n=== Parsing MATRIZ ===")

SKIP_ROW = re.compile(
    r'(^TOTAL|^SUBTOTAL|^LOCALES|OFICINA|GRUPO KFC)',
    re.IGNORECASE
)

def is_skip_ubicacion(s):
    if not s: return True
    su = s.strip().upper()
    if SKIP_ROW.search(su): return True
    # rows with just numbers or very short
    if re.match(r'^\d+$', su): return True
    return False


def parse_sheet_extintor_rows(ws, mes):
    """Return list of extintor dicts from a MATRIZ sheet."""
    all_rows = list(ws.iter_rows(values_only=True))
    # Find header row
    header_idx = 0
    for i, row in enumerate(all_rows[:8]):
        cells = [str(c).strip().upper() if c else "" for c in row]
        # Header usually has CC, TIPO, CAPACIDAD
        if sum(1 for c in cells if c in ("CC","TIPO","CAPACIDAD","UBICACION","UBICACIÓN")) >= 2:
            header_idx = i
            break
        # Also accept row where first cell looks like a header
        if any("CC" == c for c in cells):
            header_idx = i
            break

    # Detect column positions from header
    hrow = [str(c).strip().upper() if c else "" for c in all_rows[header_idx]]
    def find_col(*names):
        for name in names:
            for j, h in enumerate(hrow):
                if name in h:
                    return j
        return None

    c_cc   = find_col("CC") or 0
    c_ubic = find_col("UBICACION","UBICACIÓN","NOMBRE","LOCAL") or 1
    c_tipo = find_col("TIPO") or 2
    c_cap  = find_col("CAPACIDAD","CAP") or 3
    c_ult  = find_col("ULT","ÚLTIMA","ULTIMA","ANTERIOR")
    c_prox = find_col("PRÓX","PROX","SIGUIENTE","NEXT")

    # If not found by name, use fixed positions from known format
    if c_ult is None:  c_ult  = 9
    if c_prox is None: c_prox = 10

    rows_out = []
    current_cc_raw = None
    current_ckey   = None

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

        # Forward-fill CC (merged cells)
        if cc_raw is not None and str(cc_raw).strip():
            raw_str = str(cc_raw).strip()
            ckey = canonical(raw_str)
            if ckey:
                current_cc_raw = raw_str
                current_ckey   = ckey

        if not current_ckey:
            continue

        tipo, cap = norm_tipo_cap(tipo_r, cap_r)

        # Skip rows without valid extintor type
        if not tipo:
            continue
        if tipo not in ("PQS", "CO2", "TIPO K"):
            # Skip non-extintor rows (e.g. "MANTENIMIENTO SIST CO2")
            continue

        # Skip header-looking ubicaciones
        if is_skip_ubicacion(ubic):
            ubic = ""

        año_ult  = to_year(ult_r)
        año_prox = to_year(prox_r)

        rows_out.append({
            "ckey":      current_ckey,
            "cc_raw":    current_cc_raw,
            "mes":       mes,
            "ubic":      ubic,
            "tipo":      tipo,
            "cap":       cap,
            "año_ult":   año_ult,
            "año_prox":  año_prox,
        })

    return rows_out


matriz_rows = []
for sheet_name in wb_mat.sheetnames:
    mes = sheet_name.strip().upper()
    if mes not in MONTHS:
        print(f"  Skipping non-month sheet: {sheet_name!r}")
        continue
    ws = wb_mat[sheet_name]
    rows = parse_sheet_extintor_rows(ws, mes)
    print(f"  Sheet '{sheet_name}': {len(rows)} extintor rows")
    matriz_rows.extend(rows)

print(f"\nTotal MATRIZ rows: {len(matriz_rows)}")
unique_keys = set(r["ckey"] for r in matriz_rows)
print(f"Unique locals in MATRIZ: {len(unique_keys)}")


# ── 3. Parse PROYECCIÓN – get dates and extintor data ─────────────────────────
# Sheets are month names. Similar structure to MATRIZ.

print("\n=== Parsing PROYECCIÓN ===")

proy_rows = []
proy_local_info = {}   # ckey -> {mes, cc_raw}

for sheet_name in wb_proy.sheetnames:
    mes = sheet_name.strip().upper()
    if mes not in MONTHS:
        print(f"  Skipping non-month sheet: {sheet_name!r}")
        continue
    ws = wb_proy[sheet_name]
    rows = parse_sheet_extintor_rows(ws, mes)
    print(f"  Sheet '{sheet_name}': {len(rows)} extintor rows")
    proy_rows.extend(rows)
    # Record first-seen month per local
    for r in rows:
        if r["ckey"] not in proy_local_info:
            proy_local_info[r["ckey"]] = {"mes": r["mes"], "cc_raw": r["cc_raw"]}

print(f"\nTotal PROYECCIÓN rows: {len(proy_rows)}")
proy_keys = set(r["ckey"] for r in proy_rows)
print(f"Unique locals in PROYECCIÓN: {len(proy_keys)}")

# Locals in PROYECCIÓN but not in MATRIZ
only_proy = proy_keys - unique_keys
print(f"\nLocals only in PROYECCIÓN (not MATRIZ): {len(only_proy)}")
for k in sorted(only_proy):
    info = proy_local_info.get(k, {})
    print(f"  {fmt_code(*k)} – MES: {info.get('mes','?')} – CC: {info.get('cc_raw','?')}")


# ── 4. Build H068 extintores (to merge into K172) ────────────────────────────

H068_KEY = ('H', 68) if canonical("H068") == ('H', 68) else None
# Try to find H068 in MATRIZ
h068_ckey = None
for k in unique_keys:
    if k[1] == 68 and k[0] in ('H', 'K'):
        print(f"  Candidate for H068: {k}")
# Actually let's just search directly
h068_rows = [r for r in matriz_rows if r["ckey"] == ('H', 68)]
if not h068_rows:
    # Try alternate: sometimes "HELADERIAS" might be coded differently
    h068_rows = [r for r in matriz_rows if r["cc_raw"] and "068" in r["cc_raw"].upper() and "H" in r["cc_raw"].upper()[:3]]
print(f"\nH068 rows found in MATRIZ: {len(h068_rows)}")
for r in h068_rows:
    print(f"  {r}")


# ── 5. Apply corrections and build final dataset ──────────────────────────────

K172_KEY = ('K', 172)
G042_KEY = ('G', 42)
K079_KEY = ('K', 79)
K192_KEY = ('K', 192)
K194_KEY = ('K', 194)

# Remove G042 (closed)
matriz_rows = [r for r in matriz_rows if r["ckey"] != G042_KEY]
proy_rows   = [r for r in proy_rows   if r["ckey"] != G042_KEY]
print(f"\nAfter removing G042: {len(matriz_rows)} MATRIZ rows")

# Merge H068 into K172: add H068's extintores to K172 with same MES as K172
k172_info = next((r for r in matriz_rows if r["ckey"] == K172_KEY), None)
k172_mes  = k172_info["mes"] if k172_info else "JULIO"

for r in h068_rows:
    merged = dict(r)
    merged["ckey"]   = K172_KEY
    merged["cc_raw"] = f"K172 (ex-H068)"
    merged["mes"]    = k172_mes
    matriz_rows.append(merged)
    print(f"  Merged H068 extintor → K172: {merged['tipo']} {merged['cap']}")

# Remove original H068 entries (already done since we only added merges)
matriz_rows = [r for r in matriz_rows if r["ckey"] != ('H', 68)]

# Add K079, K192, K194 from PROYECCIÓN if not already in MATRIZ
for target_key in [K079_KEY, K192_KEY, K194_KEY]:
    existing = [r for r in matriz_rows if r["ckey"] == target_key]
    if not existing:
        proy_found = [r for r in proy_rows if r["ckey"] == target_key]
        # Deduplicate PROYECCIÓN rows (each extintor may appear twice)
        seen = set()
        deduped = []
        for r in proy_found:
            sig = (r["tipo"], r["cap"], r["ubic"])
            if sig not in seen:
                seen.add(sig)
                deduped.append(r)
        print(f"\nAdding {target_key} from PROYECCIÓN: {len(deduped)} extintores (deduped from {len(proy_found)})")
        for r in deduped:
            print(f"  {r['tipo']} {r['cap']} – {r['ubic']}")
        matriz_rows.extend(deduped)
    else:
        print(f"\n{target_key} already in MATRIZ with {len(existing)} rows")

# CN04=BS04, CN16=BS16, CN37=BS37, CN31=BS31 (same location, CINNABON+BASKIN)
# Keep only one entry per location. BS entries should map to CN (or vice versa).
# Per user: CN prefix = CINNABON. BS prefix = BASKIN ROBBINS.
# These pairs share the same physical location but are different brands.
# We should keep BOTH as separate locals (they have separate brand names).
# The issue was that BS031 was listed as CN031 in PRESUPUESTO.
# Solution: for BS locals that match CN locals, they are still separate in our DB.
# Just ensure BS031 maps correctly to BASKIN ROBBINS.

# However, user said "CN04 y BS04 son lo mismo" → same physical location, separate entries.
# We keep both. No de-duplication needed between CN and BS.


# ── 6. Determine cycle and COBRO for each local ───────────────────────────────
# Cycle: OCT 2026 – SEP 2027 = RECARGA for ALL locals
# Sheets OCT/NOV/DEC → recarga month 2026 (within cycle)
# Sheets ENE-SEP → recarga month 2027 (within cycle)
# COBRO_ANUAL = RECARGA price (which includes mantt) if in cycle
#             = MANTT price only if not in cycle (but ALL are in cycle this year!)

OCT_TO_DEC = {"OCTUBRE", "NOVIEMBRE", "DICIEMBRE"}

def get_recarga_year(mes):
    """Return the year of recarga in the current cycle."""
    if mes in OCT_TO_DEC:
        return 2026
    return 2027

def get_ult_recarga_year(mes):
    """Last recarga was 3 years before current cycle."""
    return get_recarga_year(mes) - 3


# ── 7. Determine brand name from PRESUPUESTO map or fallback ─────────────────

BRAND_FALLBACK = {
    'K':  ("INT FOOD SERVICES CORP SA", "KFC"),
    'G':  ("INT FOOD SERVICES CORP SA", "KFC"),
    'M':  ("SHEMLON SA", "MENESTRAS DEL NEGRO"),
    'T':  ("DELI-INTERNACIONAL SA", "TROPIBURGER"),
    'V':  ("PROMOTORA ECUATORIANA DE CAFÉ DE COLOMBIA SA", "JUAN VALDEZ"),
    'I':  ("PRODUCCIONES Y EVENTOS NOVOEVENTOS SA", "IL CAPPO"),
    'J':  ("GRUPO KFC", "CAJUN"),
    'BS': ("SHEMLON SA", "BASKIN ROBBINS"),
    'CN': ("SHEMLON SA", "CINNABON"),
    'DI': ("GRUPO KFC", "DOLCE INCONTRO"),
    'CA': ("GRUPO KFC", "AMERICAN DELI"),
    'F':  ("GRUPO KFC", "SPORT PLANET"),
}

def get_brand(ckey):
    pinfo = presupuesto_map.get(ckey)
    if pinfo and pinfo.get("empresa") and pinfo.get("marca"):
        return pinfo["empresa"], pinfo["marca"]
    # Fallback to prefix map
    pfx = ckey[0]
    return BRAND_FALLBACK.get(pfx, BRAND_FALLBACK.get(pfx[:1], ("DESCONOCIDO", "DESCONOCIDO")))


# ── 8. Assign MES from PROYECCIÓN if not in MATRIZ ───────────────────────────
# (already done by adding proy rows above)

# Verify/override MES from PROYECCIÓN for all locals where PROYECCIÓN has data
# The PROYECCIÓN is the source of truth for dates.
# Build a map: ckey -> mes from PROYECCIÓN
proy_mes_map = {}
for r in proy_rows:
    k = r["ckey"]
    if k not in proy_mes_map:
        proy_mes_map[k] = r["mes"]

# For locals in MATRIZ, if PROYECCIÓN has a different month, we could override.
# Per user: PROYECCIÓN is source of truth for dates.
# Apply override only where there's a mismatch and we trust PROYECCIÓN more.
mes_overrides = {}
for ckey, proy_mes in proy_mes_map.items():
    mat_rows = [r for r in matriz_rows if r["ckey"] == ckey]
    if mat_rows:
        mat_mes = mat_rows[0]["mes"]
        if mat_mes != proy_mes:
            mes_overrides[ckey] = (mat_mes, proy_mes)

if mes_overrides:
    print(f"\nMES overrides from PROYECCIÓN (MATRIZ→PROYECCIÓN):")
    for k, (old, new) in list(mes_overrides.items())[:20]:
        print(f"  {fmt_code(*k)}: {old} → {new}")
    # Apply overrides
    for r in matriz_rows:
        if r["ckey"] in mes_overrides:
            r["mes"] = mes_overrides[r["ckey"]][1]


# ── 9. Build final rows with pricing ─────────────────────────────────────────

print("\n=== Building final rows ===")

# Group MATRIZ rows by (ckey, mes) to handle unique-per-extintor
# Use (ckey, mes, tipo, cap, ubic) as de-dup key
seen_extintores = set()
final_rows = []

for r in matriz_rows:
    ckey   = r["ckey"]
    mes    = r["mes"]
    tipo   = r["tipo"]
    cap    = r["cap"]
    ubic   = r.get("ubic", "")

    price_key = (tipo, cap)
    cm = MANTT.get(price_key)
    cr = RECARGA.get(price_key)
    cap_disp = CAP_DISPLAY.get(price_key, f"{cap}")

    # Cobro logic: ALL locals are in OCT2026-SEP2027 cycle
    # RECARGA includes mantt → COBRO = RECARGA price
    cobro = cr if cr is not None else cm
    cobro_tipo = "RECARGA (mantt incluido)" if cr is not None else "MANTT"

    año_rec     = get_recarga_year(mes)
    año_ult_rec = get_ult_recarga_year(mes)

    empresa, marca = get_brand(ckey)
    codigo = fmt_code(*ckey)

    año_prox_mat = r.get("año_prox", "")
    año_ult_mat  = r.get("año_ult", "")
    notas = []
    if cm is None:
        notas.append(f"precio no hallado ({tipo}/{cap})")
    # Override years with calculated values (source of truth = cycle logic)
    # but keep MATRIZ values as reference

    final_rows.append({
        "CÓDIGO":              codigo,
        "CC_ORIGINAL":         r.get("cc_raw", ""),
        "EMPRESA":             empresa,
        "MARCA":               marca,
        "MES_SERVICIO":        mes,
        "UBICACIÓN":           ubic,
        "TIPO":                tipo,
        "CAPACIDAD":           cap_disp,
        "COSTO_MANTT":         round(cm, 2) if cm else None,
        "PRECIO_RECARGA":      round(cr, 2) if cr else None,
        "COBRO_ANUAL":         round(cobro, 2) if cobro else None,
        "TIPO_COBRO":          cobro_tipo,
        "AÑO_RECARGA":         año_rec,
        "AÑO_ULT_RECARGA":     año_ult_rec,
        "AÑO_PROX_MAT":        año_prox_mat,
        "AÑO_ULT_MAT":         año_ult_mat,
        "NOTAS":               "; ".join(notas),
        "_cm":                 cm or 0.0,
        "_cr":                 cr or 0.0,
        "_cobro":              cobro or 0.0,
        "_ckey":               ckey,
    })

print(f"Total final extintor rows: {len(final_rows)}")
unique_locals = set(r["_ckey"] for r in final_rows)
print(f"Unique locals:             {len(unique_locals)}")

# Check unique local count by ckey
local_count_by_code = Counter(r["CÓDIGO"] for r in final_rows)
print(f"Unique CÓDIGO values:      {len(set(r['CÓDIGO'] for r in final_rows))}")


# ── 10. Per-local summary ─────────────────────────────────────────────────────

local_summary = {}
for r in final_rows:
    k = r["CÓDIGO"]
    if k not in local_summary:
        local_summary[k] = {
            "CÓDIGO":        k,
            "CC_ORIGINAL":   r["CC_ORIGINAL"],
            "EMPRESA":       r["EMPRESA"],
            "MARCA":         r["MARCA"],
            "MES_SERVICIO":  r["MES_SERVICIO"],
            "N_EXTINTORES":  0,
            "TOTAL_MANTT":   0.0,
            "TOTAL_RECARGA": 0.0,
            "COBRO_ANUAL":   0.0,
            "AÑO_RECARGA":   r["AÑO_RECARGA"],
            "AÑO_ULT_RECARGA": r["AÑO_ULT_RECARGA"],
            "DESGLOSE":      [],
        }
    s = local_summary[k]
    s["N_EXTINTORES"]  += 1
    s["TOTAL_MANTT"]   += r["_cm"]
    s["TOTAL_RECARGA"] += r["_cr"]
    s["COBRO_ANUAL"]   += r["_cobro"]
    s["DESGLOSE"].append(f"{r['TIPO']} {r['CAPACIDAD']} ${r['_cobro']:.2f}")

for s in local_summary.values():
    s["TOTAL_MANTT"]   = round(s["TOTAL_MANTT"],   2)
    s["TOTAL_RECARGA"] = round(s["TOTAL_RECARGA"],  2)
    s["COBRO_ANUAL"]   = round(s["COBRO_ANUAL"],   2)
    s["FORMULA_COBRO"] = " + ".join(s["DESGLOSE"])

print(f"\nLocales en resumen: {len(local_summary)}")

# Show count by month
month_counts = Counter(s["MES_SERVICIO"] for s in local_summary.values())
print("\nLocales por mes:")
for m in MONTHS:
    print(f"  {m}: {month_counts.get(m, 0)}")

# Show total annual billing
total_anual = sum(s["COBRO_ANUAL"] for s in local_summary.values())
print(f"\nCOBRO ANUAL TOTAL: ${total_anual:,.2f}")

# Auditoría: locales en PRESUPUESTO pero no en MATRIZ/final
pres_keys = set(presupuesto_map.keys())
final_keys = set(r["_ckey"] for r in final_rows)
only_pres = pres_keys - final_keys
only_final = final_keys - pres_keys

print(f"\nEn PRESUPUESTO pero no en MATRIZ final: {len(only_pres)}")
for k in sorted(only_pres):
    info = presupuesto_map[k]
    print(f"  {fmt_code(*k)} ({info.get('cc_pres','?')})")

print(f"\nEn MATRIZ final pero no en PRESUPUESTO: {len(only_final)}")
for k in sorted(only_final):
    print(f"  {fmt_code(*k)}")


# ── 11. Write Excel ───────────────────────────────────────────────────────────

OUT = "/home/user/previfuego-facturacion/BASE_DATOS_KFC.xlsx"
wb = openpyxl.Workbook()

H_FILL  = PatternFill("solid", fgColor="1F4E79")
H_FONT  = Font(color="FFFFFF", bold=True)
H2_FILL = PatternFill("solid", fgColor="375623")
ALT     = PatternFill("solid", fgColor="D6E4F0")
ALT2    = PatternFill("solid", fgColor="E2EFDA")
CENTER  = Alignment(horizontal="center")


def auto_width(ws):
    for col in ws.columns:
        mx = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(mx + 2, 50)


# ── Sheet 1: DETALLE ─────────────────────────────────────────────────────────
ws_det = wb.active
ws_det.title = "DETALLE"

COLS_DET = [
    "CÓDIGO","CC_ORIGINAL","EMPRESA","MARCA","MES_SERVICIO","UBICACIÓN",
    "TIPO","CAPACIDAD","COSTO_MANTT","PRECIO_RECARGA","COBRO_ANUAL","TIPO_COBRO",
    "AÑO_RECARGA","AÑO_ULT_RECARGA","NOTAS"
]

ws_det.append(COLS_DET)
for cell in ws_det[1]:
    cell.fill = H_FILL; cell.font = H_FONT; cell.alignment = CENTER

# Sort by month then code
month_sort = {m: i for i, m in enumerate(MONTHS)}
sorted_rows = sorted(final_rows, key=lambda r: (
    month_sort.get(r["MES_SERVICIO"], 99),
    r["CÓDIGO"]
))

for idx, r in enumerate(sorted_rows, 2):
    ws_det.append([r.get(c) for c in COLS_DET])
    if idx % 2 == 0:
        for cell in ws_det[idx]:
            cell.fill = ALT

auto_width(ws_det)


# ── Sheet 2: RESUMEN_LOCALES ─────────────────────────────────────────────────
ws_res = wb.create_sheet("RESUMEN_LOCALES")

COLS_RES = [
    "CÓDIGO","CC_ORIGINAL","EMPRESA","MARCA","MES_SERVICIO","N_EXTINTORES",
    "TOTAL_MANTT","TOTAL_RECARGA","COBRO_ANUAL","AÑO_RECARGA","AÑO_ULT_RECARGA","FORMULA_COBRO"
]

ws_res.append(COLS_RES)
for cell in ws_res[1]:
    cell.fill = H2_FILL; cell.font = H_FONT; cell.alignment = CENTER

sorted_summaries = sorted(local_summary.values(), key=lambda s: (
    month_sort.get(s["MES_SERVICIO"], 99),
    s["CÓDIGO"]
))

for idx, s in enumerate(sorted_summaries, 2):
    ws_res.append([s.get(c) for c in COLS_RES])
    if idx % 2 == 0:
        for cell in ws_res[idx]:
            cell.fill = ALT2

auto_width(ws_res)


# ── Sheet 3: AUDITORÍA ───────────────────────────────────────────────────────
ws_aud = wb.create_sheet("AUDITORÍA")
ws_aud.append(["TIPO","CÓDIGO","DETALLE"])
for cell in ws_aud[1]:
    cell.fill = PatternFill("solid", fgColor="7B0000")
    cell.font = Font(color="FFFFFF", bold=True)

for k in sorted(only_pres):
    info = presupuesto_map.get(k, {})
    ws_aud.append(["EN_PRESUPUESTO_NO_EN_MATRIZ", fmt_code(*k),
                   f"CC presup: {info.get('cc_pres','?')} | empresa: {info.get('empresa','?')}"])

for k in sorted(only_final):
    ws_aud.append(["EN_MATRIZ_NO_EN_PRESUPUESTO", fmt_code(*k), ""])

for r in final_rows:
    if r["NOTAS"]:
        ws_aud.append(["PRECIO_NO_HALLADO", r["CÓDIGO"],
                       f"{r['TIPO']} {r['CAPACIDAD']} – {r['NOTAS']}"])

auto_width(ws_aud)


wb.save(OUT)
print(f"\n✓ Saved: {OUT}")
print(f"  DETALLE rows: {len(final_rows)}")
print(f"  Locales:      {len(local_summary)}")
print(f"  Cobro anual:  ${total_anual:,.2f}")
