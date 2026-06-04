#!/usr/bin/env python3
"""
build_database.py
Downloads MATRIZ EXTINTORES GRUPO KFC.xlsx from Dropbox and builds BASE_DATOS_KFC.xlsx.
"""

import requests
import io
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from collections import defaultdict

# ── 1. Download from Dropbox ─────────────────────────────────────────────────
DROPBOX_TOKEN = (
    "sl.u.AGgF5XKR5BwOCmo95YjLwdRE-buSMDIWiOPYXcWcE7SUTTaHr9AIpXz-9MVh42IhLVZsHJCrKnkxOwzu8i5hgLrIuJC8XlRGaCprnDHwr1Czw7hxc92BXwjNMEAMvR0iXS88gLeyV_v4Wbj4KCTCMX7qlmtztJWZSS1oN0mfnHX2dj2SP2A_4Ari0pVrlMqbUbnWlRGYRNgHrMA7K6bbUCfeBq-3IO8VaTTkZEnyYNMKPHUHu1CHyxypSrC44D5jlBEyPLguvyKppj_qwNQkfJSHHhLv-VvqohTPLRMewH0inx-UpIcq7ku28sGrd04g2-R2YJUOcEbAWOgu5Uwm2VCXXwTxTlbBPV2NQe--n_ZT_uzACDQHLV0E9c0FIA9Z8_v56DGJch7mMmR7TbavjM9v2yFDzf8Jdw4ip3LCYS1yhkHwV-_CTZ8fFdTcJJ5gm4CbjCGhI-k9eGvmp4xFeId6Wra8yWLA_DZd-Q11gHbb-FY7vCD7hWvNi0B_3N86FG4S1ha_0loOmJkav9dk6L81RhIcAzLiMsN64rO0hTS4gugUGT8MuXm9UbAwNL_7untY5spa_8ecfefeC894o-Ae3TVeAji3dvcrUYsgClYgK1fkylqt1SaeOj1K0PZFcjmxin-K9BUvfGTjUclowh-HxEDl7n6-QofS1DqczB0loetZV27LHZFoBGl4FDRpSDhydxE9XkwyoIXg6LSfUUNCwi9fFlBnISrO5eZaGrCjpqpftgEug_jPjRLSILFra6uakJOVT8gS-aZoWBZv9YacSnrJ6iuRf0mRo6X7nV9SOm2cmdtyUZvdxBxtjTfaQfFWD8VkHHmTHLR-Fmt50on5DDaocHhfzkJt8wyafU4YH7NSbz2mhfFmxcb1s5WEQrgslgW5dKmmi6VGBIcPDMeIPxoRZe43aodjcWYzZmOT7czMjPpI5J6CNalK8HZMI43oRVwKAml3TXMEg0CLhtlMRdsafEo3rIdoB81ZK3RbnnGTCXkXuXKDdzd1qHRyE3Z2DCbcHFYvwzYcT6gr4VnIc4WmsTt7f4s4bGoaHoJxD8K1zOn85J_VLHJBOXdknyDSyHxhNPSKebKo3_pAFXJe4Q4HoaCrGyxHH09xGS_1TteDVPnHqRUFK5fIa1EezABLZfDKnvKOBxLeaTe9AEO89ddAiYvpJaGkpCKKG0GIOtivpsNRwi4tczW06bmZZAUPLQFjkWhTcxvh_7QHYG_Jm63xC1B7st_7Zkez4HnAzUsTCEucTpIpTWvwAcEuu-5cgRUdbe2XHjeSeeFUUmPB"
)
DROPBOX_PATH = "/Previfuego/MATRIZ LOCALES/MATRIZ EXTINTORES GRUPO KFC.xlsx"

print("Downloading file from Dropbox...")
resp = requests.post(
    "https://content.dropboxapi.com/2/files/download",
    headers={
        "Authorization": f"Bearer {DROPBOX_TOKEN}",
        "Dropbox-API-Arg": f'{{"path": "{DROPBOX_PATH}"}}',
    },
    timeout=60,
)
resp.raise_for_status()
print(f"Downloaded {len(resp.content):,} bytes")

wb_src = openpyxl.load_workbook(io.BytesIO(resp.content), data_only=True)
print(f"Sheets: {wb_src.sheetnames}")

# ── 2. Price lists ─────────────────────────────────────────────────────────────
# Key: (tipo_normalized, capacidad_num_str)
# capacidad_num_str: "5", "10", "20", "50", "75", "2.5"
MANTT_PRICES = {
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

RECARGA_PRICES = {
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

# Display capacidad with units
CAPACIDAD_DISPLAY = {
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

MONTHS = [
    "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
    "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"
]

# ── 3. Normalize tipo ─────────────────────────────────────────────────────────
def normalize_tipo(raw):
    if raw is None:
        return ""
    s = str(raw).strip().upper()
    if s in ("K", "TIPO K", "TIPO  K", "TYPE K"):
        return "TIPO K"
    if s.startswith("CO2"):
        return "CO2"
    if s.startswith("PQS"):
        return "PQS"
    return s

def normalize_capacidad_num(raw):
    """Return the numeric string only: '5', '10', '20', '50', '75', '2.5'"""
    if raw is None:
        return ""
    if isinstance(raw, (int, float)):
        v = float(raw)
        if v == int(v):
            return str(int(v))
        return str(round(v, 1))
    s = str(raw).strip().upper()
    # strip unit letters and noise
    s = re.sub(r'[A-Z\s]+', '', s)   # remove all letters/spaces, keep digits/dots/commas
    s = s.replace(",", ".").replace("..", ".")  # fix "2..5" -> "2.5"
    # take only numeric part
    m = re.match(r'(\d+(?:\.\d+)?)', s)
    if m:
        try:
            v = float(m.group(1))
            if v == int(v):
                return str(int(v))
            return str(round(v, 1))
        except:
            pass
    return s

# ── 4. Normalize CC code ──────────────────────────────────────────────────────
# CC examples seen:
#   "KFC 02 CHIMBORAZO", "KFC 25 SHELL DURAN", "KFC 105 RIOCENTRO NORTE"
#   "M31 RIOCENTRO NORTE", "M35 GARZOTA", "M58 PEÑAS"
#   "T32 TERMINAL"
#   "I03 MALL DEL SOL"
#   "J01 MALL DEL SOL"   -> Juan Valdez
#   "K002EC", "M035EC", "T015EC"
#   "GUS49"

SKIP_CC_PATTERNS = re.compile(
    r'^(LOCALES|TOTAL|MENESTRAS|JUAN VALDEZ|DELI|ESPAÑOL|IL CAPPO|KFC$|IVORY$|BAGUET|TROPI$|OFICINAS GRUPO)',
    re.IGNORECASE
)

# Additional brand prefixes seen in the data:
# JV -> V (Juan Valdez)
# J  -> V (Juan Valdez, when just "J 08 TERMINAL")
# BS -> BS (Bagueterie)
# CN -> CN (Cajun)
# I  -> I  (Ivory, "I 04 MALL DEL SUR")
# T  -> T  (Tropiburger, "T 43 FORTIN")
# M  -> M  (Menestras, "M 14 RIOCENTRO SUR")
# V  -> V  (Juan Valdez, "V 55 MALL DEL SOL", "V 03 SAN MARINO")
# E  -> unknown/other brand - keep as-is
# A  -> unknown/other brand
# R  -> unknown/other brand
# DI -> unknown
# CA -> unknown
# F  -> unknown

def normalize_cc(raw):
    """Returns (code, prefix) e.g. ('K002', 'K')"""
    if raw is None:
        return None, None
    s = str(raw).strip()
    su = s.upper()

    # Skip header/label rows
    if SKIP_CC_PATTERNS.match(su):
        return None, None

    # Pattern: letter code + EC suffix  e.g. "K002EC", "M035EC"
    m = re.match(r'^([A-Z]+)(\d+)EC$', su)
    if m:
        prefix_raw = m.group(1)
        num = int(m.group(2))
        prefix = map_prefix(prefix_raw)
        return f"{prefix}{num:03d}", prefix

    # "KFC 02 CHIMBORAZO" / "KFC 105 ..."
    m = re.match(r'^KFC\s+(\d+)', su)
    if m:
        return f"K{int(m.group(1)):03d}", "K"

    # "KFCK124 LOJA", "KFCK69 LOJA" -> strip KFCK prefix -> K124, K069
    m = re.match(r'^KFCK(\d+)', su)
    if m:
        return f"K{int(m.group(1)):03d}", "K"

    # "GUS49" or "GUS 49"
    m = re.match(r'^GUS\s*(\d+)', su)
    if m:
        return f"G{int(m.group(1)):03d}", "G"

    # "JV15 ..." "JV75 ..." (Juan Valdez with JV prefix)
    m = re.match(r'^JV\s*(\d+)', su)
    if m:
        return f"V{int(m.group(1)):03d}", "V"

    # "BS 04 MALL DEL SOL" or "BS04 ..."
    m = re.match(r'^BS\s*(\d+)', su)
    if m:
        return f"BS{int(m.group(1)):03d}", "BS"

    # "CN 34 RIOCENTRO SUR" or "CJNC31 ..."
    m = re.match(r'^(?:CN|CJNC)\s*(\d+)', su)
    if m:
        return f"CN{int(m.group(1)):03d}", "CN"

    # Single letter + space + number patterns:
    # "V 55 ...", "V 03 ...", "M 14 ...", "M 28 ...", "T 43 ...", "T 51 ...",
    # "I 04 ...", "I 17 ...", "J 06 ...", "J 08 ..." (J=Juan Valdez)
    for pfx_raw, pfx_norm in [
        ("V", "V"), ("M", "M"), ("T", "T"), ("I", "I"), ("J", "V"),
        ("K", "K"), ("G", "G"),
    ]:
        m = re.match(rf'^{pfx_raw}\s+(\d+)', su)
        if m:
            return f"{pfx_norm}{int(m.group(1)):03d}", pfx_norm

    # "M31 ..." "M35 ..." (Menestras no space)
    m = re.match(r'^M(\d+)\b', su)
    if m:
        return f"M{int(m.group(1)):03d}", "M"

    # "T32 ..." (Tropiburger no space)
    m = re.match(r'^T(\d+)\b', su)
    if m:
        return f"T{int(m.group(1)):03d}", "T"

    # "I03 ..." (Ivory no space)
    m = re.match(r'^I(\d+)\b', su)
    if m:
        return f"I{int(m.group(1)):03d}", "I"

    # "J01 ..." (Juan Valdez no space) -> V prefix
    m = re.match(r'^J(\d+)\b', su)
    if m:
        return f"V{int(m.group(1)):03d}", "V"

    # Generic: letter(s) + optional space + digits  (covers E06, A37, R08, DI01, CA01, F003, etc.)
    m = re.match(r'^([A-Z]{1,4})\s*(\d+)', su)
    if m:
        prefix_raw = m.group(1)
        num = int(m.group(2))
        prefix = map_prefix(prefix_raw)
        return f"{prefix_raw}{num:03d}", prefix

    return s, s[:1]

PREFIX_MAP = {
    "K": "K", "KFC": "K",
    "G": "G", "GUS": "G",
    "M": "M",
    "T": "T", "TROP": "T", "TRO": "T",
    "V": "V", "JV": "V", "J": "V",
    "I": "I", "ILCI": "I", "ILC": "I",
    "BS": "BS",
    "CN": "CN", "CJNC": "CN",
    # Unknown brands - preserve raw prefix
    "E": "E", "A": "A", "R": "R", "DI": "DI", "CA": "CA", "F": "F",
}

def map_prefix(p):
    return PREFIX_MAP.get(p.upper(), p[0].upper() if p else "?")

COMPANY_MAP = {
    "K":  ("INT FOOD SERVICES CORP SA", "KFC"),
    "G":  ("INT FOOD SERVICES CORP SA", "KFC"),
    "M":  ("SHEMLON SA", "MENESTRAS DEL SEÑOR"),
    "T":  ("DELI-INTERNACIONAL SA", "TROPIBURGER"),
    "V":  ("PROMOTORA ECUATORIANA DE CAFÉ DE COLOMBIA SA", "JUAN VALDEZ"),
    "I":  ("PRODUCCIONES Y EVENTOS NOVOEVENTOS SA", "IVORY"),
    "BS": ("SHEMLON SA", "BAGUETERIE"),
    "CN": ("SHEMLON SA", "CAJUN"),
}

def get_company(prefix):
    if not prefix:
        return ("DESCONOCIDO", "DESCONOCIDO")
    if prefix in COMPANY_MAP:
        return COMPANY_MAP[prefix]
    p1 = prefix[:1].upper()
    return COMPANY_MAP.get(p1, ("DESCONOCIDO", "DESCONOCIDO"))

# ── 5. Parse all sheets ───────────────────────────────────────────────────────
COL_CC         = 0
COL_UBICACION  = 1
COL_TIPO       = 2
COL_CAPACIDAD  = 3
COL_AÑO_ULTREC = 9
COL_AÑO_PROXREC= 10

rows_data = []
skipped_sheets = []
unrecognized_tipos = set()
unrecognized_caps  = set()

for sheet_name in wb_src.sheetnames:
    month_upper = sheet_name.strip().upper()
    if month_upper not in MONTHS:
        skipped_sheets.append(sheet_name)
        continue

    ws = wb_src[sheet_name]
    all_rows = list(ws.iter_rows(values_only=True))
    print(f"\nSheet '{sheet_name}': {len(all_rows)} rows")

    # Find header row
    header_row_idx = 0
    for i, row in enumerate(all_rows[:5]):
        for cell in row:
            if cell and re.search(r'\bCC\b', str(cell).upper()):
                header_row_idx = i
                break
        else:
            continue
        break

    data_start = header_row_idx + 1

    current_cc   = None
    current_code = None
    current_pfx  = None

    for row in all_rows[data_start:]:
        if not row:
            continue

        def gc(idx):
            if idx < len(row):
                v = row[idx]
                return str(v).strip() if v is not None else ""
            return ""

        cc_raw      = row[COL_CC] if COL_CC < len(row) else None
        ubicacion   = gc(COL_UBICACION)
        tipo_raw    = row[COL_TIPO] if COL_TIPO < len(row) else None
        cap_raw     = row[COL_CAPACIDAD] if COL_CAPACIDAD < len(row) else None
        año_ultrec  = row[COL_AÑO_ULTREC]  if COL_AÑO_ULTREC  < len(row) else None
        año_proxrec = row[COL_AÑO_PROXREC] if COL_AÑO_PROXREC < len(row) else None

        # Forward-fill CC (merged cells appear as None in subsequent rows)
        if cc_raw is not None and str(cc_raw).strip():
            code, pfx = normalize_cc(cc_raw)
            if code is not None:
                current_cc   = str(cc_raw).strip()
                current_code = code
                current_pfx  = pfx
            # If code is None it was a skip pattern; keep previous current_cc

        # Skip rows with no useful data
        if not ubicacion and tipo_raw is None and cap_raw is None:
            continue
        if not ubicacion and not tipo_raw:
            continue

        # Skip label / subtotal rows
        if ubicacion and re.search(
            r'(TOTAL|SUBTOTAL|LOCALES|MENESTRAS|JUAN VALDEZ|DELI|ESPAÑOL|IL CAPPO)',
            ubicacion.upper()
        ):
            continue

        tipo     = normalize_tipo(tipo_raw)
        cap_num  = normalize_capacidad_num(cap_raw)
        key      = (tipo, cap_num)
        cap_disp = CAPACIDAD_DISPLAY.get(key, f"{cap_num}")

        costo_mantt   = MANTT_PRICES.get(key)
        costo_recarga = RECARGA_PRICES.get(key)

        if costo_mantt is None and tipo:
            unrecognized_tipos.add(f"{tipo}/{cap_num}")

        # Año values
        def to_year(v):
            if v is None:
                return ""
            try:
                return str(int(float(str(v))))
            except:
                return str(v).strip()

        año_ult = to_year(año_ultrec)
        año_prox = to_year(año_proxrec)

        # Recarga_2026: due in OCT2026-SEP2027 cycle
        if año_prox == "2026" or año_prox == "":
            recarga_2026 = "YES"
        elif año_prox in ("2027", "2028", "2029", "2030"):
            recarga_2026 = "NO"
        else:
            recarga_2026 = "YES"

        empresa, marca = get_company(current_pfx)

        notas = []
        if costo_mantt is None:
            notas.append(f"precio no hallado ({tipo}/{cap_num})")
        if not current_code:
            notas.append("CC no identificado")

        costo_rec_out = (costo_recarga or 0.0) if recarga_2026 == "YES" else 0.0

        rows_data.append({
            "CÓDIGO":             current_code or "",
            "CC_ORIGINAL":        current_cc or "",
            "EMPRESA":            empresa,
            "MARCA":              marca,
            "MES_SERVICIO":       month_upper,
            "UBICACIÓN":          ubicacion,
            "TIPO":               tipo,
            "CAPACIDAD":          cap_disp,
            "COSTO_MANTT":        costo_mantt,
            "COSTO_RECARGA":      costo_rec_out,
            "RECARGA_2026":       recarga_2026,
            "AÑO_ULTIMA_RECARGA": año_ult,
            "AÑO_PROX_RECARGA":   año_prox,
            "NOTAS":              "; ".join(notas),
            "_mantt_raw":         costo_mantt or 0.0,
            "_rec_raw":           costo_rec_out,
        })

print(f"\nTotal rows parsed: {len(rows_data)}")
if skipped_sheets:
    print(f"Skipped (non-month) sheets: {skipped_sheets}")
if unrecognized_tipos:
    print(f"Unrecognized TIPO/CAP combinations: {sorted(unrecognized_tipos)}")

# ── 6. Per-local totals ───────────────────────────────────────────────────────
local_mantt   = defaultdict(float)
local_recarga = defaultdict(float)
for r in rows_data:
    k = (r["CÓDIGO"], r["MES_SERVICIO"])
    local_mantt[k]   += r["_mantt_raw"]
    local_recarga[k] += r["_rec_raw"]

for r in rows_data:
    k = (r["CÓDIGO"], r["MES_SERVICIO"])
    r["TOTAL_MANTT_LOCAL"]   = round(local_mantt[k], 2)
    r["TOTAL_RECARGA_LOCAL"] = round(local_recarga[k], 2)

# ── 7. Sample output ──────────────────────────────────────────────────────────
COLS_OUT = [
    "CÓDIGO","CC_ORIGINAL","EMPRESA","MARCA","MES_SERVICIO","UBICACIÓN",
    "TIPO","CAPACIDAD","COSTO_MANTT","COSTO_RECARGA",
    "RECARGA_2026","AÑO_ULTIMA_RECARGA","AÑO_PROX_RECARGA",
    "TOTAL_MANTT_LOCAL","TOTAL_RECARGA_LOCAL","NOTAS"
]

print("\n--- FIRST 25 ROWS SAMPLE ---")
for i, r in enumerate(rows_data[:25]):
    parts = [f"{c}={r.get(c,'')}" for c in COLS_OUT if c not in ("EMPRESA","MARCA","NOTAS")]
    print(f"  [{i+1}] " + " | ".join(parts))
    if r["NOTAS"]:
        print(f"       NOTA: {r['NOTAS']}")

# ── 8. Write Excel ────────────────────────────────────────────────────────────
OUT = "/home/user/previfuego-facturacion/BASE_DATOS_KFC.xlsx"
wb_out = openpyxl.Workbook()

ws_main = wb_out.active
ws_main.title = "DATOS"

hdr_fill = PatternFill("solid", fgColor="1F4E79")
hdr_font = Font(color="FFFFFF", bold=True)
alt_fill = PatternFill("solid", fgColor="D6E4F0")

ws_main.append(COLS_OUT)
for cell in ws_main[1]:
    cell.fill = hdr_fill
    cell.font = hdr_font
    cell.alignment = Alignment(horizontal="center")

for idx, r in enumerate(rows_data, start=2):
    ws_main.append([r.get(c) for c in COLS_OUT])
    if idx % 2 == 0:
        for cell in ws_main[idx]:
            cell.fill = alt_fill

for col in ws_main.columns:
    max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
    ws_main.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)

# ── 9. Summary sheet ─────────────────────────────────────────────────────────
summary = {}
for r in rows_data:
    k = (r["CÓDIGO"], r["MES_SERVICIO"])
    if k not in summary:
        summary[k] = {
            "CÓDIGO":        r["CÓDIGO"],
            "CC_ORIGINAL":   r["CC_ORIGINAL"],
            "EMPRESA":       r["EMPRESA"],
            "MARCA":         r["MARCA"],
            "MES_SERVICIO":  r["MES_SERVICIO"],
            "N_EXTINTORES":  0,
            "TOTAL_MANTT":   0.0,
            "TOTAL_RECARGA": 0.0,
            "RECARGA_2026":  r["RECARGA_2026"],
        }
    s = summary[k]
    s["N_EXTINTORES"] += 1
    s["TOTAL_MANTT"]  += r["_mantt_raw"]
    s["TOTAL_RECARGA"]+= r["_rec_raw"]

ws_sum = wb_out.create_sheet("RESUMEN")
sum_cols = ["CÓDIGO","CC_ORIGINAL","EMPRESA","MARCA","MES_SERVICIO",
            "N_EXTINTORES","TOTAL_MANTT","TOTAL_RECARGA","RECARGA_2026"]
ws_sum.append(sum_cols)
for cell in ws_sum[1]:
    cell.fill = PatternFill("solid", fgColor="375623")
    cell.font = Font(color="FFFFFF", bold=True)
    cell.alignment = Alignment(horizontal="center")

month_order = {m: i for i, m in enumerate(MONTHS)}
for s in sorted(summary.values(), key=lambda x: (month_order.get(x["MES_SERVICIO"], 99), x["CÓDIGO"])):
    s["TOTAL_MANTT"]   = round(s["TOTAL_MANTT"], 2)
    s["TOTAL_RECARGA"] = round(s["TOTAL_RECARGA"], 2)
    ws_sum.append([s.get(c) for c in sum_cols])

for col in ws_sum.columns:
    max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
    ws_sum.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)

wb_out.save(OUT)
print(f"\nSaved: {OUT}")

# ── 10. Final stats ───────────────────────────────────────────────────────────
unique_lm   = len(set((r["CÓDIGO"], r["MES_SERVICIO"]) for r in rows_data))
unique_codes= len(set(r["CÓDIGO"] for r in rows_data if r["CÓDIGO"]))
with_notas  = sum(1 for r in rows_data if r["NOTAS"])
rec_yes     = sum(1 for r in rows_data if r["RECARGA_2026"] == "YES")

print(f"\n=== FINAL STATS ===")
print(f"Total extintores (rows): {len(rows_data)}")
print(f"Unique local-month pairs: {unique_lm}")
print(f"Unique local codes:       {unique_codes}")
print(f"Rows with issues/notas:   {with_notas}")
print(f"Extintores RECARGA_2026=YES: {rec_yes}")

# Per-brand breakdown
brand_counts = defaultdict(int)
for r in rows_data:
    brand_counts[r["MARCA"]] += 1
print("\nRows by brand:")
for b, c in sorted(brand_counts.items()):
    print(f"  {b}: {c}")
