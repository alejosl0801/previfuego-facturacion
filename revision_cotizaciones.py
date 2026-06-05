#!/usr/bin/env python3
"""
Compara cotizaciones Dropbox vs BASE_DATOS_KFC.xlsx.
Para cada INCONSISTENCIA genera detalle de líneas (cotización vs DB).
"""

import re
import io
import json
import requests
import openpyxl
from zipfile import BadZipFile
from collections import defaultdict
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TOKEN = "sl.u.AGjIpgezlnwaiHgKQ7Bu5YYhccUjjteEDF1RqACnZjg3vFiMjL5x5d3IaLviJ9dyORvEYZT9jSRe2ocUFRy6vc_0Ls5NPIrwAB6HDxh0Rief7rixFm6VGuCmG1QG0La0ZOoKlDsxocS9ewt1m24m3MXbNcfn88LKLmsQpMw2goxO_qOqtX1yas7JTvt5mDkdplEnlMmHEdQXg3zIgeIj0lFlIzVJwaLgfZfVNL1GwShY5RwsMi_aKAm-VWi6Y5rraWcnvPDbCl2ZNy70KAT-0GNYLKlipCrcjFkCKzBvqfWX8zffKCm9A_mTqWpb6N1ZZ-MPAQICCro_UuZ9zVUvgp6sPQVlHS7vfDRiLLVZ9LIs7-h2DbCct4f4FX3fqszjgAr0LFKak5vCNv8zymvTmzTEOucEEsi5nQ4e7UNE7ixEiZ1kZ2WwsFihVJWMxhP4kFjChzHdjSJ12ubYy2JLjNKpAtLBxhVVqV6uzqoH2rkIWmEnmdpS5EjyxRGni7kZfJQZG4Bdy5IJlqJYcBtQmxE9u-UFnBe28LHnUVJB1s50kmV8yNpkvzFlqC0CtGuvLj6FgjsfsUSMOroG8gi-eZaRGDEj9jqqpMsfW9NSb7iz6budr6_5OhzxQ3w3YZRVnB70EQBdHxWZTE_8So9s-y31vhN97tE9lWqqMl0lR9VzjeBU7c5zZY3N1N6oE1Yf-yyXKGxjdC6dU7Up0y5bPj9Z5CnLM33kc2N_-2L_M3Ze7J7t-VgKO2Jaa_B3H6eEeGGztUHYP8mo6Tnj0_1LHK6pzvJsmxIsFeayP_XosPF4gfhf3vHjFm9QgNDvu8hynZoUgAuXYugcqu2iwafkpQ5ZNH8vcpfCC7PBs_qIGVKBAE75ZeoLSovGxY17OIQIzkGHtScwgS1AUgidHDnEEwIieljdLH1n3J46UmL3s9u3UBtbpvhhg3yjd-Mu4fGWM8-fW9BUAMDMFpOTTPb4u0uS6rm31EUTBaNJFATbfV-E_7iqrAsjZaVfy92vvlLA8eEdvkImoPZsY86p7pgTIHuxqJs1OgY56SD5RJMTpcRZq8M8N6PbN134UoXO_5tpvkAbIcQ3Hpiq208kHncSeUIt7Rah3sI3LAgkBrv227x5LUKOvnJjm_mMpNbAwDt965lINR9a_-ME_J_fT-uaGKVE93nUV4JBN2YqwtsP90LGsmtFFSbKhcyWNXmKqkPjr7eTByLy4w0GOoa4cXv2XpBFGX-mfHgb4Zqi9eG6EoFbGglKX6zVzALOratprYdEXEWhZ6Rr-dKFDk9YfWm3uobt"

DB_PATH    = "/home/user/previfuego-facturacion/BASE_DATOS_KFC.xlsx"
OUTPUT_PATH = "/home/user/previfuego-facturacion/REVISION_COTIZACIONES_2026.xlsx"

MONTH_FOLDERS = [
    ("01 ENERO",      "ENERO"),
    ("02 FEBRERO",    "FEBRERO"),
    ("03 MARZO",      "MARZO"),
    ("04 ABRIL",      "ABRIL"),
    ("05 MAYO",       "MAYO"),
    ("06 JUNIO",      "JUNIO"),
    ("07 JULIO",      "JULIO"),
    ("08 AGOSTO",     "AGOSTO"),
    ("09 SEPTIEMBRE", "SEPTIEMBRE"),
    ("10 OCTUBRE",    "OCTUBRE"),
]

RECARGA_MONTHS = {"OCTUBRE", "NOVIEMBRE", "DICIEMBRE"}


# ─── CODE NORMALISATION ──────────────────────────────────────────────────────

def canonical_code(raw: str) -> str:
    s = raw.strip().upper()
    s = re.sub(r'^KFCK', 'K', s)
    s = re.sub(r'^JV',   'V', s)
    s = re.sub(r'^CJNC', 'J', s)
    s = re.sub(r'^CN',   'BS', s)
    m = re.match(r'^([A-Z]+?)0*([1-9]\d*)', s)
    return (m.group(1) + m.group(2)) if m else s


def extract_code_from_filename(filename: str) -> str:
    name = re.sub(r'\.xlsx$', '', filename, flags=re.IGNORECASE)
    m = re.match(r'^COTIZACION\d+\s+(\S+)', name, re.IGNORECASE)
    return canonical_code(m.group(1)) if m else ""


# ─── DROPBOX ─────────────────────────────────────────────────────────────────

def dbx_list(path: str) -> list:
    url = "https://api.dropboxapi.com/2/files/list_folder"
    hdrs = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}
    r = requests.post(url, headers=hdrs, data=json.dumps({"path": path, "limit": 2000}))
    if r.status_code != 200:
        print(f"  [WARN] list_folder failed {path}: {r.status_code}")
        return []
    result = r.json()
    entries = result.get("entries", [])
    while result.get("has_more"):
        r2 = requests.post(
            "https://api.dropboxapi.com/2/files/list_folder/continue",
            headers=hdrs,
            data=json.dumps({"cursor": result["cursor"]}),
        )
        result = r2.json()
        entries.extend(result.get("entries", []))
    return entries


def dbx_download(path: str) -> bytes:
    r = requests.post(
        "https://content.dropboxapi.com/2/files/download",
        headers={"Authorization": f"Bearer {TOKEN}",
                 "Dropbox-API-Arg": json.dumps({"path": path})},
    )
    r.raise_for_status()
    return r.content


# ─── COTIZACION PARSING ───────────────────────────────────────────────────────

def parse_cotizacion(wb: openpyxl.Workbook) -> dict:
    """
    Returns {subtotal, items: [{qty, desc, unit_price, line_total}]}
    Items are rows where col3=numeric qty, col4=description, col10=unit price.
    """
    ws = wb.active
    items = []
    subtotal = None
    total = None

    for row in ws.iter_rows():
        cells = {c.column: c.value for c in row if c.value is not None}
        # Line item: col3 = integer qty, col4 = string desc starting with EXTINTOR
        qty = cells.get(3)
        desc = cells.get(4, "")
        unit = cells.get(10)
        ltot = cells.get(13)
        if (isinstance(qty, (int, float)) and qty > 0
                and isinstance(desc, str) and "EXTINTOR" in desc.upper()
                and isinstance(unit, (int, float))):
            items.append({
                "qty": int(qty),
                "desc": desc.strip(),
                "unit_price": float(unit),
                "line_total": float(ltot) if isinstance(ltot, (int, float)) else float(qty * unit),
            })
        # SUBTOTAL label
        for c in row:
            if c.value and str(c.value).strip().upper() == "SUBTOTAL":
                for c2 in ws[c.row]:
                    if c2.column > c.column and isinstance(c2.value, (int, float)):
                        subtotal = float(c2.value)
                        break
            if c.value and str(c.value).strip().upper() in ("TOTAL", "TOTAL:"):
                for c2 in ws[c.row]:
                    if c2.column > c.column and isinstance(c2.value, (int, float)):
                        total = float(c2.value)
                        break

    if subtotal is None and total is not None:
        subtotal = round(total / 1.15, 2)

    return {"subtotal": subtotal, "items": items}


def fmt_items(items: list, price_key: str = "unit_price") -> str:
    """Format line items as '2×50lbCO2@$20 + 1×10lbPQS@$4'."""
    parts = []
    for it in items:
        desc_short = shorten_desc(it["desc"])
        parts.append(f"{it['qty']}×{desc_short}@${it[price_key]:.0f}")
    return " + ".join(parts) if parts else "(sin líneas)"


def shorten_desc(desc: str) -> str:
    """Convert 'EXTINTOR 50 LIBRAS - GAS CARBONICO' → '50lbCO2'."""
    d = desc.upper()
    # Extract capacity
    cap_m = re.search(r'(\d[\d,\.]*)\s*(LIBRAS?|GLNS?|KGS?|LBS?)', d)
    cap = cap_m.group(0).replace(" ", "").replace("LIBRAS", "lb").replace("LIBRA", "lb") \
                        .replace("GLNS", "gln").replace("GLN", "gln") \
                        .replace("KGS", "kg").replace("LBS", "lb") if cap_m else "?"
    cap = re.sub(r'[,]', '.', cap)
    # Extract type
    if "CARBONICO" in d or "CO2" in d or "DIÓXIDO" in d:
        tipo = "CO2"
    elif "POLVO" in d or "PQS" in d or "ABC" in d:
        tipo = "PQS"
    elif "POTASIO" in d or "ACETATO" in d or "CLASE K" in d or "COCINA" in d:
        tipo = "K"
    elif "HALOTRON" in d or "HALON" in d:
        tipo = "HAL"
    elif "AGUA" in d:
        tipo = "H2O"
    else:
        tipo = "?"
    return f"{cap}{tipo}"


# ─── DB LOADING ───────────────────────────────────────────────────────────────

def load_db_full(db_path: str):
    """
    Returns:
      resumen: dict  canonical_code → {codigo, nombre, mes, mantt, recarga, cobro, n_ext}
      detalle: dict  canonical_code → list of {tipo, cap, ubic, mantt, recarga}
    """
    wb = openpyxl.load_workbook(db_path)

    # RESUMEN_LOCALES
    ws_res = wb["RESUMEN_LOCALES"]
    resumen = {}
    for row in ws_res.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        codigo = str(row[0])
        canon = canonical_code(codigo)
        resumen[canon] = {
            "codigo": codigo,
            "nombre": str(row[2]) if row[2] else "",
            "mes": str(row[3]).upper() if row[3] else "",
            "n_ext": row[4],
            "mantt": float(row[5]) if row[5] is not None else None,
            "recarga": float(row[6]) if row[6] is not None else None,
            "cobro": float(row[7]) if row[7] is not None else None,
            "ano_recarga": int(row[9]) if row[9] is not None else None,
        }

    # DETALLE — extract canonical code from NOMBRE_LOCAL "K079 - BABAHOYO"
    ws_det = wb["DETALLE"]
    detalle = defaultdict(list)
    for row in ws_det.iter_rows(min_row=2, values_only=True):
        nombre_local = row[1]
        if not nombre_local or not isinstance(nombre_local, str):
            continue
        # Skip TOTALES rows
        if "TOTALES" in nombre_local.upper():
            continue
        # Extract code: "K079 - BABAHOYO" → "K079"
        m = re.match(r'^([A-Z0-9]+)\s*[-–]', nombre_local.strip())
        if not m:
            continue
        canon = canonical_code(m.group(1))
        detalle[canon].append({
            "tipo": str(row[4]) if row[4] else "",
            "cap":  str(row[5]) if row[5] else "",
            "ubic": str(row[3]) if row[3] else "",
            "mantt":   float(row[6]) if row[6] is not None else 0,
            "recarga": float(row[7]) if row[7] is not None else 0,
        })

    return resumen, detalle


def fmt_db_items(db_exts: list, tipo_cobro: str) -> str:
    """Format DB extintores as '2×50lbCO2@$20 + 1×10lbPQS@$4'."""
    if not db_exts:
        return "(vacío en DB)"
    parts = []
    for e in db_exts:
        cap_short = e["cap"].replace(" LBS", "lb").replace(" LB", "lb") \
                            .replace(" GLNS", "gln").replace(" GLN", "gln") \
                            .replace(" KGS", "kg").replace(" KG", "kg")
        cap_short = re.sub(r'\s+', '', cap_short)
        tipo_short = {"CO2": "CO2", "PQS": "PQS", "K": "K", "HALOTRON": "HAL"}.get(e["tipo"].upper(), e["tipo"])
        price = e["recarga"] if tipo_cobro == "RECARGA" else e["mantt"]
        parts.append(f"1×{cap_short}{tipo_short}@${price:.0f}")
    return " + ".join(parts)


def build_analysis(cot_items: list, db_exts: list, subtotal_cot: float,
                   valor_db: float, tipo_cobro: str) -> str:
    """
    Compare cotizacion line items vs DB extintores.
    Returns a short analysis string.
    """
    n_cot = sum(it["qty"] for it in cot_items)
    n_db  = len(db_exts)
    price_key = "recarga" if tipo_cobro == "RECARGA" else "mantt"

    # Sum of cotizacion prices
    sum_cot = sum(it["qty"] * it["unit_price"] for it in cot_items)
    sum_db  = sum(e[price_key] for e in db_exts)

    parts = []

    # Number difference
    if n_cot != n_db:
        diff_n = n_cot - n_db
        parts.append(f"{'+'if diff_n>0 else ''}{diff_n} ext en COT vs DB ({n_cot} vs {n_db})")

    # Price sum comparison (vs subtotal)
    if abs(sum_cot - (subtotal_cot or 0)) > 0.5:
        parts.append(f"⚠ suma líneas COT=${sum_cot:.2f} ≠ subtotal=${subtotal_cot:.2f}")

    # Possible explanation for large diffs
    diff_abs = abs((subtotal_cot or 0) - valor_db)
    if diff_abs > 0:
        # Check if cot looks like a different number of each extinguisher type
        cot_by_type = defaultdict(lambda: {"qty": 0, "unit": 0})
        for it in cot_items:
            t = shorten_desc(it["desc"])
            cot_by_type[t]["qty"] += it["qty"]
            cot_by_type[t]["unit"] = it["unit_price"]

        db_by_type = defaultdict(int)
        for e in db_exts:
            cap_short = re.sub(r'\s+', '', e["cap"].replace(" LBS","lb").replace(" GLNS","gln").replace(" KGS","kg").replace(" LB","lb"))
            tipo_s = {"CO2":"CO2","PQS":"PQS","K":"K","HALOTRON":"HAL"}.get(e["tipo"].upper(), e["tipo"])
            db_by_type[f"{cap_short}{tipo_s}"] += 1

        extras = []
        for t, v in cot_by_type.items():
            db_n = db_by_type.get(t, 0)
            if v["qty"] > db_n:
                extras.append(f"+{v['qty']-db_n}×{t}(@${v['unit']:.0f})")

        missing = []
        for t, db_n in db_by_type.items():
            cot_n = cot_by_type.get(t, {}).get("qty", 0)
            if db_n > cot_n:
                missing.append(f"-{db_n-cot_n}×{t}")

        if extras:
            parts.append(f"Extra en COT: {', '.join(extras)}")
        if missing:
            parts.append(f"Falta en COT: {', '.join(missing)}")

    return " | ".join(parts) if parts else f"Precio unitario distinto (${sum_cot/max(n_cot,1):.2f} vs ${valor_db/max(n_db,1):.2f}/ext)"


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("Cargando base de datos...")
    resumen, detalle = load_db_full(DB_PATH)
    print(f"  {len(resumen)} locales en RESUMEN, {sum(len(v) for v in detalle.values())} filas en DETALLE")

    rows = []

    for folder_name, mes_name in MONTH_FOLDERS:
        dropbox_path = f"/Previfuego/2026/{folder_name}/GRUPO KFC"
        print(f"\n{'='*60}\nProcesando: {folder_name}")

        entries = dbx_list(dropbox_path)
        if not entries:
            print(f"  [WARN] vacío o no encontrado: {dropbox_path}")
            continue

        xlsx_files = [
            e for e in entries
            if e[".tag"] == "file"
            and e["name"].lower().endswith(".xlsx")
            and e["name"].upper().startswith("COTIZACION")
        ]
        print(f"  {len(xlsx_files)} COTIZACION*.xlsx")

        for entry in xlsx_files:
            filename  = entry["name"]
            file_path = entry["path_lower"]
            local_code = extract_code_from_filename(filename)
            print(f"  [{filename}] → {local_code!r}", end="")

            # Download
            try:
                content = dbx_download(file_path)
            except Exception as e:
                print(f" → DOWNLOAD ERROR: {e}")
                rows.append(_row(mes_name, filename, local_code, "", None, None, "", None, "ERROR",
                                 "", "", f"Download: {e}"))
                continue

            # Parse xlsx
            try:
                wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
            except BadZipFile:
                print(" → CORRUPTED")
                rows.append(_row(mes_name, filename, local_code, "", None, None, "", None, "ERROR",
                                 "", "", "Archivo corrupto"))
                continue
            except Exception as e:
                print(f" → PARSE ERROR: {e}")
                rows.append(_row(mes_name, filename, local_code, "", None, None, "", None, "ERROR",
                                 "", "", str(e)))
                continue

            parsed = parse_cotizacion(wb)
            subtotal  = parsed["subtotal"]
            cot_items = parsed["items"]

            # DB lookup
            db_entry = resumen.get(local_code)
            if db_entry is None:
                print(f" → SIN MATCH DB (subtotal={subtotal})")
                rows.append(_row(mes_name, filename, local_code, "", subtotal, None, "",
                                 None, "SIN_MATCH_DB", fmt_items(cot_items), "", ""))
                continue

            nombre    = db_entry["nombre"]
            db_exts   = detalle.get(local_code, [])
            # If DB says AÑO_RECARGA=2026, this local does recarga in current ENE-SEP cycle
            if db_entry.get("ano_recarga") == 2026 and mes_name not in RECARGA_MONTHS:
                tipo = "RECARGA"
            else:
                tipo = "RECARGA" if mes_name in RECARGA_MONTHS else "MANTENIMIENTO"
            valor_db  = db_entry["recarga"] if tipo == "RECARGA" else db_entry["mantt"]

            # Status
            if subtotal is None:
                estado = "SIN_SUBTOTAL"; diff_pct = None
            elif not valor_db:
                estado = "SIN_VALOR_DB"; diff_pct = None
            else:
                diff_pct = abs(subtotal - valor_db) / valor_db * 100
                if diff_pct <= 1.0:
                    estado = "OK"
                elif tipo == "RECARGA" and db_entry["mantt"]:
                    diff_vs_mantt = abs(subtotal - db_entry["mantt"]) / db_entry["mantt"] * 100
                    estado = "COT_DEBE_RECARGA" if diff_vs_mantt <= 1.0 else "INCONSISTENCIA"
                else:
                    estado = "INCONSISTENCIA"

            print(f" → {estado}" + (f" diff={diff_pct:.1f}%" if diff_pct else ""))

            # Build detail strings only for non-OK rows
            cot_det = db_det = analisis = ""
            if estado in ("INCONSISTENCIA", "COT_DEBE_RECARGA"):
                cot_det  = fmt_items(cot_items)
                db_det   = fmt_db_items(db_exts, tipo)
                analisis = build_analysis(cot_items, db_exts, subtotal, valor_db or 0, tipo)
                if estado == "COT_DEBE_RECARGA":
                    analisis = (f"Cotización usa precio MANTENIMIENTO (${subtotal}). "
                                f"Debe ser RECARGA (${valor_db}). | {analisis}").rstrip(" |")

            rows.append(_row(mes_name, filename, local_code, nombre,
                             subtotal, valor_db, tipo,
                             round(diff_pct, 2) if diff_pct is not None else None,
                             estado, cot_det, db_det, analisis))

    # ─── EXCEL ────────────────────────────────────────────────────────────────
    print(f"\n{'='*60}\nGenerando Excel…")
    _write_excel(rows)

    # ─── PRINT SUMMARY ────────────────────────────────────────────────────────
    total = len(rows)
    ok              = sum(1 for r in rows if r["ESTADO"] == "OK")
    inconsistencias = sum(1 for r in rows if r["ESTADO"] == "INCONSISTENCIA")
    cot_rec         = sum(1 for r in rows if r["ESTADO"] == "COT_DEBE_RECARGA")
    sin_match       = sum(1 for r in rows if r["ESTADO"] == "SIN_MATCH_DB")
    errores         = sum(1 for r in rows if r["ESTADO"] == "ERROR")

    print(f"\n{'='*60}")
    print(f"  Total revisados  : {total}")
    print(f"  OK               : {ok}")
    print(f"  INCONSISTENCIAS  : {inconsistencias}")
    print(f"  COT_DEBE_RECARGA : {cot_rec}")
    print(f"  SIN MATCH DB     : {sin_match}")
    print(f"  ERRORES          : {errores}")

    print("\nINCONSISTENCIAS:")
    for r in rows:
        if r["ESTADO"] == "INCONSISTENCIA":
            print(f"  {r['MES']:12} {r['LOCAL_CODE']:8} {str(r['LOCAL_NOMBRE'])[:32]:32} "
                  f"COT=${r['SUBTOTAL_COT']:7.2f} DB=${r['VALOR_DB']:7.2f} "
                  f"({r['DIFERENCIA%']:.1f}%) | {r['ANALISIS'][:80]}")


# ─── HELPERS ─────────────────────────────────────────────────────────────────

HEADERS = ["MES", "ARCHIVO", "LOCAL_CODE", "LOCAL_NOMBRE",
           "SUBTOTAL_COT", "VALOR_DB", "TIPO", "DIFERENCIA%",
           "ESTADO", "COT_DETALLE", "DB_DETALLE", "ANALISIS"]


def _row(mes, arch, code, nombre, subtot, val_db, tipo, diff, estado, cot_det, db_det, analisis):
    return {
        "MES": mes, "ARCHIVO": arch, "LOCAL_CODE": code, "LOCAL_NOMBRE": nombre,
        "SUBTOTAL_COT": subtot, "VALOR_DB": val_db, "TIPO": tipo,
        "DIFERENCIA%": diff, "ESTADO": estado,
        "COT_DETALLE": cot_det, "DB_DETALLE": db_det, "ANALISIS": analisis,
    }


def _write_excel(rows: list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "REVISION"

    thin   = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(bold=True, color="FFFFFF")

    ws.append(HEADERS)
    for ci in range(1, len(HEADERS) + 1):
        c = ws.cell(row=1, column=ci)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    fill_ok    = PatternFill("solid", fgColor="C6EFCE")
    fill_inc   = PatternFill("solid", fgColor="FFC7CE")
    fill_rec   = PatternFill("solid", fgColor="F4B942")
    fill_warn  = PatternFill("solid", fgColor="FFEB9C")
    fill_err   = PatternFill("solid", fgColor="D9D9D9")

    for r in rows:
        ws.append([r[h] for h in HEADERS])
        rn = ws.max_row
        est = r["ESTADO"]
        fill = (fill_ok  if est == "OK"            else
                fill_inc if est == "INCONSISTENCIA" else
                fill_rec if est == "COT_DEBE_RECARGA" else
                fill_warn if est in ("SIN_MATCH_DB","SIN_SUBTOTAL","SIN_VALOR_DB") else
                fill_err)
        for ci in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=rn, column=ci)
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(ci >= 10))

    # Column widths
    for ci, w in enumerate([10, 52, 12, 38, 12, 10, 14, 11, 18, 55, 55, 80], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

    # ── RESUMEN sheet ──────────────────────────────────────────────────────
    ws2 = wb.create_sheet("RESUMEN")
    total = len(rows)
    ok   = sum(1 for r in rows if r["ESTADO"] == "OK")
    inc  = sum(1 for r in rows if r["ESTADO"] == "INCONSISTENCIA")
    rec  = sum(1 for r in rows if r["ESTADO"] == "COT_DEBE_RECARGA")
    sm   = sum(1 for r in rows if r["ESTADO"] == "SIN_MATCH_DB")
    err  = sum(1 for r in rows if r["ESTADO"] == "ERROR")

    for sd in [
        ["RESUMEN REVISION COTIZACIONES 2026", ""],
        ["", ""],
        ["Total archivos revisados",                    total],
        ["✅ OK (≤1% diferencia)",                       ok],
        ["🔴 INCONSISTENCIA de precio (>1%)",            inc],
        ["🟠 COT_DEBE_RECARGA (emitida como MANT/OCT)", rec],
        ["🟡 SIN MATCH en DB (otras marcas)",            sm],
        ["⚫ ERRORES (descarga/parseo)",                  err],
    ]:
        ws2.append(sd)
    ws2["A1"].font = Font(bold=True, size=13)
    ws2.column_dimensions["A"].width = 48
    ws2.column_dimensions["B"].width = 12

    ws2.append([""])
    ws2.append(["DETALLE POR MES", "TOTAL", "OK", "INCONSISTENCIA", "COT_DEBE_RECARGA", "SIN_MATCH", "OTROS"])
    by_month = defaultdict(list)
    for r in rows:
        by_month[r["MES"]].append(r)
    for _, mes in MONTH_FOLDERS:
        mr = by_month[mes]
        if not mr:
            continue
        m_ok  = sum(1 for r in mr if r["ESTADO"] == "OK")
        m_inc = sum(1 for r in mr if r["ESTADO"] == "INCONSISTENCIA")
        m_rec = sum(1 for r in mr if r["ESTADO"] == "COT_DEBE_RECARGA")
        m_sm  = sum(1 for r in mr if r["ESTADO"] == "SIN_MATCH_DB")
        m_oth = len(mr) - m_ok - m_inc - m_rec - m_sm
        ws2.append([mes, len(mr), m_ok, m_inc, m_rec, m_sm, m_oth])

    wb.save(OUTPUT_PATH)
    print(f"Guardado: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
