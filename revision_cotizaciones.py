#!/usr/bin/env python3
"""
Script to compare Dropbox cotizacion files against BASE_DATOS_KFC.xlsx
and generate a revision report.
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

# ─── CONFIG ─────────────────────────────────────────────────────────────────
TOKEN = "sl.u.AGjIpgezlnwaiHgKQ7Bu5YYhccUjjteEDF1RqACnZjg3vFiMjL5x5d3IaLviJ9dyORvEYZT9jSRe2ocUFRy6vc_0Ls5NPIrwAB6HDxh0Rief7rixFm6VGuCmG1QG0La0ZOoKlDsxocS9ewt1m24m3MXbNcfn88LKLmsQpMw2goxO_qOqtX1yas7JTvt5mDkdplEnlMmHEdQXg3zIgeIj0lFlIzVJwaLgfZfVNL1GwShY5RwsMi_aKAm-VWi6Y5rraWcnvPDbCl2ZNy70KAT-0GNYLKlipCrcjFkCKzBvqfWX8zffKCm9A_mTqWpb6N1ZZ-MPAQICCro_UuZ9zVUvgp6sPQVlHS7vfDRiLLVZ9LIs7-h2DbCct4f4FX3fqszjgAr0LFKak5vCNv8zymvTmzTEOucEEsi5nQ4e7UNE7ixEiZ1kZ2WwsFihVJWMxhP4kFjChzHdjSJ12ubYy2JLjNKpAtLBxhVVqV6uzqoH2rkIWmEnmdpS5EjyxRGni7kZfJQZG4Bdy5IJlqJYcBtQmxE9u-UFnBe28LHnUVJB1s50kmV8yNpkvzFlqC0CtGuvLj6FgjsfsUSMOroG8gi-eZaRGDEj9jqqpMsfW9NSb7iz6budr6_5OhzxQ3w3YZRVnB70EQBdHxWZTE_8So9s-y31vhN97tE9lWqqMl0lR9VzjeBU7c5zZY3N1N6oE1Yf-yyXKGxjdC6dU7Up0y5bPj9Z5CnLM33kc2N_-2L_M3Ze7J7t-VgKO2Jaa_B3H6eEeGGztUHYP8mo6Tnj0_1LHK6pzvJsmxIsFeayP_XosPF4gfhf3vHjFm9QgNDvu8hynZoUgAuXYugcqu2iwafkpQ5ZNH8vcpfCC7PBs_qIGVKBAE75ZeoLSovGxY17OIQIzkGHtScwgS1AUgidHDnEEwIieljdLH1n3J46UmL3s9u3UBtbpvhhg3yjd-Mu4fGWM8-fW9BUAMDMFpOTTPb4u0uS6rm31EUTBaNJFATbfV-E_7iqrAsjZaVfy92vvlLA8eEdvkImoPZsY86p7pgTIHuxqJs1OgY56SD5RJMTpcRZq8M8N6PbN134UoXO_5tpvkAbIcQ3Hpiq208kHncSeUIt7Rah3sI3LAgkBrv227x5LUKOvnJjm_mMpNbAwDt965lINR9a_-ME_J_fT-uaGKVE93nUV4JBN2YqwtsP90LGsmtFFSbKhcyWNXmKqkPjr7eTByLy4w0GOoa4cXv2XpBFGX-mfHgb4Zqi9eG6EoFbGglKX6zVzALOratprYdEXEWhZ6Rr-dKFDk9YfWm3uobt"

DB_PATH = "/home/user/previfuego-facturacion/BASE_DATOS_KFC.xlsx"
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

# Months Jan-Sep → MANTENIMIENTO, Oct-Dec → RECARGA
RECARGA_MONTHS = {"OCTUBRE", "NOVIEMBRE", "DICIEMBRE"}


# ─── HELPERS ────────────────────────────────────────────────────────────────

def canonical_code(raw: str) -> str:
    """
    Normalise a raw code to canonical form for matching.
    Examples:
      KFCK02 → K2    (strip KFCK prefix, strip leading zeros)
      KFCK002 → K2
      JV55   → V55   (strip J prefix)
      M058   → M58   (strip leading zeros in numeric part)
      BS009  → BS9
      K002   → K2
      V055   → V55
      CJNC01 → J1
    """
    s = raw.strip().upper()
    # Strip prefix KFCK → K
    s = re.sub(r'^KFCK', 'K', s)
    # Strip prefix JV → V
    s = re.sub(r'^JV', 'V', s)
    # Strip prefix CJNC → J (CAJUN codes)
    s = re.sub(r'^CJNC', 'J', s)
    # CN → BS (Cinnabon = Baskin-Robin, misma ubicación)
    s = re.sub(r'^CN', 'BS', s)
    # Now remove leading zeros from the numeric suffix
    # e.g. K002 → K2, BS009 → BS9, M058 → M58
    # Also strip any trailing non-numeric suffix (e.g. K91PASEO → K91)
    m = re.match(r'^([A-Z]+?)0*([1-9]\d*)', s)
    if m:
        return m.group(1) + m.group(2)
    return s


def extract_code_from_filename(filename: str) -> str:
    """
    Extract local code from cotizacion filename.
    E.g.:
      'COTIZACION1262 KFCK02 9 DE OCT Y CHIMBORAZO.xlsx' → K2
      'COTIZACION1281 M58 PENAS.xlsx'                    → M58
      'COTIZACION1260 JV55 MALL DEL SOL.xlsx'            → V55
      'COTIZACION1275 BS09 ...'                          → BS9
    """
    # Remove extension
    name = re.sub(r'\.xlsx$', '', filename, flags=re.IGNORECASE)
    # Pattern: COTIZACION<num> <CODE> <rest...>
    m = re.match(r'^COTIZACION\d+\s+(\S+)', name, re.IGNORECASE)
    if m:
        return canonical_code(m.group(1))
    return ""


def dropbox_list_folder(path: str) -> list:
    """List files in a Dropbox folder, handling pagination."""
    url = "https://api.dropboxapi.com/2/files/list_folder"
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json",
    }
    data = json.dumps({"path": path, "limit": 2000})
    r = requests.post(url, headers=headers, data=data)
    if r.status_code != 200:
        print(f"  [WARN] list_folder failed for {path}: {r.status_code} {r.text[:200]}")
        return []
    result = r.json()
    entries = result.get("entries", [])
    while result.get("has_more"):
        cursor = result["cursor"]
        r2 = requests.post(
            "https://api.dropboxapi.com/2/files/list_folder/continue",
            headers=headers,
            data=json.dumps({"cursor": cursor}),
        )
        result = r2.json()
        entries.extend(result.get("entries", []))
    return entries


def dropbox_download(path: str) -> bytes:
    """Download a file from Dropbox. Use json.dumps for header to handle unicode."""
    url = "https://content.dropboxapi.com/2/files/download"
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Dropbox-API-Arg": json.dumps({"path": path}),
    }
    r = requests.post(url, headers=headers)
    r.raise_for_status()
    return r.content


def parse_subtotal(wb: openpyxl.Workbook) -> tuple:
    """
    Scan all cells for 'SUBTOTAL' label and return the numeric value
    adjacent to it (same row, looking right).
    Returns (subtotal_value_or_None, note_string)
    """
    ws = wb.active
    subtotal_val = None
    total_val = None

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            cell_str = str(cell.value).strip().upper()
            if cell_str == "SUBTOTAL":
                # Look right on the same row for the first numeric value
                for c2 in ws[cell.row]:
                    if c2.column > cell.column and isinstance(c2.value, (int, float)):
                        subtotal_val = float(c2.value)
                        break
            elif cell_str in ("TOTAL", "TOTAL:"):
                for c2 in ws[cell.row]:
                    if c2.column > cell.column and isinstance(c2.value, (int, float)):
                        total_val = float(c2.value)
                        break

    if subtotal_val is not None:
        return subtotal_val, ""

    # Fallback: back-calculate from TOTAL / 1.15
    if total_val is not None:
        subtotal_val = round(total_val / 1.15, 2)
        return subtotal_val, "SUBTOTAL back-calculated from TOTAL/1.15"

    return None, "SUBTOTAL not found"


# ─── LOAD DATABASE ──────────────────────────────────────────────────────────

def load_database(db_path: str) -> dict:
    """
    Load RESUMEN_LOCALES sheet.
    Returns dict: canonical_code → {codigo, nombre, mes, mantt, recarga, cobro}
    """
    wb = openpyxl.load_workbook(db_path)
    ws = wb["RESUMEN_LOCALES"]
    db = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        codigo = row[0]
        marca = row[1]
        nombre = row[2]
        mes = row[3]
        n_ext = row[4]
        mantt = row[5]
        recarga = row[6]
        cobro = row[7]
        canon = canonical_code(str(codigo))
        db[canon] = {
            "codigo": str(codigo),
            "marca": str(marca) if marca else "",
            "nombre": str(nombre) if nombre else "",
            "mes": str(mes).upper() if mes else "",
            "n_extintores": n_ext,
            "mantt": float(mantt) if mantt is not None else None,
            "recarga": float(recarga) if recarga is not None else None,
            "cobro": float(cobro) if cobro is not None else None,
        }
    return db


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    print("Loading database...")
    db = load_database(DB_PATH)
    print(f"  Loaded {len(db)} locals from database")

    # Print some canonical keys to verify
    sample_keys = list(db.keys())[:10]
    print(f"  Sample canonical keys: {sample_keys}")

    rows = []  # Result dicts

    for folder_name, mes_name in MONTH_FOLDERS:
        dropbox_path = f"/Previfuego/2026/{folder_name}/GRUPO KFC"
        print(f"\n{'='*60}")
        print(f"Processing: {folder_name}")

        entries = dropbox_list_folder(dropbox_path)
        if not entries:
            print(f"  [WARN] Empty or not found: {dropbox_path}")
            continue

        xlsx_files = [
            e for e in entries
            if e[".tag"] == "file"
            and e["name"].lower().endswith(".xlsx")
            and e["name"].upper().startswith("COTIZACION")
        ]
        print(f"  Found {len(xlsx_files)} COTIZACION*.xlsx files")

        for entry in xlsx_files:
            filename = entry["name"]
            file_path = entry["path_lower"]

            local_code = extract_code_from_filename(filename)
            print(f"  [{filename}] → code={local_code!r}", end="")

            # Download file
            try:
                content = dropbox_download(file_path)
            except Exception as e:
                print(f" → DOWNLOAD ERROR: {e}")
                rows.append({
                    "MES": mes_name,
                    "ARCHIVO": filename,
                    "LOCAL_CODE": local_code,
                    "LOCAL_NOMBRE": "",
                    "SUBTOTAL_COT": None,
                    "VALOR_DB": None,
                    "TIPO": "",
                    "DIFERENCIA%": None,
                    "ESTADO": "ERROR",
                    "NOTAS": f"Download error: {e}",
                })
                continue

            # Parse xlsx
            try:
                wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
            except BadZipFile:
                print(f" → CORRUPTED (BadZipFile)")
                rows.append({
                    "MES": mes_name,
                    "ARCHIVO": filename,
                    "LOCAL_CODE": local_code,
                    "LOCAL_NOMBRE": "",
                    "SUBTOTAL_COT": None,
                    "VALOR_DB": None,
                    "TIPO": "",
                    "DIFERENCIA%": None,
                    "ESTADO": "ERROR",
                    "NOTAS": "Archivo corrupto (BadZipFile)",
                })
                continue
            except Exception as e:
                print(f" → PARSE ERROR: {e}")
                rows.append({
                    "MES": mes_name,
                    "ARCHIVO": filename,
                    "LOCAL_CODE": local_code,
                    "LOCAL_NOMBRE": "",
                    "SUBTOTAL_COT": None,
                    "VALOR_DB": None,
                    "TIPO": "",
                    "DIFERENCIA%": None,
                    "ESTADO": "ERROR",
                    "NOTAS": f"Parse error: {e}",
                })
                continue

            subtotal, note = parse_subtotal(wb)

            # Look up DB
            db_entry = db.get(local_code)
            if db_entry is None:
                print(f" → NOT IN DB (subtotal={subtotal})")
                rows.append({
                    "MES": mes_name,
                    "ARCHIVO": filename,
                    "LOCAL_CODE": local_code,
                    "LOCAL_NOMBRE": "",
                    "SUBTOTAL_COT": subtotal,
                    "VALOR_DB": None,
                    "TIPO": "",
                    "DIFERENCIA%": None,
                    "ESTADO": "SIN_MATCH_DB",
                    "NOTAS": f"Codigo {local_code!r} no encontrado en DB. {note}".strip(),
                })
                continue

            nombre = db_entry["nombre"]
            mes_db = db_entry["mes"]

            # Determine correct value based on month
            if mes_name in RECARGA_MONTHS:
                tipo = "RECARGA"
                valor_db = db_entry["recarga"]
            else:
                tipo = "MANTENIMIENTO"
                valor_db = db_entry["mantt"]

            # Determine status
            if subtotal is None:
                estado = "SIN_SUBTOTAL"
                diff_pct = None
            elif valor_db is None or valor_db == 0:
                estado = "SIN_VALOR_DB"
                diff_pct = None
            else:
                diff_pct = abs(subtotal - valor_db) / valor_db * 100
                if diff_pct <= 1.0:
                    estado = "OK"
                else:
                    estado = "INCONSISTENCIA"

            if diff_pct is not None:
                print(f" → {estado} subtotal={subtotal} db={valor_db} diff={diff_pct:.1f}%")
            else:
                print(f" → {estado} subtotal={subtotal} db={valor_db}")

            rows.append({
                "MES": mes_name,
                "ARCHIVO": filename,
                "LOCAL_CODE": local_code,
                "LOCAL_NOMBRE": nombre,
                "SUBTOTAL_COT": subtotal,
                "VALOR_DB": valor_db,
                "TIPO": tipo,
                "DIFERENCIA%": round(diff_pct, 2) if diff_pct is not None else None,
                "ESTADO": estado,
                "NOTAS": note if note else "",
            })

    # ─── GENERATE EXCEL REPORT ────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"Generating report: {OUTPUT_PATH}")

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "REVISION"

    headers = [
        "MES", "ARCHIVO", "LOCAL_CODE", "LOCAL_NOMBRE",
        "SUBTOTAL_COT", "VALOR_DB", "TIPO",
        "DIFERENCIA%", "ESTADO", "NOTAS"
    ]

    # Styles
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws_out.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws_out.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    fill_ok = PatternFill("solid", fgColor="C6EFCE")
    fill_inconsistency = PatternFill("solid", fgColor="FFC7CE")
    fill_warn = PatternFill("solid", fgColor="FFEB9C")
    fill_error = PatternFill("solid", fgColor="D9D9D9")

    for r_dict in rows:
        row_data = [r_dict[h] for h in headers]
        ws_out.append(row_data)
        row_num = ws_out.max_row
        estado = r_dict["ESTADO"]
        if estado == "OK":
            fill = fill_ok
        elif estado == "INCONSISTENCIA":
            fill = fill_inconsistency
        elif estado in ("SIN_SUBTOTAL", "SIN_MATCH_DB", "SIN_VALOR_DB"):
            fill = fill_warn
        else:
            fill = fill_error
        for col_idx in range(1, len(headers) + 1):
            cell = ws_out.cell(row=row_num, column=col_idx)
            cell.fill = fill
            cell.border = border

    # Column widths
    col_widths = [12, 55, 14, 42, 14, 12, 16, 12, 18, 55]
    for i, w in enumerate(col_widths, 1):
        ws_out.column_dimensions[get_column_letter(i)].width = w

    ws_out.freeze_panes = "A2"

    # ─── SUMMARY SHEET ────────────────────────────────────────────────────────
    ws_sum = wb_out.create_sheet("RESUMEN")
    total = len(rows)
    ok = sum(1 for r in rows if r["ESTADO"] == "OK")
    inconsistencias = sum(1 for r in rows if r["ESTADO"] == "INCONSISTENCIA")
    sin_subtotal = sum(1 for r in rows if r["ESTADO"] == "SIN_SUBTOTAL")
    sin_match = sum(1 for r in rows if r["ESTADO"] == "SIN_MATCH_DB")
    sin_valor_db = sum(1 for r in rows if r["ESTADO"] == "SIN_VALOR_DB")
    errors = sum(1 for r in rows if r["ESTADO"] == "ERROR")

    summary_data = [
        ["RESUMEN DE REVISION COTIZACIONES 2026", ""],
        ["", ""],
        ["TOTAL archivos revisados", total],
        ["OK (diferencia <= 1%)", ok],
        ["INCONSISTENCIAS (diferencia > 1%)", inconsistencias],
        ["SIN SUBTOTAL en cotizacion", sin_subtotal],
        ["SIN MATCH en base de datos", sin_match],
        ["SIN VALOR en base de datos", sin_valor_db],
        ["ERRORES (descarga/parseo)", errors],
    ]
    for sd_row in summary_data:
        ws_sum.append(sd_row)
    ws_sum.column_dimensions["A"].width = 42
    ws_sum.column_dimensions["B"].width = 15
    ws_sum["A1"].font = Font(bold=True, size=13)

    ws_sum.append(["", ""])
    ws_sum.append(["DETALLE POR MES", ""])
    ws_sum.append(["MES", "TOTAL", "OK", "INCONSISTENCIA", "OTROS"])

    by_month = defaultdict(list)
    for r in rows:
        by_month[r["MES"]].append(r)

    for folder_name, mes_name in MONTH_FOLDERS:
        month_rows = by_month[mes_name]
        if not month_rows:
            continue
        m_total = len(month_rows)
        m_ok = sum(1 for r in month_rows if r["ESTADO"] == "OK")
        m_inc = sum(1 for r in month_rows if r["ESTADO"] == "INCONSISTENCIA")
        m_other = m_total - m_ok - m_inc
        ws_sum.append([mes_name, m_total, m_ok, m_inc, m_other])

    wb_out.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")

    # ─── PRINT SUMMARY ────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print("SUMMARY:")
    print(f"  Total files reviewed : {total}")
    print(f"  OK                   : {ok}")
    print(f"  INCONSISTENCIAS      : {inconsistencias}")
    print(f"  SIN SUBTOTAL         : {sin_subtotal}")
    print(f"  SIN MATCH DB         : {sin_match}")
    print(f"  SIN VALOR DB         : {sin_valor_db}")
    print(f"  ERRORES              : {errors}")
    print()

    if inconsistencias > 0:
        print("INCONSISTENCIAS DETAIL:")
        for r in rows:
            if r["ESTADO"] == "INCONSISTENCIA":
                print(f"  {r['MES']:12} {r['LOCAL_CODE']:8} {str(r['LOCAL_NOMBRE'])[:35]:35} "
                      f"COT={r['SUBTOTAL_COT']:8.2f} DB={r['VALOR_DB']:8.2f} "
                      f"DIFF={r['DIFERENCIA%']:.1f}%")

    if sin_match > 0:
        print("\nSIN MATCH DB DETAIL:")
        for r in rows:
            if r["ESTADO"] == "SIN_MATCH_DB":
                print(f"  {r['MES']:12} {r['LOCAL_CODE']:8} {r['ARCHIVO']}")


if __name__ == "__main__":
    main()
