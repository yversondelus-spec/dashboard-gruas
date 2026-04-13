import os, sys, json, requests, pandas as pd
from datetime import datetime, date
from io import BytesIO

# ── Configuración ──────────────────────────────────────────────────────────
LIMIT_HRS     = 160
SHEET_URL_IMP = os.environ.get("SHEET_URL_IMPORT", "")
SHEET_URL_EXP = os.environ.get("SHEET_URL_EXPORT", "")

MESES_ES = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
            7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}

# ── FLOTA IMPORT ───────────────────────────────────────────────────────────
GRUAS_IMPORT = [
    {"id":"LINDE 11728","empresa":"Linde Leasing"},
    {"id":"LINDE 11731","empresa":"Linde Leasing"},
    {"id":"LINDE 11732","empresa":"Linde Leasing"},
    {"id":"LINDE 11733","empresa":"Linde Leasing"},
    {"id":"LINDE 11734","empresa":"Linde Leasing"},
    {"id":"LINDE 11735","empresa":"Linde Leasing"},
    {"id":"LINDE 11736","empresa":"Linde Leasing"},
    {"id":"LINDE 11738","empresa":"Linde Leasing"},
    {"id":"LINDE 11739","empresa":"Linde Leasing"},
]
IDS_IMP = [g["id"] for g in GRUAS_IMPORT]

# ── FLOTA EXPORT ───────────────────────────────────────────────────────────
GRUAS_EXPORT = [
    {"id":"EXP 11720","empresa":"Export Leasing"},
    {"id":"EXP 11721","empresa":"Export Leasing"},
    {"id":"EXP 11722","empresa":"Export Leasing"},
    {"id":"EXP 11723","empresa":"Export Leasing"},
    {"id":"EXP 11724","empresa":"Export Leasing"},
    {"id":"EXP 11725","empresa":"Export Leasing"},
    {"id":"EXP 11726","empresa":"Export Leasing"},
    {"id":"EXP 11727","empresa":"Export Leasing"},
    {"id":"EXP 11729","empresa":"Export Leasing"},
    {"id":"EXP 11730","empresa":"Export Leasing"},
    {"id":"EXP 11737","empresa":"Export Leasing"},
    {"id":"EXP 11740","empresa":"Export Leasing"},
]
IDS_EXP = [g["id"] for g in GRUAS_EXPORT]

# ── Utils ──────────────────────────────────────────────────────────────────
def download_excel(url):
    if not url:
        return None
    if "docs.google.com/spreadsheets" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return BytesIO(r.content)

def _parse_val(val):
    if pd.isna(val): return None
    try: return float(val)
    except: return None

# ── Lectura ────────────────────────────────────────────────────────────────
def leer_hoja_import(excel_bytes, year):
    try:
        df = pd.read_excel(excel_bytes, sheet_name=f"SEMANAS {year}",
                           header=None, skiprows=5, usecols=range(18))
    except: return []

    rows = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.iloc[1]).date()
            sem   = int(row.iloc[0])
        except:
            continue

        entry = {"sem": sem, "fecha": fecha}
        for i, g in enumerate(GRUAS_IMPORT):
            entry[g["id"]] = _parse_val(row.iloc[8 + i])
        rows.append(entry)
    return rows

def leer_hoja_export(excel_bytes, year):
    try:
        df = pd.read_excel(excel_bytes, sheet_name=f"SEMANAS {year}",
                           header=None, skiprows=5, usecols=range(14))
    except: return []

    rows = []
    for _, row in df.iterrows():
        try:
            fecha = pd.to_datetime(row.iloc[1]).date()
            sem   = int(row.iloc[0])
        except:
            continue

        entry = {"sem": sem, "fecha": fecha}
        for i, g in enumerate(GRUAS_EXPORT):
            entry[g["id"]] = _parse_val(row.iloc[2 + i])
        rows.append(entry)
    return rows

# ── FIX 1: diferencial con fecha previa ────────────────────────────────────
def calcular_horas_semanales(rows, grua_ids):
    prev_val  = {gid: None for gid in grua_ids}
    prev_date = {gid: None for gid in grua_ids}

    for row in rows:
        for gid in grua_ids:
            val = row.get(gid)

            if isinstance(val, float) and prev_val[gid] is not None:
                row[f"{gid}_hrs"]       = round(max(val - prev_val[gid], 0), 1)
                row[f"{gid}_prev_date"] = prev_date[gid]
            else:
                row[f"{gid}_hrs"]       = None
                row[f"{gid}_prev_date"] = None

            if isinstance(val, float):
                prev_val[gid]  = val
                prev_date[gid] = row["fecha"]

    return rows

# ── Helper corte 20 ────────────────────────────────────────────────────────
def get_20th_boundaries(d_start, d_end):
    boundaries = []
    m, y = d_start.month, d_start.year
    c = date(y, m, 20)

    if c <= d_start:
        m = m % 12 + 1
        y = y + (1 if d_start.month == 12 else 0)
        c = date(y, m, 20)

    while c < d_end:
        boundaries.append(c)
        m = c.month % 12 + 1
        y = c.year + (1 if c.month == 12 else 0)
        c = date(y, m, 20)

    return boundaries

# ── Periodo ────────────────────────────────────────────────────────────────
def periodo_key_label(fecha):
    if fecha.day >= 20:
        inicio = date(fecha.year, fecha.month, 20)
        fin = date(fecha.year + (fecha.month == 12), (fecha.month % 12) + 1, 20)
    else:
        inicio = date(fecha.year - (fecha.month == 1), (fecha.month - 2) % 12 + 1, 20)
        fin = date(fecha.year, fecha.month, 20)

    key   = f"{inicio.strftime('%Y%m%d')}_{fin.strftime('%Y%m%d')}"
    label = f"20 {MESES_ES[inicio.month]} – 20 {MESES_ES[fin.month]} {fin.year}"
    return key, label, inicio, fin

# ── FIX 2: distribución correcta ───────────────────────────────────────────
def agrupar_por_periodo(rows, grua_ids, hoy):
    periodos = {}

    def add(key, label, inicio, fin, gid, hrs):
        if key not in periodos:
            periodos[key] = {
                "key": key, "label": label,
                "inicio": inicio, "fin": fin,
                "hrsporgid": {g:0 for g in grua_ids},
                "tiene_dato": {g:False for g in grua_ids},
                "semanas":[]
            }
        periodos[key]["hrsporgid"][gid] += hrs
        periodos[key]["tiene_dato"][gid] = True

    for row in rows:
        for gid in grua_ids:
            hrs = row.get(f"{gid}_hrs")
            prev = row.get(f"{gid}_prev_date")
            curr = row["fecha"]

            if not hrs or not prev:
                continue

            boundaries = get_20th_boundaries(prev, curr)

            if not boundaries:
                k,l,i,f = periodo_key_label(prev)
                add(k,l,i,f,gid,hrs)
            else:
                total = (curr - prev).days
                pts = [prev] + boundaries + [curr]

                for i in range(len(pts)-1):
                    d1, d2 = pts[i], pts[i+1]
                    dias = (d2 - d1).days
                    if dias == 0: continue

                    seg_hrs = hrs * dias / total
                    k,l,i,f = periodo_key_label(d1)
                    add(k,l,i,f,gid,seg_hrs)

    return periodos

# ── FIX 3: merge sin reset ────────────────────────────────────────────────
def merge_anos(excel_bytes, leer_fn, grua_ids, hoy):
    all_rows = []

    for y in [2024, 2025, 2026]:
        excel_bytes.seek(0)
        all_rows += leer_fn(excel_bytes, y)

    all_rows.sort(key=lambda x: x["fecha"])
    all_rows = calcular_horas_semanales(all_rows, grua_ids)

    return agrupar_por_periodo(all_rows, grua_ids, hoy)

# ── MAIN ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":

    now = datetime.now()
    hoy = now.date()

    imp = download_excel(SHEET_URL_IMP)
    exp = download_excel(SHEET_URL_EXP)

    data_imp = merge_anos(imp, leer_hoja_import, IDS_IMP, hoy)
    data_exp = merge_anos(exp, leer_hoja_export, IDS_EXP, hoy) if exp else {}

    print("OK calculado correctamente")
