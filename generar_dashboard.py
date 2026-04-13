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

# ── FLOTA EXPORT ──────────────────────────────────────────────────────────
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
    try:
        return float(val)
    except:
        return None

# ── Lectura ────────────────────────────────────────────────────────────────
def leer_hoja_import(excel_bytes, year):
    try:
        df = pd.read_excel(excel_bytes, sheet_name=f"SEMANAS {year}",
                           header=None, skiprows=5)
    except ValueError:
        return []
    
    rows = []
    for _, row in df.iterrows():
        try:
            fecha_val = pd.to_datetime(row.iloc[1], dayfirst=True)
            if pd.isna(fecha_val):
                continue
            fecha = fecha_val.date()
        except:
            continue
        
        entry = {"fecha": fecha}
        for i, g in enumerate(GRUAS_IMPORT):
            try:
                entry[g["id"]] = _parse_val(row.iloc[8+i])
            except IndexError:
                entry[g["id"]] = None
        rows.append(entry)
    
    return rows

def leer_hoja_export(excel_bytes, year):
    try:
        df = pd.read_excel(excel_bytes, sheet_name=f"SEMANAS {year}",
                           header=None, skiprows=5)
    except ValueError:
        return []
    
    rows = []
    for _, row in df.iterrows():
        try:
            fecha_val = pd.to_datetime(row.iloc[1], dayfirst=True)
            if pd.isna(fecha_val):
                continue
            fecha = fecha_val.date()
        except:
            continue
        
        entry = {"fecha": fecha}
        for i, gid in enumerate(IDS_EXP):
            try:
                entry[gid] = _parse_val(row.iloc[2+i])
            except IndexError:
                entry[gid] = None
        rows.append(entry)
    
    return rows

# ── FIX 1: Diferenciales con fecha previa ──────────────────────────────────
def calcular_horas_semanales(rows, grua_ids):
    prev_val  = {gid: None for gid in grua_ids}
    prev_date = {gid: None for gid in grua_ids}

    for row in rows:
        for gid in grua_ids:
            val = row.get(gid)

            if isinstance(val, float) and prev_val[gid] is not None:
                row[f"{gid}_hrs"] = max(val - prev_val[gid], 0)
                row[f"{gid}_prev_date"] = prev_date[gid]
            else:
                row[f"{gid}_hrs"] = None
                row[f"{gid}_prev_date"] = None

            if isinstance(val, float):
                prev_val[gid] = val
                prev_date[gid] = row["fecha"]

    return rows

# ── Helper cortes ──────────────────────────────────────────────────────��───
def get_cutoff_boundaries(d_start, d_end):
    """Obtiene los cortes de 21-20 entre dos fechas"""
    boundaries = []
    
    # Comenzar desde el primer corte después de d_start
    m, y = d_start.month, d_start.year
    
    # Primer corte: día 20 del mes siguiente a d_start
    if d_start.day <= 20:
        c = date(y, m, 20)
    else:
        m = m % 12 + 1
        y = y + (1 if m == 1 else 0)
        c = date(y, m, 20)
    
    while c < d_end:
        boundaries.append(c)
        m = c.month % 12 + 1
        y = c.year + (1 if c.month == 12 else 0)
        c = date(y, m, 20)
    
    return boundaries

def periodo_key_label(fecha):
    """Retorna el período 21-20 que contiene la fecha"""
    if fecha.day > 20:
        # Día 21+: período va desde 21 de este mes hasta 20 del siguiente
        inicio = date(fecha.year, fecha.month, 21)
        fin_m = fecha.month % 12 + 1
        fin_y = fecha.year + (1 if fecha.month == 12 else 0)
        fin = date(fin_y, fin_m, 20)
    else:
        # Día 1-20: período va desde 21 del mes anterior hasta 20 de este mes
        inicio_m = fecha.month - 1 if fecha.month > 1 else 12
        inicio_y = fecha.year if fecha.month > 1 else fecha.year - 1
        inicio = date(inicio_y, inicio_m, 21)
        fin = date(fecha.year, fecha.month, 20)
    
    key = f"{inicio}_{fin}"
    label = f"21 {MESES_ES[inicio.month]} – 20 {MESES_ES[fin.month]} {fin.year}"
    return key, label, inicio, fin

# ── FIX 2: Distribución proporcional ───────────────────────────────────────
def agrupar_por_periodo(rows, grua_ids, hoy):
    periodos = {}

    def add(key, label, inicio, fin, gid, hrs):
        if key not in periodos:
            periodos[key] = {
                "label": label,
                "inicio": str(inicio),
                "fin": str(fin),
                "hrsporgid": {g:0 for g in grua_ids}
            }
        periodos[key]["hrsporgid"][gid] += hrs

    for row in rows:
        curr = row["fecha"]

        for gid in grua_ids:
            hrs = row.get(f"{gid}_hrs")
            prev = row.get(f"{gid}_prev_date")

            if not hrs or not prev:
                continue

            bounds = get_cutoff_boundaries(prev, curr)

            if not bounds:
                # No hay cortes entre prev y curr, todo va a un período
                key, label, ini, fin = periodo_key_label(prev)
                add(key, label, ini, fin, gid, hrs)

            else:
                # Hay cortes, distribuir proporcionalmente
                total_days = (curr - prev).days
                tramos = [prev] + bounds + [curr]

                for i in range(len(tramos)-1):
                    d1, d2 = tramos[i], tramos[i+1]
                    days = (d2 - d1).days
                    if days == 0: continue

                    h = hrs * days / total_days
                    key, label, ini, fin = periodo_key_label(d1)
                    add(key, label, ini, fin, gid, h)

    return periodos

# ── FIX 3: Merge sin reset ─────────────────────────────────────────────────
def merge_anos(excel_bytes, leer_fn, grua_ids, hoy):
    all_rows = []

    for y in [2024, 2025, 2026]:
        excel_bytes.seek(0)
        rows = leer_fn(excel_bytes, y)
        if rows:
            all_rows.extend(rows)

    all_rows = [r for r in all_rows if isinstance(r["fecha"], date)]
    all_rows.sort(key=lambda x: x["fecha"])

    all_rows = calcular_horas_semanales(all_rows, grua_ids)

    return agrupar_por_periodo(all_rows, grua_ids, hoy)

# ── GENERAR DASHBOARD DATA ─────────────────────────────────────────────────
def generar_dashboard_data(periodos_imp, periodos_exp, hoy):
    """Convierte periodos a formato para dashboard HTML"""
    
    def procesar_periodo(periodos, grua_ids, limit_hrs=160):
        result = {}
        for key, p in periodos.items():
            hrs_por_grua = p["hrsporgid"]
            
            ok = sum(1 for h in hrs_por_grua.values() if h and h < limit_hrs * 0.6)
            prec = sum(1 for h in hrs_por_grua.values() if h and limit_hrs * 0.6 <= h < limit_hrs * 0.86)
            alert = sum(1 for h in hrs_por_grua.values() if h and limit_hrs * 0.86 <= h < limit_hrs)
            limit = sum(1 for h in hrs_por_grua.values() if h and h >= limit_hrs)
            
            # Cards HTML para grúas
            cards_html = ""
            for gid in grua_ids:
                h = hrs_por_grua.get(gid, 0) or 0
                pct = (h / limit_hrs) * 100 if limit_hrs > 0 else 0
                
                if h < limit_hrs * 0.6:
                    status, badge = "s-ok", "ok"
                    status_text = "✅ OK"
                elif h < limit_hrs * 0.86:
                    status, badge = "s-precaution", "precaution"
                    status_text = "⚠️ Precaución"
                elif h < limit_hrs:
                    status, badge = "s-alert", "alert"
                    status_text = "🔶 Alerta"
                else:
                    status, badge = "s-limit", "limit"
                    status_text = "🔴 Límite"
                
                cards_html += f'''
                <div class="crane-card {status}">
                  <div class="crane-header">
                    <div class="crane-name">{gid.split()[-1]}</div>
                    <span class="status-badge {badge}">{status_text}</span>
                  </div>
                  <div class="crane-plate">{gid}</div>
                  <div class="crane-km">
                    <div class="crane-km-val">{h:.0f}</div>
                    <div class="crane-km-of">/ {limit_hrs} hrs</div>
                  </div>
                  <div class="prog-bar">
                    <div class="prog-fill {badge.replace('-','')}" style="width:{min(pct,100)}%"></div>
                  </div>
                  <div class="crane-footer">
                    <span>{pct:.0f}%</span>
                    <span class="disp">Disponible: {max(0, limit_hrs - h):.0f} hrs</span>
                  </div>
                </div>
                '''
            
            # Bar chart data
            bar_labels = list(hrs_por_grua.keys())
            bar_data = [hrs_por_grua.get(g, 0) or 0 for g in bar_labels]
            bar_labels = [g.split()[-1] for g in bar_labels]  # Solo número
            
            result[key] = {
                "label": p["label"],
                "inicio_label": p["inicio"],
                "fin_label": p["fin"],
                "ok": ok,
                "prec": prec,
                "alert": alert,
                "limit": limit,
                "cards": cards_html,
                "bar_labels": bar_labels,
                "bar_data": bar_data,
                "n_sem": len([1 for h in hrs_por_grua.values() if h])
            }
        return result
    
    data_imp = procesar_periodo(periodos_imp, IDS_IMP)
    data_exp = procesar_periodo(periodos_exp, IDS_EXP) if periodos_exp else {}
    
    # Combinar por período
    all_keys = set(data_imp.keys()) | set(data_exp.keys())
    gruas_data = {}
    
    for key in sorted(all_keys, reverse=True):
        gruas_data[key] = {
            "inicio_label": data_imp[key]["inicio_label"] if key in data_imp else data_exp[key]["inicio_label"],
            "fin_label": data_imp[key]["fin_label"] if key in data_imp else data_exp[key]["fin_label"],
            "imp": data_imp.get(key),
            "exp": data_exp.get(key) if periodos_exp else None,
        }
    
    # Período actual (más reciente)
    periodo_actual = sorted(all_keys, reverse=True)[0] if all_keys else ""
    
    return gruas_data, periodo_actual

# ── MAIN ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    now = datetime.now()
    hoy = now.date()

    raw_imp = download_excel(SHEET_URL_IMP)
    raw_exp = download_excel(SHEET_URL_EXP) if SHEET_URL_EXP else None

    periodos_imp = merge_anos(raw_imp, leer_hoja_import, IDS_IMP, hoy)
    periodos_exp = merge_anos(raw_exp, leer_hoja_export, IDS_EXP, hoy) if raw_exp else {}

    # Generar datos para dashboard
    gruas_data, periodo_actual = generar_dashboard_data(periodos_imp, periodos_exp, hoy)
    
    # Crear directorio si no existe
    os.makedirs("docs", exist_ok=True)
    
    # Guardar JSON
    with open("docs/data.json", "w", encoding="utf-8") as f:
        json.dump({
            "data": gruas_data,
            "periodo_actual": periodo_actual,
            "tiene_export": bool(periodos_exp)
        }, f, ensure_ascii=False, indent=2)
    
    print(f"✅ Dashboard data generado: {len(gruas_data)} períodos")
    print(f"📅 Período actual: {periodo_actual}")
    print(f"✈️ Tiene export: {bool(periodos_exp)}")
