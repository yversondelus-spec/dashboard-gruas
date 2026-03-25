import os, sys, json, requests, pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO

# ── Configuración ──────────────────────────────────────────────────────────
LIMIT_HRS      = 160
SHEET_URL_IMP  = os.environ.get("SHEET_URL_IMPORT", "")
SHEET_URL_EXP  = os.environ.get("SHEET_URL_EXPORT", "")

MESES_ES = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
            7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
MESES_FULL = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
              7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

# Linde = ALTAS (control crítico) — van primero y con más protagonismo
# Royal = BAJAS (costo fijo)      — van en sección secundaria
GRUAS_LINDE = [
    {"id":"LINDE 11728", "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11731", "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11732", "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11733", "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11734", "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11735", "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11736", "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11738", "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11739", "tipo":"ALTA", "empresa":"Linde Leasing"},
]
GRUAS_ROYAL = [
    {"id":"ROYAL 9023",  "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9024",  "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9025",  "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9026",  "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9027",  "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9028",  "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"TRACTOR EQUIPAJE","tipo":"MULA","empresa":"—"},
]
GRUAS_ALL = GRUAS_LINDE + GRUAS_ROYAL
GRUA_IDS  = [g["id"] for g in GRUAS_ALL]

# Colores corporativos Aerosan Airport Services
# Azul institucional + paleta de estado clara
AEROSAN_BLUE   = "#003087"
AEROSAN_BLUE2  = "#0055B3"
AEROSAN_CYAN   = "#0099CC"
AEROSAN_GRAY   = "#4A5568"

GRUA_COLORS_LINDE = {
    "LINDE 11728":"#0055B3","LINDE 11731":"#0066CC","LINDE 11732":"#0077DD",
    "LINDE 11733":"#0088EE","LINDE 11734":"#0099CC","LINDE 11735":"#00AADD",
    "LINDE 11736":"#00BBEE","LINDE 11738":"#33AADD","LINDE 11739":"#55BBEE",
}
GRUA_COLORS_ROYAL = {
    "ROYAL 9023":"#94A3B8","ROYAL 9024":"#A0AEC0","ROYAL 9025":"#8896A8",
    "ROYAL 9026":"#7A8898","ROYAL 9027":"#6B7A8A","ROYAL 9028":"#5C6B7A",
    "TRACTOR EQUIPAJE":"#E67E22",
}
GRUA_COLORS = {**GRUA_COLORS_LINDE, **GRUA_COLORS_ROYAL}

def download_excel(url, label=""):
    if not url:
        return None
    if "docs.google.com/spreadsheets" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    print(f"Descargando Excel {label}...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return BytesIO(r.content)

def leer_hoja_anual(excel_bytes, year):
    """Lee hoja 'SEMANAS {year}', devuelve lista de rows con fecha y horómetros."""
    nombre = f"SEMANAS {year}"
    try:
        df = pd.read_excel(excel_bytes, sheet_name=nombre, header=None,
                           skiprows=5, usecols=range(19))
    except Exception as e:
        print(f"  [WARN] '{nombre}' no encontrada: {e}")
        return []

    rows = []
    for _, row in df.iterrows():
        sem = row.iloc[0]
        if pd.isna(sem) or not str(sem).strip().isdigit():
            continue
        sem = int(sem)
        if sem > 57:
            continue
        try:
            fecha = pd.to_datetime(row.iloc[1]).date()
        except Exception:
            continue

        entry = {"sem": sem, "fecha": fecha}
        for i, gid in enumerate(GRUA_IDS):
            val = row.iloc[2 + i]
            if pd.isna(val):
                entry[gid] = None
            elif isinstance(val, str):
                entry[gid] = val.strip()
            else:
                try:
                    entry[gid] = float(val)
                except Exception:
                    entry[gid] = None
        rows.append(entry)
    return rows

def calcular_horas_semanales(rows):
    """Horómetro acumulado → horas semana = lectura_actual - lectura_anterior."""
    prev = {gid: None for gid in GRUA_IDS}
    for row in rows:
        for gid in GRUA_IDS:
            val = row[gid]
            if isinstance(val, (int, float)) and prev[gid] is not None and isinstance(prev[gid], (int, float)):
                row[f"{gid}_hrs"] = round(max(val - prev[gid], 0), 1)
            else:
                row[f"{gid}_hrs"] = None
            if isinstance(val, (int, float)):
                prev[gid] = val
    return rows

def periodo_key_label(fecha):
    """
    Determina el período 20→20 al que pertenece una fecha.
    Retorna (key, label, inicio, fin).
    día >= 20 → período empieza este mes
    día <  20 → período empezó el mes anterior
    """
    if fecha.day >= 20:
        inicio = date(fecha.year, fecha.month, 20)
        if fecha.month == 12:
            fin = date(fecha.year + 1, 1, 20)
        else:
            fin = date(fecha.year, fecha.month + 1, 20)
    else:
        if fecha.month == 1:
            inicio = date(fecha.year - 1, 12, 20)
        else:
            inicio = date(fecha.year, fecha.month - 1, 20)
        fin = date(fecha.year, fecha.month, 20)

    key   = f"{inicio.strftime('%Y%m%d')}_{fin.strftime('%Y%m%d')}"
    label = f"20 {MESES_ES[inicio.month]} – 20 {MESES_ES[fin.month]} {fin.year}"
    return key, label, inicio, fin

def agrupar_por_periodo(rows):
    """Agrupa horas semanales por período 20→20."""
    periodos = {}
    for row in rows:
        key, label, inicio, fin = periodo_key_label(row["fecha"])
        if key not in periodos:
            periodos[key] = {
                "key": key, "label": label,
                "inicio": inicio, "fin": fin,
                "semanas": [],
                "hrsporgid":  {gid: 0.0  for gid in GRUA_IDS},
                "tiene_dato": {gid: False for gid in GRUA_IDS},
            }
        periodos[key]["semanas"].append(row["fecha"].strftime("%d/%m"))
        for gid in GRUA_IDS:
            hrs = row.get(f"{gid}_hrs")
            if hrs is not None and isinstance(hrs, float):
                periodos[key]["hrsporgid"][gid]  += hrs
                periodos[key]["tiene_dato"][gid]  = True
    return periodos

def sumar_periodos(periodos_imp, periodos_exp):
    """Suma horas de importación + exportación por el mismo período."""
    todas_keys = set(list(periodos_imp.keys()) + list(periodos_exp.keys()))
    resultado  = {}
    for key in todas_keys:
        p_imp = periodos_imp.get(key)
        p_exp = periodos_exp.get(key)
        base  = p_imp if p_imp else p_exp

        merged = {
            "key":       base["key"],
            "label":     base["label"],
            "inicio":    base["inicio"],
            "fin":       base["fin"],
            "semanas":   base["semanas"],
            "hrsporgid": {},
            "tiene_dato":{},
            "hrs_imp":   {},
            "hrs_exp":   {},
        }
        for gid in GRUA_IDS:
            hi = p_imp["hrsporgid"][gid] if p_imp and p_imp["tiene_dato"][gid] else 0.0
            he = p_exp["hrsporgid"][gid] if p_exp and p_exp["tiene_dato"][gid] else 0.0
            tiene = (p_imp and p_imp["tiene_dato"][gid]) or (p_exp and p_exp["tiene_dato"][gid])
            merged["hrsporgid"][gid]  = round(hi + he, 1)
            merged["tiene_dato"][gid] = tiene
            merged["hrs_imp"][gid]    = round(hi, 1)
            merged["hrs_exp"][gid]    = round(he, 1)
        resultado[key] = merged
    return resultado

def get_status(hrs, tiene_dato):
    if not tiene_dato:
        return {"key":"sin_dato",   "label":"Sin dato",        "cls":"no-data"}
    pct = hrs / LIMIT_HRS
    if hrs >= LIMIT_HRS:
        return {"key":"limit",      "label":"Límite Superado", "cls":"limit"}
    if pct >= 0.86:
        return {"key":"alert",      "label":"Alerta",          "cls":"alert"}
    if pct >= 0.60:
        return {"key":"precaution", "label":"Precaución",      "cls":"precaution"}
    return     {"key":"ok",         "label":"OK",              "cls":"ok"}

def build_card(g, hrs_total, hrs_imp, hrs_exp, tiene_dato, idx, compact=False):
    gid  = g["id"]
    st   = get_status(hrs_total, tiene_dato)
    pct  = min(hrs_total / LIMIT_HRS * 100, 100) if tiene_dato else 0
    disp = max(LIMIT_HRS - hrs_total, 0) if tiene_dato else LIMIT_HRS
    val  = f"{hrs_total:.1f}" if tiene_dato else "—"
    imp_lbl = f"IMP {hrs_imp:.1f}" if tiene_dato else ""
    exp_lbl = f"EXP {hrs_exp:.1f}" if tiene_dato else ""

    size_cls = " compact" if compact else ""
    delay    = f"{idx * 0.04:.2f}s"

    return f"""<div class="crane-card s-{st['cls']}{size_cls}" style="animation-delay:{delay}">
  <div class="crane-header">
    <div class="crane-name">{gid}</div>
    <div class="status-badge {st['cls']}">● {st['label']}</div>
  </div>
  <div class="crane-plate">{g['tipo']} · {g['empresa']}</div>
  <div class="crane-km">
    <span class="crane-km-val">{val}</span>
    <span class="crane-km-of">/ {LIMIT_HRS} hrs</span>
  </div>
  <div class="prog-bar"><div class="prog-fill {st['cls']}" style="width:{pct:.0f}%"></div></div>
  <div class="crane-footer">
    <span class="imp-exp">{imp_lbl} · {exp_lbl}</span>
    <span class="disp">{disp:.0f} hrs disp.</span>
  </div>
</div>"""

def build_periodo_entry(p):
    # ── Linde (ALTAS — críticas) ──────────────────────────────────────────
    cards_linde = ""
    ok_l = prec_l = alert_l = limit_l = sin_l = 0
    bar_l_labels, bar_l_imp, bar_l_exp = [], [], []

    for i, g in enumerate(GRUAS_LINDE):
        gid        = g["id"]
        hrs        = p["hrsporgid"][gid]
        tiene      = p["tiene_dato"][gid]
        hi, he     = p["hrs_imp"][gid], p["hrs_exp"][gid]
        st         = get_status(hrs, tiene)
        cards_linde += build_card(g, hrs, hi, he, tiene, i, compact=False)
        bar_l_labels.append(gid.replace("LINDE ",""))
        bar_l_imp.append(hi if tiene else 0)
        bar_l_exp.append(he if tiene else 0)
        if   st["key"]=="ok":         ok_l    +=1
        elif st["key"]=="precaution": prec_l  +=1
        elif st["key"]=="alert":      alert_l +=1
        elif st["key"]=="limit":      limit_l +=1
        else:                         sin_l   +=1

    # ── Royal + Tractor (BAJAS — costo fijo) ─────────────────────────────
    cards_royal = ""
    ok_r = prec_r = alert_r = limit_r = 0
    bar_r_labels, bar_r_data, bar_r_colors = [], [], []

    for i, g in enumerate(GRUAS_ROYAL):
        gid    = g["id"]
        hrs    = p["hrsporgid"][gid]
        tiene  = p["tiene_dato"][gid]
        hi, he = p["hrs_imp"][gid], p["hrs_exp"][gid]
        st     = get_status(hrs, tiene)
        cards_royal += build_card(g, hrs, hi, he, tiene, i, compact=True)
        bar_r_labels.append(gid.replace("ROYAL ","").replace("TRACTOR EQUIPAJE","TRACTOR"))
        bar_r_data.append(round(hrs,1) if tiene else 0)
        bar_r_colors.append(GRUA_COLORS.get(gid,"#94A3B8"))
        if   st["key"]=="ok":         ok_r    +=1
        elif st["key"]=="precaution": prec_r  +=1
        elif st["key"]=="alert":      alert_r +=1
        elif st["key"]=="limit":      limit_r +=1

    n_sem = len(p["semanas"])
    return {
        "label":   p["label"],
        "n_sem":   n_sem,
        # KPIs Linde
        "ok_l":    ok_l,  "prec_l": prec_l, "alert_l": alert_l, "limit_l": limit_l,
        # KPIs Royal
        "ok_r":    ok_r,  "prec_r": prec_r, "alert_r": alert_r, "limit_r": limit_r,
        # Cards
        "cards_linde": cards_linde,
        "cards_royal": cards_royal,
        # Chart Linde stacked (IMP + EXP)
        "bar_l_labels": bar_l_labels,
        "bar_l_imp":    bar_l_imp,
        "bar_l_exp":    bar_l_exp,
        # Chart Royal simple
        "bar_r_labels": bar_r_labels,
        "bar_r_data":   bar_r_data,
        "bar_r_colors": bar_r_colors,
        # Donut linde
        "donut_l": [ok_l, prec_l, alert_l, limit_l],
    }

def make_selectors(data_keys_sorted, actual_key):
    """data_keys_sorted = lista de keys ordenadas más reciente primero."""
    opts = ""
    for key in data_keys_sorted:
        # Obtener label del primer entry que tenga esta key
        sel = "selected" if key == actual_key else ""
        opts += f'<option value="{key}" {sel}></option>'  # label se inyecta después
    return opts

# ── Main ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not SHEET_URL_IMP:
        print("ERROR: Variable SHEET_URL_IMPORT no configurada."); sys.exit(1)

    now        = datetime.now()
    hoy        = now.date()

    # Descargar ambos Excel
    raw_imp = download_excel(SHEET_URL_IMP, "IMPORTACIONES")
    raw_exp = download_excel(SHEET_URL_EXP, "EXPORTACIONES") if SHEET_URL_EXP else None

    # Leer y calcular horas por semana para cada año
    periodos_imp_total = {}
    periodos_exp_total = {}

    for year in [2024, 2025, 2026]:
        # Importaciones
        raw_imp.seek(0)
        rows_imp = leer_hoja_anual(raw_imp, year)
        if rows_imp:
            rows_imp = calcular_horas_semanales(rows_imp)
            for k, v in agrupar_por_periodo(rows_imp).items():
                periodos_imp_total[k] = v

        # Exportaciones
        if raw_exp:
            raw_exp.seek(0)
            rows_exp = leer_hoja_anual(raw_exp, year)
            if rows_exp:
                rows_exp = calcular_horas_semanales(rows_exp)
                for k, v in agrupar_por_periodo(rows_exp).items():
                    periodos_exp_total[k] = v

    # Si no hay exportaciones, usar ceros
    if not periodos_exp_total:
        print("  [INFO] Sin datos de exportaciones — se mostrará solo importaciones.")
        for k, v in periodos_imp_total.items():
            periodos_exp_total[k] = {
                "key": v["key"], "label": v["label"],
                "inicio": v["inicio"], "fin": v["fin"],
                "semanas": [], "hrsporgid": {gid: 0.0 for gid in GRUA_IDS},
                "tiene_dato": {gid: False for gid in GRUA_IDS},
            }

    # Combinar IMP + EXP
    periodos_merged = sumar_periodos(periodos_imp_total, periodos_exp_total)

    if not periodos_merged:
        print("ERROR: No se encontraron datos."); sys.exit(1)

    # Construir entries para el JS
    gruas_js = {}
    for key, p in periodos_merged.items():
        if any(p["tiene_dato"].values()):
            gruas_js[key] = build_periodo_entry(p)
            gruas_js[key]["label"] = p["label"]  # asegurar label

    # Keys ordenadas más reciente primero
    keys_sorted = sorted(gruas_js.keys(), reverse=True)

    # Período actual
    actual_key, actual_label, _, _ = periodo_key_label(hoy)
    if actual_key not in gruas_js:
        actual_key = keys_sorted[0]

    # Generar options con labels reales
    periodo_opts = ""
    for key in keys_sorted:
        sel   = "selected" if key == actual_key else ""
        label = gruas_js[key]["label"]
        periodo_opts += f'<option value="{key}" {sel}>{label}</option>'

    print(f"Períodos procesados: {len(gruas_js)}")
    print(f"Mostrando: {gruas_js[actual_key]['label']}")

    with open("template.html", "r", encoding="utf-8") as f:
        t = f.read()

    html = t \
        .replace("{{HORA}}",             now.strftime("%H:%M")) \
        .replace("{{FECHA}}",            now.strftime("%d/%m/%Y")) \
        .replace("{{PERIODO_OPTIONS}}",  periodo_opts) \
        .replace("{{GRUAS_JS_DATA}}",    json.dumps(gruas_js, ensure_ascii=False)) \
        .replace("{{PERIODO_ACTUAL}}",   actual_key) \
        .replace("{{TIENE_EXPORT}}",     "true" if SHEET_URL_EXP else "false")

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente.")
