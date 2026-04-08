import os, sys, json, requests, pandas as pd
from datetime import datetime, date
from io import BytesIO

# ── Configuración ──────────────────────────────────────────────────────────
LIMIT_HRS     = 160
SHEET_URL_IMP = os.environ.get("SHEET_URL_IMPORT", "")
SHEET_URL_EXP = os.environ.get("SHEET_URL_EXPORT", "")

MESES_ES = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
            7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}

# ── FLOTA IMPORT: solo Linde (todas BAJA) — Royal en Excel pero no se renderiza
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

COLORS_IMP = {
    "LINDE 11728":"#0055B3","LINDE 11731":"#0066CC","LINDE 11732":"#0077DD",
    "LINDE 11733":"#0088EE","LINDE 11734":"#0099CC","LINDE 11735":"#00AADD",
    "LINDE 11736":"#00BBEE","LINDE 11738":"#33AADD","LINDE 11739":"#55BBEE",
}
COLORS_EXP = {
    "EXP 11720":"#0055B3","EXP 11721":"#0066CC","EXP 11722":"#0077DD",
    "EXP 11723":"#0088EE","EXP 11724":"#0099CC","EXP 11725":"#00AADD",
    "EXP 11726":"#00BBEE","EXP 11727":"#33AADD","EXP 11729":"#55BBEE",
    "EXP 11730":"#1B3A6B","EXP 11737":"#2A4A7A","EXP 11740":"#3A5A8A",
}

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

def _parse_val(val):
    if pd.isna(val): return None
    def _parse_val(val):
    if pd.isna(val): return None
    if isinstance(val, str):
        v = val.strip()
        try:
            return float(v) if v else None
        except ValueError:
            return None
    try: return float(val)
    except: return None
def leer_hoja_import(excel_bytes, year):
    """col0=Nº, col1=Fecha, cols2-7=Royal(skip), cols8-16=Linde(9 grúas)"""
    nombre = f"SEMANAS {year}"
    try:
        df = pd.read_excel(excel_bytes, sheet_name=nombre, header=None,
                           skiprows=5, usecols=range(18))
    except Exception as e:
        print(f"  [WARN] '{nombre}' Import: {e}"); return []
    rows = []
    for _, row in df.iterrows():
        sem = row.iloc[0]
        if pd.isna(sem) or not str(sem).strip().isdigit(): continue
        if int(sem) > 57: continue
        try: fecha = pd.to_datetime(row.iloc[1]).date()
        except: continue
        entry = {"sem": int(sem), "fecha": fecha}
        for i, g in enumerate(GRUAS_IMPORT):
            entry[g["id"]] = _parse_val(row.iloc[8 + i])
        rows.append(entry)
    return rows

def leer_hoja_export(excel_bytes, year):
    """col0=Nº, col1=Fecha, cols2-13=12 grúas export"""
    nombre = f"SEMANAS {year}"
    try:
        df = pd.read_excel(excel_bytes, sheet_name=nombre, header=None,
                           skiprows=5, usecols=range(14))
    except Exception as e:
        print(f"  [WARN] '{nombre}' Export: {e}"); return []
    rows = []
    for _, row in df.iterrows():
        sem = row.iloc[0]
        if pd.isna(sem) or not str(sem).strip().isdigit(): continue
        if int(sem) > 57: continue
        try: fecha = pd.to_datetime(row.iloc[1]).date()
        except: continue
        entry = {"sem": int(sem), "fecha": fecha}
        for i, g in enumerate(GRUAS_EXPORT):
            entry[g["id"]] = _parse_val(row.iloc[2 + i])
        rows.append(entry)
    return rows

def calcular_horas_semanales(rows, grua_ids):
    prev = {gid: None for gid in grua_ids}
    for row in rows:
        for gid in grua_ids:
            val = row.get(gid)
            if isinstance(val, float) and prev[gid] is not None:
                row[f"{gid}_hrs"] = round(max(val - prev[gid], 0), 1)
            else:
                row[f"{gid}_hrs"] = None
            if isinstance(val, float):
                prev[gid] = val
    return rows

def periodo_key_label(fecha):
    if fecha.day >= 20:
        inicio = date(fecha.year, fecha.month, 20)
        fin = date(fecha.year + 1, 1, 20) if fecha.month == 12 else date(fecha.year, fecha.month + 1, 20)
    else:
        inicio = date(fecha.year - 1, 12, 20) if fecha.month == 1 else date(fecha.year, fecha.month - 1, 20)
        fin = date(fecha.year, fecha.month, 20)
    key   = f"{inicio.strftime('%Y%m%d')}_{fin.strftime('%Y%m%d')}"
    label = f"20 {MESES_ES[inicio.month]} – 20 {MESES_ES[fin.month]} {fin.year}"
    return key, label, inicio, fin

def agrupar_por_periodo(rows, grua_ids, hoy):
    periodos = {}
    for row in rows:
        fecha = row["fecha"]
        if fecha > hoy: continue                       # ignorar fechas futuras
        key, label, inicio, fin = periodo_key_label(fecha)
        if inicio > hoy: continue                      # ignorar períodos futuros
        if key not in periodos:
            periodos[key] = {
                "key": key, "label": label,
                "inicio": inicio, "fin": fin,
                "semanas": [],
                "hrsporgid":  {gid: 0.0  for gid in grua_ids},
                "tiene_dato": {gid: False for gid in grua_ids},
            }
        periodos[key]["semanas"].append(fecha.strftime("%d/%m"))
        for gid in grua_ids:
            hrs = row.get(f"{gid}_hrs")
            if hrs is not None:
                periodos[key]["hrsporgid"][gid]  += hrs
                periodos[key]["tiene_dato"][gid]  = True
    return periodos

def merge_anos(excel_bytes, leer_fn, grua_ids, hoy):
    total = {}
    for year in [2024, 2025, 2026]:
        excel_bytes.seek(0)
        rows = leer_fn(excel_bytes, year)
        if not rows: continue
        rows = calcular_horas_semanales(rows, grua_ids)
        for k, v in agrupar_por_periodo(rows, grua_ids, hoy).items():
            if k not in total:
                total[k] = v
            else:
                for gid in grua_ids:
                    if v["tiene_dato"][gid]:
                        total[k]["hrsporgid"][gid] += v["hrsporgid"][gid]
                        total[k]["tiene_dato"][gid] = True
                for s in v["semanas"]:
                    if s not in total[k]["semanas"]:
                        total[k]["semanas"].append(s)
    return total

def get_status(hrs, tiene_dato):
    if not tiene_dato: return {"key":"sin_dato","label":"Sin dato","cls":"no-data"}
    pct = hrs / LIMIT_HRS
    if hrs >= LIMIT_HRS: return {"key":"limit","label":"Límite Superado","cls":"limit"}
    if pct >= 0.86:      return {"key":"alert","label":"Alerta","cls":"alert"}
    if pct >= 0.60:      return {"key":"precaution","label":"Precaución","cls":"precaution"}
    return                      {"key":"ok","label":"OK","cls":"ok"}

def build_card(gid, empresa, hrs, tiene_dato, idx):
    st    = get_status(hrs, tiene_dato)
    pct   = min(hrs / LIMIT_HRS * 100, 100) if tiene_dato else 0
    disp  = max(LIMIT_HRS - hrs, 0) if tiene_dato else LIMIT_HRS
    val   = f"{hrs:.1f}" if tiene_dato else "—"
    delay = f"{idx * 0.04:.2f}s"
    name  = gid.replace("EXP ","").replace("LINDE ","")
    return f"""<div class="crane-card s-{st['cls']}" style="animation-delay:{delay}">
  <div class="crane-header">
    <div class="crane-name">{name}</div>
    <div class="status-badge {st['cls']}">● {st['label']}</div>
  </div>
  <div class="crane-plate">BAJA · {empresa}</div>
  <div class="crane-km">
    <span class="crane-km-val">{val}</span>
    <span class="crane-km-of">/ {LIMIT_HRS} hrs</span>
  </div>
  <div class="prog-bar"><div class="prog-fill {st['cls']}" style="width:{pct:.0f}%"></div></div>
  <div class="crane-footer">
    <span class="crane-pct">{pct:.0f}% del límite</span>
    <span class="disp">{disp:.0f} hrs disp.</span>
  </div>
</div>"""

def build_entry(p, gruas, colors):
    cards = ""
    ok = prec = alert = limit = 0
    bar_labels, bar_data, bar_colors = [], [], []
    for i, g in enumerate(gruas):
        gid   = g["id"]
        hrs   = p["hrsporgid"][gid]
        tiene = p["tiene_dato"][gid]
        st    = get_status(hrs, tiene)
        cards += build_card(gid, g["empresa"], hrs, tiene, i)
        bar_labels.append(gid.replace("EXP ","").replace("LINDE ",""))
        bar_data.append(round(hrs, 1) if tiene else 0)
        bar_colors.append(colors.get(gid, "#0099D6"))
        if   st["key"] == "ok":         ok    += 1
        elif st["key"] == "precaution": prec  += 1
        elif st["key"] == "alert":      alert += 1
        elif st["key"] == "limit":      limit += 1
    return {
        "cards":      cards,
        "ok": ok, "prec": prec, "alert": alert, "limit": limit,
        "bar_labels": bar_labels,
        "bar_data":   bar_data,
        "bar_colors": bar_colors,
        "donut":      [ok, prec, alert, limit],
        "n_sem":      len(p["semanas"]),
    }

# ── Main ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not SHEET_URL_IMP:
        print("ERROR: Variable SHEET_URL_IMPORT no configurada."); sys.exit(1)

    now = datetime.now()
    hoy = now.date()

    raw_imp = download_excel(SHEET_URL_IMP, "IMPORTACIONES")
    raw_exp = download_excel(SHEET_URL_EXP, "EXPORTACIONES") if SHEET_URL_EXP else None

    periodos_imp = merge_anos(raw_imp, leer_hoja_import, IDS_IMP, hoy)
    periodos_exp = merge_anos(raw_exp, leer_hoja_export, IDS_EXP, hoy) if raw_exp else {}

    all_keys = set(list(periodos_imp.keys()) + list(periodos_exp.keys()))
    if not all_keys:
        print("ERROR: No se encontraron datos."); sys.exit(1)

    gruas_js = {}
    for key in all_keys:
        p_imp = periodos_imp.get(key)
        p_exp = periodos_exp.get(key)
        has_imp = p_imp and any(p_imp["tiene_dato"].values())
        has_exp = p_exp and any(p_exp["tiene_dato"].values())
        if not has_imp and not has_exp:
            continue
        base = p_imp or p_exp
        gruas_js[key] = {
            "label":       base["label"],
            "inicio_label": f"20 {MESES_ES[base['inicio'].month]} {base['inicio'].year}",
            "fin_label":    f"20 {MESES_ES[base['fin'].month]} {base['fin'].year}",
            "imp": build_entry(p_imp, GRUAS_IMPORT, COLORS_IMP) if has_imp else None,
            "exp": build_entry(p_exp, GRUAS_EXPORT, COLORS_EXP) if has_exp else None,
        }

    keys_sorted = sorted(gruas_js.keys(), reverse=True)
    actual_key, _, _, _ = periodo_key_label(hoy)
    if actual_key not in gruas_js:
        actual_key = keys_sorted[0]

    periodo_opts = ""
    for key in keys_sorted:
        sel   = "selected" if key == actual_key else ""
        gruas_js[key]["label"]
        periodo_opts += f'<option value="{key}" {"selected" if key == actual_key else ""}>{gruas_js[key]["label"]}</option>'

    print(f"Períodos procesados: {len(gruas_js)}")
    print(f"Mostrando: {gruas_js[actual_key]['label']}")

    with open("template.html", "r", encoding="utf-8") as f:
        t = f.read()

    html = t \
        .replace("{{HORA}}",            now.strftime("%H:%M")) \
        .replace("{{FECHA}}",           now.strftime("%d/%m/%Y")) \
        .replace("{{PERIODO_OPTIONS}}", periodo_opts) \
        .replace("{{GRUAS_JS_DATA}}",   json.dumps(gruas_js, ensure_ascii=False)) \
        .replace("{{PERIODO_ACTUAL}}",  actual_key) \
        .replace("{{TIENE_EXPORT}}",    "true" if raw_exp else "false")

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente.")
