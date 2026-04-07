import os, sys, json, requests, pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO

# ── Configuración ──────────────────────────────────────────────────────────
LIMIT_HRS      = 160
SHEET_URL_IMP  = os.environ.get("SHEET_URL_IMPORT", "")
SHEET_URL_EXP  = os.environ.get("SHEET_URL_EXPORT", "")

MESES_ES = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
            7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}

# ── FLOTA IMPORT: Royal (BAJAS) → Linde (ALTAS) → Tractor ─────────────────
GRUAS_IMPORT = [
    {"id":"ROYAL 9023",       "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9024",       "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9025",       "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9026",       "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9027",       "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9028",       "tipo":"BAJA", "empresa":"Royal Leasing"},
    {"id":"LINDE 11728",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11731",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11732",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11733",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11734",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11735",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11736",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11738",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11739",      "tipo":"ALTA", "empresa":"Linde Leasing"},
    {"id":"TRACTOR EQUIPAJE", "tipo":"MULA", "empresa":"—"},
]
IDS_IMP = [g["id"] for g in GRUAS_IMPORT]

# ── FLOTA EXPORT: todas BAJAS ──────────────────────────────────────────────
GRUAS_EXPORT = [
    {"id":"EXP 11720", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11721", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11722", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11723", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11724", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11725", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11726", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11727", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11729", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11730", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11737", "tipo":"BAJA", "empresa":"Export Leasing"},
    {"id":"EXP 11740", "tipo":"BAJA", "empresa":"Export Leasing"},
]
IDS_EXP = [g["id"] for g in GRUAS_EXPORT]

# ── Colores ────────────────────────────────────────────────────────────────
COLORS_IMP = {
    "ROYAL 9023":"#94A3B8","ROYAL 9024":"#A0AEC0","ROYAL 9025":"#8896A8",
    "ROYAL 9026":"#7A8898","ROYAL 9027":"#6B7A8A","ROYAL 9028":"#5C6B7A",
    "LINDE 11728":"#0055B3","LINDE 11731":"#0066CC","LINDE 11732":"#0077DD",
    "LINDE 11733":"#0088EE","LINDE 11734":"#0099CC","LINDE 11735":"#00AADD",
    "LINDE 11736":"#00BBEE","LINDE 11738":"#33AADD","LINDE 11739":"#55BBEE",
    "TRACTOR EQUIPAJE":"#E67E22",
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

def leer_hoja(excel_bytes, year, grua_ids, n_cols):
    """Lee una hoja anual para una flota dada.
    n_cols = número total de columnas a leer (2 fijas + len(grua_ids))
    """
    nombre = f"SEMANAS {year}"
    try:
        df = pd.read_excel(excel_bytes, sheet_name=nombre, header=None,
                           skiprows=5, usecols=range(n_cols))
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
        for i, gid in enumerate(grua_ids):
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

def calcular_horas_semanales(rows, grua_ids):
    prev = {gid: None for gid in grua_ids}
    for row in rows:
        for gid in grua_ids:
            val = row[gid]
            if isinstance(val, (int, float)) and prev[gid] is not None and isinstance(prev[gid], (int, float)):
                row[f"{gid}_hrs"] = round(max(val - prev[gid], 0), 1)
            else:
                row[f"{gid}_hrs"] = None
            if isinstance(val, (int, float)):
                prev[gid] = val
    return rows

def periodo_key_label(fecha):
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

def agrupar_por_periodo(rows, grua_ids):
    periodos = {}
    for row in rows:
        key, label, inicio, fin = periodo_key_label(row["fecha"])
        if key not in periodos:
            periodos[key] = {
                "key": key, "label": label,
                "inicio": inicio, "fin": fin,
                "semanas": [],
                "hrsporgid":  {gid: 0.0  for gid in grua_ids},
                "tiene_dato": {gid: False for gid in grua_ids},
            }
        periodos[key]["semanas"].append(row["fecha"].strftime("%d/%m"))
        for gid in grua_ids:
            hrs = row.get(f"{gid}_hrs")
            if hrs is not None and isinstance(hrs, float):
                periodos[key]["hrsporgid"][gid]  += hrs
                periodos[key]["tiene_dato"][gid]  = True
    return periodos

def merge_anos(excel_bytes, grua_ids, n_cols):
    """Lee todos los años disponibles y fusiona periodos."""
    total = {}
    for year in [2024, 2025, 2026]:
        excel_bytes.seek(0)
        rows = leer_hoja(excel_bytes, year, grua_ids, n_cols)
        if rows:
            rows = calcular_horas_semanales(rows, grua_ids)
            for k, v in agrupar_por_periodo(rows, grua_ids).items():
                if k not in total:
                    total[k] = v
                else:
                    # sumar horas si el periodo ya existe (cruce de años)
                    for gid in grua_ids:
                        if v["tiene_dato"][gid]:
                            total[k]["hrsporgid"][gid] += v["hrsporgid"][gid]
                            total[k]["tiene_dato"][gid] = True
    return total

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

def build_card(g, hrs, tiene_dato, idx, compact=False):
    gid  = g["id"]
    st   = get_status(hrs, tiene_dato)
    pct  = min(hrs / LIMIT_HRS * 100, 100) if tiene_dato else 0
    disp = max(LIMIT_HRS - hrs, 0) if tiene_dato else LIMIT_HRS
    val  = f"{hrs:.1f}" if tiene_dato else "—"
    size_cls = " compact" if compact else ""
    delay    = f"{idx * 0.04:.2f}s"
    label_display = gid.replace("EXP ", "")

    return f"""<div class="crane-card s-{st['cls']}{size_cls}" style="animation-delay:{delay}">
  <div class="crane-header">
    <div class="crane-name">{label_display}</div>
    <div class="status-badge {st['cls']}">● {st['label']}</div>
  </div>
  <div class="crane-plate">{g['tipo']} · {g['empresa']}</div>
  <div class="crane-km{' compact' if compact else ''}">
    <span class="crane-km-val">{val}</span>
    <span class="crane-km-of">/ {LIMIT_HRS} hrs</span>
  </div>
  <div class="prog-bar"><div class="prog-fill {st['cls']}" style="width:{pct:.0f}%"></div></div>
  <div class="crane-footer">
    <span class="imp-exp">{g['empresa']}</span>
    <span class="disp">{disp:.0f} hrs disp.</span>
  </div>
</div>"""

def build_periodo_entry(p_imp, p_exp):
    """Construye el entry del período con secciones separadas IMP y EXP."""
    # ── Sección IMPORT ─────────────────────────────────────────────────────
    # Linde (ALTAS)
    gruas_linde = [g for g in GRUAS_IMPORT if g["tipo"] == "ALTA"]
    gruas_royal = [g for g in GRUAS_IMPORT if g["tipo"] in ("BAJA","MULA")]

    cards_linde = ""
    ok_l = prec_l = alert_l = limit_l = 0
    bar_l_labels, bar_l_data = [], []

    for i, g in enumerate(gruas_linde):
        gid   = g["id"]
        hrs   = p_imp["hrsporgid"][gid] if p_imp else 0.0
        tiene = p_imp["tiene_dato"][gid] if p_imp else False
        st    = get_status(hrs, tiene)
        cards_linde += build_card(g, hrs, tiene, i, compact=False)
        bar_l_labels.append(gid.replace("LINDE ",""))
        bar_l_data.append(round(hrs,1) if tiene else 0)
        if   st["key"]=="ok":         ok_l    +=1
        elif st["key"]=="precaution": prec_l  +=1
        elif st["key"]=="alert":      alert_l +=1
        elif st["key"]=="limit":      limit_l +=1

    cards_royal = ""
    ok_r = prec_r = alert_r = limit_r = 0
    bar_r_labels, bar_r_data, bar_r_colors = [], [], []

    for i, g in enumerate(gruas_royal):
        gid   = g["id"]
        hrs   = p_imp["hrsporgid"][gid] if p_imp else 0.0
        tiene = p_imp["tiene_dato"][gid] if p_imp else False
        st    = get_status(hrs, tiene)
        cards_royal += build_card(g, hrs, tiene, i, compact=True)
        lbl = gid.replace("ROYAL ","").replace("TRACTOR EQUIPAJE","TRACTOR")
        bar_r_labels.append(lbl)
        bar_r_data.append(round(hrs,1) if tiene else 0)
        bar_r_colors.append(COLORS_IMP.get(gid,"#94A3B8"))
        if   st["key"]=="ok":         ok_r    +=1
        elif st["key"]=="precaution": prec_r  +=1
        elif st["key"]=="alert":      alert_r +=1
        elif st["key"]=="limit":      limit_r +=1

    # ── Sección EXPORT ────────────────────────────────────────────────────
    cards_exp = ""
    ok_e = prec_e = alert_e = limit_e = 0
    bar_e_labels, bar_e_data, bar_e_colors = [], [], []

    for i, g in enumerate(GRUAS_EXPORT):
        gid   = g["id"]
        hrs   = p_exp["hrsporgid"][gid] if p_exp else 0.0
        tiene = p_exp["tiene_dato"][gid] if p_exp else False
        st    = get_status(hrs, tiene)
        cards_exp += build_card(g, hrs, tiene, i, compact=False)
        bar_e_labels.append(gid.replace("EXP ",""))
        bar_e_data.append(round(hrs,1) if tiene else 0)
        bar_e_colors.append(COLORS_EXP.get(gid,"#0099D6"))
        if   st["key"]=="ok":         ok_e    +=1
        elif st["key"]=="precaution": prec_e  +=1
        elif st["key"]=="alert":      alert_e +=1
        elif st["key"]=="limit":      limit_e +=1

    n_sem = len(p_imp["semanas"]) if p_imp else (len(p_exp["semanas"]) if p_exp else 0)

    return {
        "label":   (p_imp or p_exp)["label"],
        "n_sem":   n_sem,
        # KPIs import
        "ok_l":    ok_l,  "prec_l": prec_l, "alert_l": alert_l, "limit_l": limit_l,
        "ok_r":    ok_r,  "prec_r": prec_r, "alert_r": alert_r, "limit_r": limit_r,
        # KPIs export
        "ok_e":    ok_e,  "prec_e": prec_e, "alert_e": alert_e, "limit_e": limit_e,
        # Cards HTML
        "cards_linde": cards_linde,
        "cards_royal": cards_royal,
        "cards_exp":   cards_exp,
        # Chart data import
        "bar_l_labels": bar_l_labels,
        "bar_l_data":   bar_l_data,
        "bar_r_labels": bar_r_labels,
        "bar_r_data":   bar_r_data,
        "bar_r_colors": bar_r_colors,
        # Chart data export
        "bar_e_labels": bar_e_labels,
        "bar_e_data":   bar_e_data,
        "bar_e_colors": bar_e_colors,
        # Donut
        "donut_l": [ok_l, prec_l, alert_l, limit_l],
        "donut_e": [ok_e, prec_e, alert_e, limit_e],
    }

# ── Main ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not SHEET_URL_IMP:
        print("ERROR: Variable SHEET_URL_IMPORT no configurada."); sys.exit(1)

    now = datetime.now()
    hoy = now.date()

    raw_imp = download_excel(SHEET_URL_IMP, "IMPORTACIONES")
    raw_exp = download_excel(SHEET_URL_EXP, "EXPORTACIONES") if SHEET_URL_EXP else None

    # Import: 16 grúas → cols 0..17 (2 fijas + 16)
    periodos_imp = merge_anos(raw_imp, IDS_IMP, 18) if raw_imp else {}

    # Export: 12 grúas → cols 0..13 (2 fijas + 12)
    periodos_exp = merge_anos(raw_exp, IDS_EXP, 14) if raw_exp else {}

    all_keys = set(list(periodos_imp.keys()) + list(periodos_exp.keys()))
    if not all_keys:
        print("ERROR: No se encontraron datos."); sys.exit(1)

    gruas_js = {}
    for key in all_keys:
        p_imp = periodos_imp.get(key)
        p_exp = periodos_exp.get(key)
        if p_imp and any(p_imp["tiene_dato"].values()) or \
           p_exp and any(p_exp["tiene_dato"].values()):
            gruas_js[key] = build_periodo_entry(p_imp, p_exp)

    keys_sorted = sorted(gruas_js.keys(), reverse=True)
    actual_key, _, _, _ = periodo_key_label(hoy)
    if actual_key not in gruas_js:
        actual_key = keys_sorted[0]

    periodo_opts = ""
    for key in keys_sorted:
        sel   = "selected" if key == actual_key else ""
        label = gruas_js[key]["label"]
        periodo_opts += f'<option value="{key}" {sel}>{label}</option>'

    print(f"Períodos procesados: {len(gruas_js)}")
    print(f"Mostrando: {gruas_js[actual_key]['label']}")
    tiene_export = bool(raw_exp)

    with open("template.html", "r", encoding="utf-8") as f:
        t = f.read()

    html = t \
        .replace("{{HORA}}",             now.strftime("%H:%M")) \
        .replace("{{FECHA}}",            now.strftime("%d/%m/%Y")) \
        .replace("{{PERIODO_OPTIONS}}",  periodo_opts) \
        .replace("{{GRUAS_JS_DATA}}",    json.dumps(gruas_js, ensure_ascii=False)) \
        .replace("{{PERIODO_ACTUAL}}",   actual_key) \
        .replace("{{TIENE_EXPORT}}",     "true" if tiene_export else "false")

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente.")
