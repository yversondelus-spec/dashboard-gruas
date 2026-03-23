import os, sys, json, requests, pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO

# ── Configuración ──────────────────────────────────────────────────────────────
LIMIT_HRS  = 160          # horas máximas mensuales por contrato
SHEET_URL  = os.environ.get("SHEET_URL", "")

GRUAS = [
    {"id": "ROYAL 9023",       "tipo": "ALTA",  "empresa": "Royal Leasing"},
    {"id": "ROYAL 9024",       "tipo": "ALTA",  "empresa": "Royal Leasing"},
    {"id": "ROYAL 9025",       "tipo": "ALTA",  "empresa": "Royal Leasing"},
    {"id": "ROYAL 9026",       "tipo": "ALTA",  "empresa": "Royal Leasing"},
    {"id": "ROYAL 9027",       "tipo": "ALTA",  "empresa": "Royal Leasing"},
    {"id": "ROYAL 9028",       "tipo": "ALTA",  "empresa": "Royal Leasing"},
    {"id": "LINDE 11728",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "LINDE 11731",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "LINDE 11732",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "LINDE 11733",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "LINDE 11734",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "LINDE 11735",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "LINDE 11736",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "LINDE 11738",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "LINDE 11739",      "tipo": "BAJA",  "empresa": "Linde Leasing"},
    {"id": "TRACTOR EQUIPAJE", "tipo": "MULA",  "empresa": "—"},
]
GRUA_IDS = [g["id"] for g in GRUAS]

GRUA_COLORS = {
    "ROYAL 9023": "#00b8d9", "ROYAL 9024": "#00c8e8", "ROYAL 9025": "#0098b9",
    "ROYAL 9026": "#0078a0", "ROYAL 9027": "#005a80", "ROYAL 9028": "#004060",
    "LINDE 11728": "#00e676","LINDE 11731": "#00cc60","LINDE 11732": "#00b050",
    "LINDE 11733": "#009040","LINDE 11734": "#007030","LINDE 11735": "#005020",
    "LINDE 11736": "#00e699","LINDE 11738": "#00cc80","LINDE 11739": "#00aa66",
    "TRACTOR EQUIPAJE": "#ffa726",
}

def download_excel(url):
    if "docs.google.com/spreadsheets" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    print("Descargando Excel desde Google Sheets...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return BytesIO(r.content)

def primer_lunes_iso(year):
    """Primer lunes de la semana ISO 1 del año dado."""
    jan4 = date(year, 1, 4)
    return jan4 - timedelta(days=jan4.weekday())

def leer_hoja_anual(excel_bytes, year):
    """
    Lee la hoja 'SEMANAS {year}' del Excel.
    Devuelve lista de dicts:
      { "sem": int, "fecha": date, "grua_id": valor_horometro_o_None }
    Columnas: N°(A), FECHA(B), ROYAL9023..TRACTOR(C..R), TOTAL(S)
    """
    nombre = f"SEMANAS {year}"
    try:
        df = pd.read_excel(excel_bytes, sheet_name=nombre, header=None,
                           skiprows=5, usecols=range(19))  # cols A..S (0..18)
    except Exception as e:
        print(f"  [WARN] Hoja '{nombre}' no encontrada: {e}")
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
            fecha = primer_lunes_iso(year) + timedelta(weeks=sem - 1)

        entry = {"sem": sem, "fecha": fecha}
        for i, gid in enumerate(GRUA_IDS):
            val = row.iloc[2 + i]  # cols C..R
            if pd.isna(val):
                entry[gid] = None
            elif isinstance(val, str):
                entry[gid] = val.strip()  # DETENIDA / MANTENCIÓN
            else:
                try:
                    entry[gid] = float(val)
                except Exception:
                    entry[gid] = None
        rows.append(entry)
    return rows

def calcular_horas_semana(rows):
    """
    Dado que las lecturas son ACUMULADAS, las horas de cada semana =
    lectura_actual - lectura_semana_anterior.
    Devuelve los mismos rows pero con campo '{gid}_hrs' con horas del período.
    """
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

def get_status(hrs):
    if hrs is None:
        return {"key": "sin_dato", "label": "Sin dato", "cls": "no-data"}
    pct = hrs / LIMIT_HRS
    if hrs >= LIMIT_HRS:
        return {"key": "limit",      "label": "Límite",     "cls": "limit"}
    if pct >= 0.86:
        return {"key": "alert",      "label": "Alerta",     "cls": "alert"}
    if pct >= 0.60:
        return {"key": "precaution", "label": "Precaución", "cls": "precaution"}
    return     {"key": "ok",         "label": "OK",         "cls": "ok"}

def build_semana_entry(row):
    """Construye el dict de datos para una semana específica."""
    cards_html = ""
    bar_labels, bar_data, bar_colors = [], [], []
    total_ok = total_prec = total_alert = total_limit = total_sin = 0

    for i, g in enumerate(GRUAS):
        gid  = g["id"]
        hrs  = row.get(f"{gid}_hrs")
        acum = row.get(gid)
        st   = get_status(hrs)

        # Contadores KPI
        if   st["key"] == "ok":       total_ok    += 1
        elif st["key"] == "precaution":total_prec  += 1
        elif st["key"] == "alert":    total_alert  += 1
        elif st["key"] == "limit":    total_limit  += 1
        else:                         total_sin    += 1

        if hrs is not None:
            pct  = min(hrs / LIMIT_HRS * 100, 100)
            disp = max(LIMIT_HRS - hrs, 0)
            hrs_disp  = f"{hrs:.1f}"
            acum_disp = f"Acum: {acum:,.1f} hrs" if isinstance(acum, float) else ""
        else:
            pct  = 0
            disp = LIMIT_HRS
            hrs_disp  = "—"
            acum_disp = "Sin lectura"

        cards_html += f"""<div class="crane-card s-{st['cls']}" style="animation-delay:{i*0.04:.2f}s">
          <div class="crane-header">
            <div class="crane-name">{gid}</div>
            <div class="status-badge {st['cls']}">● {st['label']}</div>
          </div>
          <div class="crane-plate">{g['tipo']} · {g['empresa']}</div>
          <div class="crane-km">
            <span class="crane-km-val">{hrs_disp}</span>
            <span class="crane-km-of">/ {LIMIT_HRS} hrs</span>
          </div>
          <div class="prog-bar"><div class="prog-fill {st['cls']}" style="width:{pct:.0f}%"></div></div>
          <div class="crane-footer">
            <span>{pct:.0f}% usado</span>
            <span class="disp">{acum_disp}</span>
          </div>
        </div>"""

        bar_labels.append(gid)
        bar_data.append(round(hrs, 1) if hrs is not None else 0)
        bar_colors.append(GRUA_COLORS.get(gid, "#00b8d9"))

    return {
        "sem":     row["sem"],
        "fecha":   row["fecha"].strftime("%d/%m/%Y"),
        "total":   len(GRUAS),
        "ok":      total_ok,
        "prec":    total_prec,
        "alert":   total_alert,
        "limit":   total_limit,
        "sin":     total_sin,
        "cards":   cards_html,
        "bar_labels": bar_labels,
        "bar_data":   bar_data,
        "bar_colors": bar_colors,
    }

def make_selectors(data_keys, actual_yr, actual_sem):
    """Genera los <option> para año y semana."""
    years = sorted(set(int(k.split("_")[0]) for k in data_keys), reverse=True)
    sems_por_year = {}
    for k in data_keys:
        yr, sem = k.split("_")
        sems_por_year.setdefault(yr, []).append(int(sem))

    yr_opts = ""
    for yr in years:
        sel = "selected" if yr == actual_yr else ""
        yr_opts += f'<option value="{yr}" {sel}>{yr}</option>'

    sem_opts = ""
    for sem in sorted(sems_por_year.get(str(actual_yr), [actual_sem])):
        sel = "selected" if sem == actual_sem else ""
        sem_opts += f'<option value="{sem}" {sel}>Sem {sem}</option>'

    return yr_opts, sem_opts

# ── Main ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not SHEET_URL:
        print("ERROR: Variable SHEET_URL no configurada."); sys.exit(1)

    now         = datetime.now()
    año_actual  = now.year
    sem_actual  = now.isocalendar()[1]

    raw = download_excel(SHEET_URL)

    gruas_js = {}
    for year in [2024, 2025, 2026]:
        raw.seek(0)
        rows = leer_hoja_anual(raw, year)
        if not rows:
            continue
        rows = calcular_horas_semana(rows)
        for row in rows:
            # Solo incluir semanas que tengan al menos una lectura real
            tiene_dato = any(row.get(f"{gid}_hrs") is not None for gid in GRUA_IDS)
            if not tiene_dato:
                continue
            key = f"{year}_{row['sem']}"
            gruas_js[key] = build_semana_entry(row)

    if not gruas_js:
        print("ERROR: No se encontraron datos en las hojas SEMANAS."); sys.exit(1)

    # Semana más reciente disponible
    ultima_key  = sorted(gruas_js.keys(), key=lambda k: (int(k.split("_")[0]), int(k.split("_")[1])))[-1]
    ultima_yr   = int(ultima_key.split("_")[0])
    ultima_sem  = int(ultima_key.split("_")[1])
    if año_actual in [int(k.split("_")[0]) for k in gruas_js]:
        año_mostrar = año_actual
        sem_mostrar = max(int(k.split("_")[1]) for k in gruas_js if k.startswith(str(año_actual)))
    else:
        año_mostrar = ultima_yr
        sem_mostrar = ultima_sem

    yr_opts, sem_opts = make_selectors(gruas_js.keys(), año_mostrar, sem_mostrar)

    print(f"Semanas procesadas: {len(gruas_js)}")
    print(f"Mostrando por defecto: Año {año_mostrar} Sem {sem_mostrar}")

    with open("template.html", "r", encoding="utf-8") as f:
        t = f.read()

    html = t \
        .replace("{{HORA}}",             now.strftime("%H:%M")) \
        .replace("{{FECHA}}",            now.strftime("%d/%m/%Y")) \
        .replace("{{GRUAS_YEAR_OPTIONS}}", yr_opts) \
        .replace("{{GRUAS_SEM_OPTIONS}}", sem_opts) \
        .replace("{{GRUAS_JS_DATA}}",    json.dumps(gruas_js, ensure_ascii=False)) \
        .replace("{{AÑO_ACTUAL}}",       str(año_mostrar)) \
        .replace("{{SEM_ACTUAL}}",       str(sem_mostrar))

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente.")
