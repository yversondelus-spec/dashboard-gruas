import os, sys, json, requests, pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO

# ── Configuración ──────────────────────────────────────────────────────────
LIMIT_HRS = 160
SHEET_URL  = os.environ.get("SHEET_URL", "")

MESES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
         7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

GRUAS = [
    {"id":"ROYAL 9023",       "tipo":"ALTA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9024",       "tipo":"ALTA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9025",       "tipo":"ALTA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9026",       "tipo":"ALTA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9027",       "tipo":"ALTA", "empresa":"Royal Leasing"},
    {"id":"ROYAL 9028",       "tipo":"ALTA", "empresa":"Royal Leasing"},
    {"id":"LINDE 11728",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11731",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11732",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11733",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11734",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11735",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11736",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11738",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"LINDE 11739",      "tipo":"BAJA", "empresa":"Linde Leasing"},
    {"id":"TRACTOR EQUIPAJE", "tipo":"MULA", "empresa":"—"},
]
GRUA_IDS = [g["id"] for g in GRUAS]

GRUA_COLORS = {
    "ROYAL 9023":"#00b8d9","ROYAL 9024":"#00c8e8","ROYAL 9025":"#0098b9",
    "ROYAL 9026":"#0078a0","ROYAL 9027":"#005a80","ROYAL 9028":"#004a6e",
    "LINDE 11728":"#00e676","LINDE 11731":"#00cc60","LINDE 11732":"#00b050",
    "LINDE 11733":"#009040","LINDE 11734":"#007030","LINDE 11735":"#005020",
    "LINDE 11736":"#00e699","LINDE 11738":"#00cc80","LINDE 11739":"#00aa66",
    "TRACTOR EQUIPAJE":"#ffa726",
}

def download_excel(url):
    if "docs.google.com/spreadsheets" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    print("Descargando Excel desde Google Sheets...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return BytesIO(r.content)

def leer_hoja_anual(excel_bytes, year):
    nombre = f"SEMANAS {year}"
    try:
        df = pd.read_excel(excel_bytes, sheet_name=nombre, header=None,
                           skiprows=5, usecols=range(19))
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
    """Horómetro acumulado → horas por semana = lectura actual - lectura anterior."""
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

def agrupar_por_mes(rows):
    """Suma horas semanales de cada grúa agrupadas por mes real de la fecha."""
    meses = {}
    for row in rows:
        fecha = row["fecha"]
        key   = f"{fecha.year}_{fecha.month}"
        if key not in meses:
            meses[key] = {
                "year": fecha.year, "month": fecha.month,
                "semanas": [],
                "hrsporgid":  {gid: 0.0   for gid in GRUA_IDS},
                "tiene_dato": {gid: False  for gid in GRUA_IDS},
            }
        meses[key]["semanas"].append(row["fecha"].strftime("%d/%m"))
        for gid in GRUA_IDS:
            hrs = row.get(f"{gid}_hrs")
            if hrs is not None and isinstance(hrs, float):
                meses[key]["hrsporgid"][gid]  += hrs
                meses[key]["tiene_dato"][gid]  = True
    return meses

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

def build_mes_entry(mes_data):
    cards_html = ""
    bar_labels, bar_data, bar_colors = [], [], []
    total_ok = total_prec = total_alert = total_limit = total_sin = 0

    for i, g in enumerate(GRUAS):
        gid        = g["id"]
        hrs        = mes_data["hrsporgid"][gid]
        tiene_dato = mes_data["tiene_dato"][gid]
        st         = get_status(hrs, tiene_dato)

        if   st["key"] == "ok":         total_ok    += 1
        elif st["key"] == "precaution": total_prec  += 1
        elif st["key"] == "alert":      total_alert += 1
        elif st["key"] == "limit":      total_limit += 1
        else:                           total_sin   += 1

        if tiene_dato:
            pct      = min(hrs / LIMIT_HRS * 100, 100)
            disp     = max(LIMIT_HRS - hrs, 0)
            hrs_disp = f"{hrs:.1f}"
            footer_r = f"{disp:.1f} hrs disponibles"
        else:
            pct = 0; hrs_disp = "—"; footer_r = "Sin lectura"

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
            <span>{pct:.0f}% del límite mensual</span>
            <span class="disp">{footer_r}</span>
          </div>
        </div>"""

        bar_labels.append(gid)
        bar_data.append(round(hrs, 1) if tiene_dato else 0)
        bar_colors.append(GRUA_COLORS.get(gid, "#00b8d9"))

    return {
        "year":    mes_data["year"],
        "month":   mes_data["month"],
        "mes_nom": MESES[mes_data["month"]],
        "n_sem":   len(mes_data["semanas"]),
        "semanas": ", ".join(mes_data["semanas"]),
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

def make_selectors(data_keys, actual_yr, actual_mo):
    years = sorted(set(int(k.split("_")[0]) for k in data_keys), reverse=True)
    meses_por_year = {}
    for k in data_keys:
        yr, mo = k.split("_")
        meses_por_year.setdefault(yr, []).append(int(mo))

    yr_opts = ""
    for yr in years:
        sel = "selected" if yr == actual_yr else ""
        yr_opts += f'<option value="{yr}" {sel}>{yr}</option>'

    mo_opts = ""
    for mo in sorted(meses_por_year.get(str(actual_yr), [actual_mo])):
        sel = "selected" if mo == actual_mo else ""
        mo_opts += f'<option value="{mo}" {sel}>{MESES[mo]}</option>'

    return yr_opts, mo_opts

# ── Main ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not SHEET_URL:
        print("ERROR: Variable SHEET_URL no configurada."); sys.exit(1)

    now        = datetime.now()
    año_actual = now.year
    mes_actual = now.month

    raw = download_excel(SHEET_URL)

    gruas_js = {}
    for year in [2024, 2025, 2026]:
        raw.seek(0)
        rows = leer_hoja_anual(raw, year)
        if not rows:
            continue
        rows = calcular_horas_semanales(rows)
        meses = agrupar_por_mes(rows)
        for key, mes_data in meses.items():
            if any(mes_data["tiene_dato"].values()):
                gruas_js[key] = build_mes_entry(mes_data)

    if not gruas_js:
        print("ERROR: No se encontraron datos."); sys.exit(1)

    # Default: mes actual si tiene datos, si no el más reciente
    key_actual = f"{año_actual}_{mes_actual}"
    if key_actual in gruas_js:
        yr_default, mo_default = año_actual, mes_actual
    else:
        ultima = sorted(gruas_js.keys(), key=lambda k:(int(k.split("_")[0]),int(k.split("_")[1])))[-1]
        yr_default = int(ultima.split("_")[0])
        mo_default = int(ultima.split("_")[1])

    yr_opts, mo_opts = make_selectors(gruas_js.keys(), yr_default, mo_default)

    print(f"Meses procesados: {sorted(gruas_js.keys())}")
    print(f"Mostrando: {MESES[mo_default]} {yr_default}")

    with open("template.html", "r", encoding="utf-8") as f:
        t = f.read()

    html = t \
        .replace("{{HORA}}",               now.strftime("%H:%M")) \
        .replace("{{FECHA}}",              now.strftime("%d/%m/%Y")) \
        .replace("{{GRUAS_YEAR_OPTIONS}}", yr_opts) \
        .replace("{{GRUAS_MES_OPTIONS}}",  mo_opts) \
        .replace("{{GRUAS_JS_DATA}}",      json.dumps(gruas_js, ensure_ascii=False)) \
        .replace("{{AÑO_ACTUAL}}",         str(yr_default)) \
        .replace("{{MES_ACTUAL}}",         str(mo_default))

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente.")
