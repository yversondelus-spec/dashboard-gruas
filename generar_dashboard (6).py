import os, sys, json, requests, pandas as pd
from datetime import datetime
from io import BytesIO

LIMIT_KM = 160
SHEET_URL = os.environ.get("SHEET_URL", "")

def download_excel(url):
    # Convierte URL de Google Sheets a descarga directa xlsx
    if "docs.google.com/spreadsheets" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    print(f"Descargando desde Google Sheets...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return BytesIO(r.content)

def parse_gruas(excel_file):
    try:
        df = pd.read_excel(excel_file, sheet_name='KM_MES_ACTUAL', header=3, usecols="A:E")
    except Exception as e:
        print(f"Error leyendo hoja KM_MES_ACTUAL: {e}")
        sys.exit(1)

    df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    print(f"Columnas: {list(df.columns)}")

    gruas = []
    for _, row in df.iterrows():
        vals = list(row)
        grua_id = str(vals[1]).strip() if len(vals) > 1 else ''
        if not grua_id or grua_id == 'nan': continue
        patente = str(vals[2]).strip() if len(vals) > 2 and str(vals[2]) != 'nan' else '—'
        km = float(str(vals[4]).replace(",",".")) if len(vals) > 4 and str(vals[4]) != 'nan' else 0
        gruas.append({"id": grua_id, "patente": patente, "km": km})

    if not gruas:
        print("ERROR: No se encontraron grúas en KM_MES_ACTUAL.")
        sys.exit(1)
    return gruas

def get_status(km):
    pct = km / LIMIT_KM
    if km >= LIMIT_KM:  return {"key":"limit",     "label":"Límite",    "cls":"limit"}
    if pct >= 0.86:     return {"key":"alert",     "label":"Alerta",    "cls":"alert"}
    if pct >= 0.60:     return {"key":"precaution","label":"Precaución","cls":"precaution"}
    return                     {"key":"ok",        "label":"OK",        "cls":"ok"}

def build_html(gruas):
    now = datetime.now()
    fecha = now.strftime("%d/%m/%Y")
    hora  = now.strftime("%H:%M")
    total     = len(gruas)
    ok        = sum(1 for g in gruas if get_status(g["km"])["key"] == "ok")
    prec      = sum(1 for g in gruas if get_status(g["km"])["key"] == "precaution")
    alert     = sum(1 for g in gruas if get_status(g["km"])["key"] == "alert")
    limit     = sum(1 for g in gruas if get_status(g["km"])["key"] == "limit")
    con_datos = sum(1 for g in gruas if g["km"] > 0)

    cards_html = ""
    for i, g in enumerate(gruas):
        st   = get_status(g["km"])
        pct  = min(g["km"] / LIMIT_KM * 100, 100)
        disp = max(LIMIT_KM - g["km"], 0)
        delay = f"{i * 0.05:.2f}"
        cards_html += f"""
        <div class="crane-card s-{st['cls']}" style="animation-delay:{delay}s">
          <div class="crane-header">
            <div class="crane-name">{g['id']}</div>
            <div class="status-badge {st['cls']}">● {st['label']}</div>
          </div>
          <div class="crane-plate">{g['patente']}</div>
          <div class="crane-km">
            <span class="crane-km-val">{int(g['km'])}</span>
            <span class="crane-km-of">/ {LIMIT_KM} km</span>
          </div>
          <div class="prog-bar"><div class="prog-fill {st['cls']}" style="width:{pct:.0f}%"></div></div>
          <div class="crane-footer">
            <span>{pct:.0f}% usado</span>
            <span class="disp">{int(disp)} km disp.</span>
          </div>
        </div>"""

    color_map  = {"ok":"#00e676","precaution":"#ffa726","alert":"#ffca28","limit":"#f44336"}
    bar_labels = json.dumps([str(g["id"]) for g in gruas[:12]])
    bar_data   = json.dumps([g["km"] for g in gruas[:12]])
    bar_colors = json.dumps([color_map[get_status(g["km"])["key"]] for g in gruas[:12]])
    donut_data = json.dumps([ok, prec, alert, limit])

    with open("template.html", "r", encoding="utf-8") as f:
        template = f.read()

    return template \
        .replace("{{FECHA}}", fecha).replace("{{HORA}}", hora) \
        .replace("{{TOTAL}}", str(total)).replace("{{CON_DATOS}}", str(con_datos)) \
        .replace("{{OK}}", str(ok)).replace("{{PRECAUTION}}", str(prec)) \
        .replace("{{ALERT}}", str(alert)).replace("{{LIMIT}}", str(limit)) \
        .replace("{{CARDS_HTML}}", cards_html) \
        .replace("{{BAR_LABELS}}", bar_labels).replace("{{BAR_DATA}}", bar_data) \
        .replace("{{BAR_COLORS}}", bar_colors).replace("{{DONUT_DATA}}", donut_data)

if __name__ == "__main__":
    if not SHEET_URL:
        print("ERROR: Variable SHEET_URL no configurada.")
        sys.exit(1)
    excel_file = download_excel(SHEET_URL)
    gruas = parse_gruas(excel_file)
    print(f"Grúas cargadas: {len(gruas)}")
    html = build_html(gruas)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente.")
