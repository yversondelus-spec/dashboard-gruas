import os
import sys
import json
import requests
import pandas as pd
from datetime import datetime
from io import BytesIO

# ─── CONFIG ────────────────────────────────────────────────────────────────
LIMIT_KM = 160
ONEDRIVE_URL = os.environ.get("ONEDRIVE_URL", "")

# ─── DOWNLOAD EXCEL ────────────────────────────────────────────────────────
def download_excel(url):
    # Convert OneDrive share URL to direct download URL
    if "1drv.ms" in url or "sharepoint.com" in url or "onedrive.live.com" in url:
        if "1drv.ms" in url:
            r = requests.get(url, allow_redirects=True)
            url = r.url
        # Convert to download URL
        if "resid=" in url:
            url = url.replace("redir?", "download?").replace("embed?", "download?")
        elif "sharepoint.com" in url:
            url = url.replace("/:x:/", "/:x:/").split("?")[0] + "?download=1"
    
    print(f"Descargando Excel desde OneDrive...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return BytesIO(r.content)

# ─── PARSE DATA ────────────────────────────────────────────────────────────
def parse_gruas(excel_file):
    # Try to read the Excel - expects columns: GRUA, PATENTE, KM_MES
    # Adjust sheet_name if needed
    try:
        df = pd.read_excel(excel_file, sheet_name=0)
    except Exception as e:
        print(f"Error leyendo Excel: {e}")
        sys.exit(1)

    # Normalize column names
    df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    print(f"Columnas detectadas: {list(df.columns)}")

    # Map flexible column names
    col_map = {}
    for col in df.columns:
        if any(k in col for k in ["GRUA", "GRÚA", "ID", "NUMERO", "NÚMERO"]): col_map["grua"] = col
        if any(k in col for k in ["PATENTE", "PLACA", "PLATE"]): col_map["patente"] = col
        if any(k in col for k in ["KM", "KILOMETRO", "KILÓMETRO", "RECORRIDO"]): col_map["km"] = col

    required = ["grua", "km"]
    for r in required:
        if r not in col_map:
            print(f"ERROR: No se encontró columna para '{r}'. Columnas disponibles: {list(df.columns)}")
            sys.exit(1)

    gruas = []
    for _, row in df.iterrows():
        grua_id = str(row[col_map["grua"]]).strip()
        if not grua_id or grua_id == "nan":
            continue
        km = float(str(row[col_map["km"]]).replace(",", ".")) if str(row[col_map["km"]]) != "nan" else 0
        patente = str(row[col_map["patente"]]).strip() if "patente" in col_map else "—"
        gruas.append({"id": grua_id, "patente": patente, "km": km})

    return gruas

# ─── STATUS ────────────────────────────────────────────────────────────────
def get_status(km):
    pct = km / LIMIT_KM
    if km >= LIMIT_KM:   return {"key": "limit",     "label": "Límite",    "cls": "limit"}
    if pct >= 0.86:      return {"key": "alert",     "label": "Alerta",    "cls": "alert"}
    if pct >= 0.60:      return {"key": "precaution","label": "Precaución","cls": "precaution"}
    return               {"key": "ok",          "label": "OK",        "cls": "ok"}

# ─── BUILD HTML ────────────────────────────────────────────────────────────
def build_html(gruas):
    now = datetime.now()
    fecha = now.strftime("%d/%m/%Y")
    hora  = now.strftime("%H:%M")

    # KPIs
    total  = len(gruas)
    ok     = sum(1 for g in gruas if get_status(g["km"])["key"] == "ok")
    prec   = sum(1 for g in gruas if get_status(g["km"])["key"] == "precaution")
    alert  = sum(1 for g in gruas if get_status(g["km"])["key"] == "alert")
    limit  = sum(1 for g in gruas if get_status(g["km"])["key"] == "limit")
    con_datos = sum(1 for g in gruas if g["km"] > 0)

    # Crane cards HTML
    cards_html = ""
    for i, g in enumerate(gruas):
        st  = get_status(g["km"])
        pct = min(g["km"] / LIMIT_KM * 100, 100)
        disp = max(LIMIT_KM - g["km"], 0)
        delay = f"{i * 0.05:.2f}"
        cards_html += f"""
        <div class="crane-card s-{st['cls']}" style="animation-delay:{delay}s">
          <div class="crane-header">
            <div class="crane-name">GRÚA {g['id']}</div>
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

    # Chart data
    bar_labels = json.dumps([str(g["id"]) for g in gruas[:12]])
    bar_data   = json.dumps([g["km"] for g in gruas[:12]])
    bar_colors_map = {"ok":"#00e676","precaution":"#ffa726","alert":"#ffca28","limit":"#f44336"}
    bar_colors = json.dumps([bar_colors_map[get_status(g["km"])["key"]] for g in gruas[:12]])
    donut_data = json.dumps([ok, prec, alert, limit])

    # Read template
    with open("template.html", "r", encoding="utf-8") as f:
        template = f.read()

    html = template \
        .replace("{{FECHA}}", fecha) \
        .replace("{{HORA}}", hora) \
        .replace("{{TOTAL}}", str(total)) \
        .replace("{{CON_DATOS}}", str(con_datos)) \
        .replace("{{OK}}", str(ok)) \
        .replace("{{PRECAUTION}}", str(prec)) \
        .replace("{{ALERT}}", str(alert)) \
        .replace("{{LIMIT}}", str(limit)) \
        .replace("{{CARDS_HTML}}", cards_html) \
        .replace("{{BAR_LABELS}}", bar_labels) \
        .replace("{{BAR_DATA}}", bar_data) \
        .replace("{{BAR_COLORS}}", bar_colors) \
        .replace("{{DONUT_DATA}}", donut_data)

    return html

# ─── MAIN ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not ONEDRIVE_URL:
        print("ERROR: Variable ONEDRIVE_URL no configurada.")
        sys.exit(1)

    excel_file = download_excel(ONEDRIVE_URL)
    gruas = parse_gruas(excel_file)
    print(f"Grúas cargadas: {len(gruas)}")

    html = build_html(gruas)

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("✅ index.html generado correctamente.")
