import os, sys, json, requests, pandas as pd
from datetime import datetime
from io import BytesIO

LIMIT_KM = 160
SHEET_URL = os.environ.get("SHEET_URL", "")

def download_excel(url):
    if "docs.google.com/spreadsheets" in url:
        sheet_id = url.split("/d/")[1].split("/")[0]
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    print("Descargando desde Google Sheets...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return BytesIO(r.content)

def parse_gruas(excel_file):
    try:
        df = pd.read_excel(excel_file, sheet_name='KM_MES_ACTUAL', header=3, usecols="A:E")
    except Exception as e:
        print(f"Error leyendo KM_MES_ACTUAL: {e}"); sys.exit(1)
    gruas = []
    for _, row in df.iterrows():
        vals = list(row)
        grua_id = str(vals[1]).strip() if len(vals) > 1 else ''
        if not grua_id or grua_id == 'nan': continue
        patente = str(vals[2]).strip() if len(vals) > 2 and str(vals[2]) != 'nan' else '—'
        km = float(str(vals[4]).replace(",",".")) if len(vals) > 4 and str(vals[4]) != 'nan' else 0
        gruas.append({"id": grua_id, "patente": patente, "km": km})
    if not gruas:
        print("ERROR: Sin grúas en KM_MES_ACTUAL"); sys.exit(1)
    return gruas

def parse_combustible(excel_file):
    # Defaults si no hay hoja
    defaults = {
        "stock": 850, "stock_inicial": 2000, "consumo": 1150, "cargas": 8, "alertas": 2,
        "vehiculos": [
            {"id":"Vehículo 1","patente":"AA-1234","litros":142,"tipo":"Diesel"},
            {"id":"Vehículo 2","patente":"BB-5678","litros":98,"tipo":"Diesel"},
            {"id":"Vehículo 3","patente":"CC-9012","litros":215,"tipo":"Diesel"},
            {"id":"Vehículo 4","patente":"DD-3456","litros":87,"tipo":"Diesel"},
            {"id":"Vehículo 5","patente":"EE-7890","litros":178,"tipo":"Diesel"},
            {"id":"Vehículo 6","patente":"FF-1234","litros":134,"tipo":"Diesel"},
            {"id":"Vehículo 7","patente":"GG-5678","litros":165,"tipo":"Diesel"},
            {"id":"Vehículo 8","patente":"HH-9012","litros":131,"tipo":"Diesel"},
        ]
    }
    try:
        df = pd.read_excel(excel_file, sheet_name='COMBUSTIBLE', header=None)
        # Buscar stock inicial en la hoja
        stock_inicial = 2000
        for i, row in df.iterrows():
            for val in row:
                if str(val).replace('.','',1).isdigit() and float(str(val)) > 500:
                    try: stock_inicial = float(val)
                    except: pass
                    break

        # Intentar leer tabla de vehículos (header en fila 5)
        df2 = pd.read_excel(excel_file, sheet_name='COMBUSTIBLE', header=4, usecols="A:J")
        df2.columns = [str(c).strip().upper() for c in df2.columns]

        vehiculos = []
        consumo_total = 0
        cargas_total = 0
        alertas = 0
        for _, row in df2.iterrows():
            vals = list(row)
            if not vals or str(vals[0]) == 'nan': continue
            patente = str(vals[1]).strip() if len(vals) > 1 and str(vals[1]) != 'nan' else '—'
            litros = 0
            for v in vals[3:]:
                try:
                    litros += float(str(v).replace(',','.'))
                except: pass
            tipo = str(vals[2]).strip() if len(vals) > 2 and str(vals[2]) != 'nan' else 'Diesel'
            if litros > 0:
                consumo_total += litros
                cargas_total += 1
                if litros > 150: alertas += 1
                vehiculos.append({"id": str(vals[0]).strip(), "patente": patente, "litros": round(litros,1), "tipo": tipo})

        if not vehiculos:
            return defaults

        stock = max(stock_inicial - consumo_total, 0)
        return {
            "stock": round(stock), "stock_inicial": round(stock_inicial),
            "consumo": round(consumo_total), "cargas": cargas_total,
            "alertas": alertas, "vehiculos": vehiculos
        }
    except Exception as e:
        print(f"Usando datos por defecto para combustible: {e}")
        return defaults

def get_km_status(km):
    pct = km / LIMIT_KM
    if km >= LIMIT_KM:  return {"key":"limit",     "label":"Límite",    "cls":"limit"}
    if pct >= 0.86:     return {"key":"alert",     "label":"Alerta",    "cls":"alert"}
    if pct >= 0.60:     return {"key":"precaution","label":"Precaución","cls":"precaution"}
    return                     {"key":"ok",        "label":"OK",        "cls":"ok"}

def get_vehicle_status(litros, avg):
    if litros > avg * 1.4: return "v-low"
    if litros > avg * 1.1: return "v-warn"
    return "v-ok"

def get_vehicle_badge(litros, avg):
    if litros > avg * 1.4: return ("limit","Alto")
    if litros > avg * 1.1: return ("alert","Moderado")
    return ("ok","Normal")

def build_html(gruas, comb):
    now = datetime.now()
    fecha = now.strftime("%d/%m/%Y")
    hora  = now.strftime("%H:%M")

    total     = len(gruas)
    ok        = sum(1 for g in gruas if get_km_status(g["km"])["key"] == "ok")
    prec      = sum(1 for g in gruas if get_km_status(g["km"])["key"] == "precaution")
    alert     = sum(1 for g in gruas if get_km_status(g["km"])["key"] == "alert")
    limit     = sum(1 for g in gruas if get_km_status(g["km"])["key"] == "limit")
    con_datos = sum(1 for g in gruas if g["km"] > 0)

    # Grúas cards
    cards_html = ""
    for i, g in enumerate(gruas):
        st = get_km_status(g["km"])
        pct = min(g["km"] / LIMIT_KM * 100, 100)
        disp = max(LIMIT_KM - g["km"], 0)
        cards_html += f"""
        <div class="crane-card s-{st['cls']}" style="animation-delay:{i*0.05:.2f}s">
          <div class="crane-header"><div class="crane-name">{g['id']}</div><div class="status-badge {st['cls']}">● {st['label']}</div></div>
          <div class="crane-plate">{g['patente']}</div>
          <div class="crane-km"><span class="crane-km-val">{int(g['km'])}</span><span class="crane-km-of">/ {LIMIT_KM} km</span></div>
          <div class="prog-bar"><div class="prog-fill {st['cls']}" style="width:{pct:.0f}%"></div></div>
          <div class="crane-footer"><span>{pct:.0f}% usado</span><span class="disp">{int(disp)} km disp.</span></div>
        </div>"""

    # Combustible cards
    vehiculos = comb["vehiculos"]
    avg_litros = sum(v["litros"] for v in vehiculos) / len(vehiculos) if vehiculos else 1
    vehicle_cards = ""
    for i, v in enumerate(vehiculos):
        vcls = get_vehicle_status(v["litros"], avg_litros)
        badge_cls, badge_label = get_vehicle_badge(v["litros"], avg_litros)
        max_litros = max(v2["litros"] for v2 in vehiculos) if vehiculos else 1
        pct = min(v["litros"] / max_litros * 100, 100) if max_litros else 0
        vehicle_cards += f"""
        <div class="vehicle-card {vcls}" style="animation-delay:{i*0.05:.2f}s">
          <div class="vehicle-header"><div class="vehicle-name">{v['id']}</div><div class="status-badge {badge_cls}">● {badge_label}</div></div>
          <div class="vehicle-plate">{v['patente']}</div>
          <div class="vehicle-litros"><span class="vehicle-litros-val">{int(v['litros'])}</span><span class="vehicle-litros-of">L consumidos</span></div>
          <div class="prog-bar"><div class="prog-fill {'limit' if badge_cls=='limit' else 'alert' if badge_cls=='alert' else 'ok'}" style="width:{pct:.0f}%"></div></div>
          <div class="vehicle-footer"><span>{v['tipo']}</span><span class="comb-tag">🔴 Diesel</span></div>
        </div>"""

    stock_pct = min(round(comb["stock"] / comb["stock_inicial"] * 100), 100) if comb["stock_inicial"] else 0
    color_map = {"ok":"#00e676","precaution":"#ffa726","alert":"#ffca28","limit":"#f44336"}
    comb_colors = ['#ffa726','#ffca28','#00b8d9','#00e676','#f44336','#ab47bc','#26c6da','#66bb6a','#ef5350','#29b6f6']

    bar_labels = json.dumps([str(g["id"]) for g in gruas[:12]])
    bar_data   = json.dumps([g["km"] for g in gruas[:12]])
    bar_colors = json.dumps([color_map[get_km_status(g["km"])["key"]] for g in gruas[:12]])
    donut_data = json.dumps([ok, prec, alert, limit])
    comb_bar_labels = json.dumps([v["id"] for v in vehiculos])
    comb_bar_data   = json.dumps([v["litros"] for v in vehiculos])
    comb_bar_colors = json.dumps(comb_colors[:len(vehiculos)])

    with open("template.html", "r", encoding="utf-8") as f:
        t = f.read()

    return t \
        .replace("{{FECHA}}", fecha).replace("{{HORA}}", hora) \
        .replace("{{TOTAL}}", str(total)).replace("{{CON_DATOS}}", str(con_datos)) \
        .replace("{{OK}}", str(ok)).replace("{{PRECAUTION}}", str(prec)) \
        .replace("{{ALERT}}", str(alert)).replace("{{LIMIT}}", str(limit)) \
        .replace("{{CARDS_HTML}}", cards_html) \
        .replace("{{BAR_LABELS}}", bar_labels).replace("{{BAR_DATA}}", bar_data) \
        .replace("{{BAR_COLORS}}", bar_colors).replace("{{DONUT_DATA}}", donut_data) \
        .replace("{{COMB_STOCK}}", str(comb["stock"])) \
        .replace("{{COMB_STOCK_INICIAL}}", str(comb["stock_inicial"])) \
        .replace("{{COMB_STOCK_PCT}}", str(stock_pct)) \
        .replace("{{COMB_CONSUMO}}", str(comb["consumo"])) \
        .replace("{{COMB_CARGAS}}", str(comb["cargas"])) \
        .replace("{{COMB_ALERTAS}}", str(comb["alertas"])) \
        .replace("{{VEHICLE_CARDS_HTML}}", vehicle_cards) \
        .replace("{{COMB_BAR_LABELS}}", comb_bar_labels) \
        .replace("{{COMB_BAR_DATA}}", comb_bar_data) \
        .replace("{{COMB_BAR_COLORS}}", comb_bar_colors)

if __name__ == "__main__":
    if not SHEET_URL:
        print("ERROR: Variable SHEET_URL no configurada."); sys.exit(1)
    data = download_excel(SHEET_URL)
    gruas = parse_gruas(data)
    print(f"Grúas cargadas: {len(gruas)}")
    data.seek(0)
    comb = parse_combustible(data)
    print(f"Combustible - Stock: {comb['stock']}L, Vehículos: {len(comb['vehiculos'])}")
    html = build_html(gruas, comb)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente.")
