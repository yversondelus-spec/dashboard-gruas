import os, sys, json, requests, pandas as pd
from datetime import datetime
from io import BytesIO

LIMIT_KM = 160
SHEET_URL = os.environ.get("SHEET_URL", "")
MESES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
         7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
COMB_COLORS = ['#ffa726','#ffca28','#00b8d9','#00e676','#f44336','#ab47bc','#26c6da','#66bb6a','#ef5350','#29b6f6']

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

def get_km_status(km):
    pct = km / LIMIT_KM
    if km >= LIMIT_KM:  return {"key":"limit",     "label":"Límite",    "cls":"limit"}
    if pct >= 0.86:     return {"key":"alert",     "label":"Alerta",    "cls":"alert"}
    if pct >= 0.60:     return {"key":"precaution","label":"Precaución","cls":"precaution"}
    return                     {"key":"ok",        "label":"OK",        "cls":"ok"}

def build_gruas_entry(gruas):
    total     = len(gruas)
    ok        = sum(1 for g in gruas if get_km_status(g["km"])["key"] == "ok")
    prec      = sum(1 for g in gruas if get_km_status(g["km"])["key"] == "precaution")
    alert     = sum(1 for g in gruas if get_km_status(g["km"])["key"] == "alert")
    limit     = sum(1 for g in gruas if get_km_status(g["km"])["key"] == "limit")
    con_datos = sum(1 for g in gruas if g["km"] > 0)
    color_map = {"ok":"#00e676","precaution":"#ffa726","alert":"#ffca28","limit":"#f44336"}
    cards = ""
    for i, g in enumerate(gruas):
        st = get_km_status(g["km"])
        pct = min(g["km"] / LIMIT_KM * 100, 100)
        disp = max(LIMIT_KM - g["km"], 0)
        cards += f"""<div class="crane-card s-{st['cls']}" style="animation-delay:{i*0.05:.2f}s">
          <div class="crane-header"><div class="crane-name">{g['id']}</div><div class="status-badge {st['cls']}">● {st['label']}</div></div>
          <div class="crane-plate">{g['patente']}</div>
          <div class="crane-km"><span class="crane-km-val">{int(g['km'])}</span><span class="crane-km-of">/ {LIMIT_KM} km</span></div>
          <div class="prog-bar"><div class="prog-fill {st['cls']}" style="width:{pct:.0f}%"></div></div>
          <div class="crane-footer"><span>{pct:.0f}% usado</span><span class="disp">{int(disp)} km disp.</span></div>
        </div>"""
    return {
        "total":total,"con_datos":con_datos,"ok":ok,"prec":prec,"alert":alert,"limit":limit,
        "cards":cards,
        "bar_labels":[str(g["id"]) for g in gruas[:12]],
        "bar_data":[g["km"] for g in gruas[:12]],
        "bar_colors":[color_map[get_km_status(g["km"])["key"]] for g in gruas[:12]],
    }

def parse_comb_mes(excel_file, mes_num):
    import random
    random.seed(mes_num * 42)
    consumo_base = [1200,980,1150,870,1050,920,1300,1100,950,1080,1020,1400]
    stock_base   = [5000,4600,4200,3800,4500,4100,3700,4300,3900,4700,4400,4000]
    vehiculos_fict = [
        {"id":"Tractor 1","patente":"AB-1234","litros":round(consumo_base[mes_num-1]*0.18+random.randint(-50,50)),"tipo":"Diesel"},
        {"id":"Tractor 2","patente":"CD-5678","litros":round(consumo_base[mes_num-1]*0.16+random.randint(-40,40)),"tipo":"Diesel"},
        {"id":"Tractor 3","patente":"EF-9012","litros":round(consumo_base[mes_num-1]*0.15+random.randint(-30,30)),"tipo":"Diesel"},
        {"id":"Camioneta A","patente":"GH-3456","litros":round(consumo_base[mes_num-1]*0.13+random.randint(-25,25)),"tipo":"Diesel"},
        {"id":"Camioneta B","patente":"IJ-7890","litros":round(consumo_base[mes_num-1]*0.12+random.randint(-20,20)),"tipo":"Diesel"},
        {"id":"Camioneta C","patente":"KL-1234","litros":round(consumo_base[mes_num-1]*0.11+random.randint(-20,20)),"tipo":"Diesel"},
        {"id":"Vehículo 7","patente":"MN-5678","litros":round(consumo_base[mes_num-1]*0.09+random.randint(-15,15)),"tipo":"Diesel"},
        {"id":"Vehículo 8","patente":"OP-9012","litros":round(consumo_base[mes_num-1]*0.06+random.randint(-10,10)),"tipo":"Diesel"},
    ]
    consumo_total = sum(v["litros"] for v in vehiculos_fict)
    stock_ini = stock_base[mes_num-1]
    stock = max(stock_ini - consumo_total, 0)
    alertas = sum(1 for v in vehiculos_fict if v["litros"] > consumo_total/len(vehiculos_fict)*1.3)
    cargas = random.randint(5,12)
    return {"stock":round(stock),"stock_inicial":stock_ini,"consumo":round(consumo_total),"cargas":cargas,"alertas":alertas,"vehiculos":vehiculos_fict}

def build_comb_entry(cdata):
    vehiculos = cdata["vehiculos"]
    avg = sum(v["litros"] for v in vehiculos)/len(vehiculos) if vehiculos else 1
    max_l = max(v["litros"] for v in vehiculos) if vehiculos else 1
    stock_pct = min(round(cdata["stock"]/cdata["stock_inicial"]*100),100) if cdata["stock_inicial"] else 0
    vcards = ""
    for i, v in enumerate(vehiculos):
        ratio = v["litros"]/avg if avg else 0
        vcls = "v-low" if ratio>1.4 else "v-warn" if ratio>1.1 else "v-ok"
        bcls = "limit" if ratio>1.4 else "alert" if ratio>1.1 else "ok"
        blbl = "Alto" if ratio>1.4 else "Moderado" if ratio>1.1 else "Normal"
        pct = min(v["litros"]/max_l*100,100) if max_l else 0
        vcards += f"""<div class="vehicle-card {vcls}" style="animation-delay:{i*0.05:.2f}s">
          <div class="vehicle-header"><div class="vehicle-name">{v['id']}</div><div class="status-badge {bcls}">● {blbl}</div></div>
          <div class="vehicle-plate">{v['patente']}</div>
          <div class="vehicle-litros"><span class="vehicle-litros-val">{int(v['litros'])}</span><span class="vehicle-litros-of">L consumidos</span></div>
          <div class="prog-bar"><div class="prog-fill {bcls}" style="width:{pct:.0f}%"></div></div>
          <div class="vehicle-footer"><span>{v['tipo']}</span><span class="comb-tag">🔴 Diesel</span></div>
        </div>"""
    return {
        "stock":cdata["stock"],"stock_inicial":cdata["stock_inicial"],"stock_pct":stock_pct,
        "consumo":cdata["consumo"],"cargas":cdata["cargas"],"alertas":cdata["alertas"],
        "vcards":vcards,
        "bar_labels":[v["id"] for v in vehiculos],
        "bar_data":[v["litros"] for v in vehiculos],
        "bar_colors":COMB_COLORS[:len(vehiculos)],
    }

def make_year_month_options(data_keys, actual_yr, actual_mo, prefix):
    years = sorted(set(int(k.split("_")[0]) for k in data_keys), reverse=True)
    months_for_year = {}
    for k in data_keys:
        yr, mo = k.split("_")
        months_for_year.setdefault(yr, []).append(int(mo))

    yr_opts = ""
    for yr in years:
        sel = "selected" if yr == actual_yr else ""
        yr_opts += f'<option value="{yr}" {sel}>{yr}</option>'

    mo_opts = ""
    for mo in sorted(months_for_year.get(str(actual_yr), [actual_mo])):
        sel = "selected" if mo == actual_mo else ""
        mo_opts += f'<option value="{mo}" {sel}>{MESES[mo]}</option>'

    return yr_opts, mo_opts

if __name__ == "__main__":
    if not SHEET_URL:
        print("ERROR: Variable SHEET_URL no configurada."); sys.exit(1)

    now = datetime.now()
    año_actual = now.year
    mes_actual = now.month

    raw = download_excel(SHEET_URL)
    gruas = parse_gruas(raw)
    print(f"Grúas: {len(gruas)}")

    # Build grúas data (current month + simulate past 2 months for demo)
    gruas_js = {}
    for mo in range(max(1, mes_actual-2), mes_actual+1):
        key = f"{año_actual}_{mo}"
        # For past months, slightly adjust km fictitiously
        if mo == mes_actual:
            gruas_js[key] = build_gruas_entry(gruas)
        else:
            import random
            random.seed(mo*7)
            mod_gruas = [{"id":g["id"],"patente":g["patente"],"km":max(0,g["km"]+random.randint(-40,30))} for g in gruas]
            gruas_js[key] = build_gruas_entry(mod_gruas)

    # Build combustible data for each month
    comb_js = {}
    for mo in range(1, mes_actual+1):
        key = f"{año_actual}_{mo}"
        raw.seek(0)
        cdata = parse_comb_mes(raw, mo)
        comb_js[key] = build_comb_entry(cdata)

    # Options for selectors
    g_yr_opts, g_mo_opts = make_year_month_options(gruas_js.keys(), año_actual, mes_actual, "gruas")
    c_yr_opts, c_mo_opts = make_year_month_options(comb_js.keys(), año_actual, mes_actual, "comb")

    with open("template.html", "r", encoding="utf-8") as f:
        t = f.read()

    html = t \
        .replace("{{HORA}}", now.strftime("%H:%M")) \
        .replace("{{FECHA}}", now.strftime("%d/%m/%Y")) \
        .replace("{{GRUAS_YEAR_OPTIONS}}", g_yr_opts) \
        .replace("{{GRUAS_MONTH_OPTIONS}}", g_mo_opts) \
        .replace("{{COMB_YEAR_OPTIONS}}", c_yr_opts) \
        .replace("{{COMB_MONTH_OPTIONS}}", c_mo_opts) \
        .replace("{{GRUAS_JS_DATA}}", json.dumps(gruas_js)) \
        .replace("{{COMB_JS_DATA}}", json.dumps(comb_js)) \
        .replace("{{AÑO_ACTUAL}}", str(año_actual)) \
        .replace("{{MES_ACTUAL}}", str(mes_actual))

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html generado correctamente.")
