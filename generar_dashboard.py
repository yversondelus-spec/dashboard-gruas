import pandas as pd
import json
import datetime
import sys
import os

# ─────────────────────────────────────────────────────────────────
# CONFIGURACION
# El Excel debe estar subido al repositorio con este nombre exacto
# ─────────────────────────────────────────────────────────────────
EXCEL_FILE = "CONTROL_HITOS.xlsx"

def fmt_time(t):
    if pd.isna(t) or t is None:
        return None
    if isinstance(t, datetime.time):
        return f"{t.hour:02d}:{t.minute:02d}"
    if isinstance(t, datetime.timedelta):
        total = int(t.total_seconds())
        return f"{total//3600:02d}:{(total%3600)//60:02d}"
    return str(t)

def to_min(t):
    if pd.isna(t) or t is None:
        return 0
    if isinstance(t, datetime.time):
        return t.hour * 60 + t.minute
    if isinstance(t, datetime.timedelta):
        return int(t.total_seconds() // 60)
    try:
        return int(float(t))
    except:
        return 0

def procesar_hitos(df):
    """Procesa una hoja de hitos (DHL 2026, DHL 2025, MAS AIR)."""
    rows = []
    for _, r in df.iterrows():
        fecha = r.iloc[0]
        if not isinstance(fecha, (datetime.datetime, datetime.date)):
            continue
        if isinstance(fecha, datetime.datetime) and fecha.year < 2020:
            continue
        vuelo = r.iloc[1]
        if pd.isna(vuelo) or str(vuelo).strip() == '':
            continue
        mot = str(r.iloc[11]).strip() if not pd.isna(r.iloc[11]) else ""
        rows.append({
            "f":   fecha.strftime("%Y-%m-%d"),
            "mn":  fecha.month,
            "an":  fecha.year,
            "v":   str(vuelo).strip(),
            "m":   str(r.iloc[2]).strip() if not pd.isna(r.iloc[2]) else "",
            "p":   str(r.iloc[3]).strip() if not pd.isna(r.iloc[3]) else "",
            "eta": fmt_time(r.iloc[4]),
            "ata": fmt_time(r.iloc[5]),
            "g":   to_min(r.iloc[8]),
            "d":   to_min(r.iloc[9]),
            "c":   str(r.iloc[10]).strip() if not pd.isna(r.iloc[10]) else "",
            "mo":  mot,
            "sup": str(r.iloc[25]).strip() if len(r) > 25 and not pd.isna(r.iloc[25]) else "",
        })
    return rows

def procesar_dly(df):
    """Procesa la hoja DLY DETAILS."""
    rows = []
    for _, r in df.iterrows():
        fecha = r.iloc[0]
        if not isinstance(fecha, (datetime.datetime, datetime.date)):
            continue
        flight = r.iloc[1]
        if pd.isna(flight):
            continue
        dly = to_min(r.iloc[7])
        if dly == 0:
            continue
        mot = str(r.iloc[9]).strip() if not pd.isna(r.iloc[9]) else ""
        rows.append({
            "f":   fecha.strftime("%Y-%m-%d"),
            "mn":  fecha.month,
            "an":  fecha.year,
            "fl":  str(flight).strip(),
            "eta": fmt_time(r.iloc[2]),
            "ata": fmt_time(r.iloc[3]),
            "etd": fmt_time(r.iloc[4]),
            "atd": fmt_time(r.iloc[5]),
            "grt": fmt_time(r.iloc[6]) if not isinstance(r.iloc[6], str) else str(r.iloc[6] or ''),
            "d":   dly,
            "c":   str(r.iloc[8]).strip() if not pd.isna(r.iloc[8]) else "",
            "mo":  mot,
        })
    return rows

def generar_html(dhl_data, mas_data, dly_data, fecha_update):
    """Lee el template HTML e inyecta los datos frescos."""
    template = open("template.html", "r", encoding="utf-8").read()
    template = template.replace("__DHL_DATA__",  json.dumps(dhl_data,  ensure_ascii=False))
    template = template.replace("__MAS_DATA__",  json.dumps(mas_data,  ensure_ascii=False))
    template = template.replace("__DLY_DATA__",  json.dumps(dly_data,  ensure_ascii=False))
    template = template.replace("__FECHA_UPD__", fecha_update)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(template)
    print(f"✓ index.html generado con datos frescos")

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: No se encontró el archivo {EXCEL_FILE}")
        print("Sube tu Excel al repositorio con ese nombre exacto.")
        sys.exit(1)

    # Leer hojas
    print(f"Leyendo {EXCEL_FILE}...")
    xls = pd.ExcelFile(EXCEL_FILE)
    print(f"Hojas encontradas: {xls.sheet_names}")

    # DHL 2026
    dhl_data = []
    for sh in ["DHL 2026", "DHL 2025 ", "DHL 2025"]:
        if sh in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sh, header=None, skiprows=3)
            dhl_data = procesar_hitos(df)
            print(f"✓ {sh}: {len(dhl_data)} filas")
            break

    # MAS AIR
    mas_data = []
    for sh in ["MAS AIR ", "MAS AIR"]:
        if sh in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sh, header=None, skiprows=3)
            mas_data = procesar_hitos(df)
            print(f"✓ MAS AIR: {len(mas_data)} filas")
            break

    # DLY DETAILS
    dly_data = []
    if "DLY DETAILS" in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name="DLY DETAILS", header=None, skiprows=1)
        dly_data = procesar_dly(df)
        print(f"✓ DLY DETAILS: {len(dly_data)} filas con demora")

    fecha_update = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    generar_html(dhl_data, mas_data, dly_data, fecha_update)
    print(f"\n✅ Dashboard actualizado — {fecha_update}")

if __name__ == "__main__":
    main()
