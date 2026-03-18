"""
BRIAV32 - Script de actualización del dashboard (modo OFFLINE)
==============================================================
Uso:
    python actualizar.py NombreDetuArchivo.xlsx

Requisitos (instalar una sola vez):
    pip install pandas openpyxl requests

El script genera un index.html que funciona SIN internet.
"""

import sys, os, json, re, base64
import pandas as pd

# ── Verificar argumentos ──────────────────────────────────────────────────────
if len(sys.argv) < 2:
    print("ERROR: Debes indicar el nombre del archivo Excel.")
    print("  Ejemplo: python actualizar.py MiArchivo.xlsx")
    sys.exit(1)

excel_path = sys.argv[1]
if not os.path.exists(excel_path):
    print(f"ERROR: No encontre el archivo: {excel_path}")
    sys.exit(1)

HERE = os.path.dirname(os.path.abspath(__file__))

# ── Verificar archivos de plantilla ──────────────────────────────────────────
tmpl_before_path = os.path.join(HERE, 'tmpl_before.dat')
tmpl_after_path  = os.path.join(HERE, 'tmpl_after.dat')

if not os.path.exists(tmpl_before_path) or not os.path.exists(tmpl_after_path):
    print("ERROR: Faltan los archivos tmpl_before.dat y tmpl_after.dat")
    print("  Asegurate de que esten en la misma carpeta que este script.")
    sys.exit(1)

print(f"Leyendo: {excel_path}")

# ── Leer hoja Base_Unificada ──────────────────────────────────────────────────
try:
    df = pd.read_excel(excel_path, sheet_name='Base_Unificada')
except Exception as e:
    print(f"ERROR leyendo el Excel: {e}")
    sys.exit(1)

print(f"  {len(df)} registros encontrados.")

# ── Limpiar datos ─────────────────────────────────────────────────────────────
def clean(v):
    if pd.isna(v): return ''
    if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
    return str(v).strip()

def num(v):
    try: return float(v) if pd.notna(v) else 0
    except: return 0

records = []
for _, row in df.iterrows():
    records.append({
        'tipo':                clean(row.get('tipo', '')),
        'vigencia':            str(int(row['vigencia'])) if pd.notna(row.get('vigencia')) else '',
        'unidad':              clean(row.get('unidad', '')),
        'contrato':            clean(row.get('contrato', '')),
        'supervisor':          clean(row.get('supervisor', '')),
        'nuevo_supervisor':    clean(row.get('nuevo_supervisor', '')),
        'objeto':              clean(row.get('objeto', '')),
        'plazo_ejecucion':     clean(row.get('plazo_ejecucion', '')),
        'valor_contrato':      num(row.get('valor_contrato')),
        'valor_cxp':           num(row.get('valor_cxp')),
        'valor_reserva':       num(row.get('valor_reserva')),
        'valor_ejecutado':     num(row.get('valor_ejecutado')),
        'valor_pendiente':     num(row.get('valor_pendiente')),
        'porcentaje_ejecucion':num(row.get('porcentaje_ejecucion')),
        'fecha_prorroga':      clean(row.get('fecha_prorroga', '')),
        'empresa':             clean(row.get('empresa', '')),
        'ultima_actuacion':    clean(row.get('ultima_actuacion', '')),
        'cantidad':            int(row['cantidad']) if pd.notna(row.get('cantidad')) else 0,
        'fecha_corte':         clean(row.get('fecha_corte', '')),
        'observaciones':       clean(row.get('observaciones', '')),
    })

DATA_JSON = json.dumps(records, ensure_ascii=False, separators=(',', ':'))
print(f"  Datos listos: {len(records)} registros.")

# ── Leer plantillas ───────────────────────────────────────────────────────────
with open(tmpl_before_path, 'r', encoding='utf-8') as f:
    tmpl_before = f.read()
with open(tmpl_after_path, 'r', encoding='utf-8') as f:
    tmpl_after = f.read()

# ── Descargar recursos offline (solo primera vez) ─────────────────────────────
CHARTJS_CACHE = os.path.join(HERE, '_chartjs_cache.js')
FONTS_CACHE   = os.path.join(HERE, '_fonts_cache.css')

def descargar_recursos():
    try:
        import requests
    except ImportError:
        print("  AVISO: instala requests: pip install requests")
        return None, None

    chartjs = None
    fonts   = None

    if os.path.exists(CHARTJS_CACHE):
        print("  Chart.js: usando cache local.")
        with open(CHARTJS_CACHE, 'r', encoding='utf-8') as f: chartjs = f.read()
    else:
        print("  Descargando Chart.js (solo esta vez)...")
        try:
            r = requests.get('https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js', timeout=30)
            chartjs = r.text
            with open(CHARTJS_CACHE, 'w', encoding='utf-8') as f: f.write(chartjs)
            print(f"  Chart.js descargado: {len(chartjs):,} chars")
        except Exception as e:
            print(f"  No se pudo descargar Chart.js: {e}")

    if os.path.exists(FONTS_CACHE):
        print("  Fuentes: usando cache local.")
        with open(FONTS_CACHE, 'r', encoding='utf-8') as f: fonts = f.read()
    else:
        print("  Descargando fuentes Roboto (solo esta vez)...")
        try:
            r2 = requests.get(
                'https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&family=Roboto+Mono:wght@400;500&display=swap',
                timeout=30, headers={'User-Agent': 'Mozilla/5.0'}
            )
            font_css = r2.text
            urls = re.findall(r'url\((https://[^)]+)\)', font_css)
            for url in urls:
                try:
                    rf = requests.get(url, timeout=30)
                    b64 = base64.b64encode(rf.content).decode()
                    fmt = 'woff2' if 'woff2' in url else 'woff'
                    font_css = font_css.replace(url, f'data:font/{fmt};base64,{b64}')
                except: pass
            fonts = font_css
            with open(FONTS_CACHE, 'w', encoding='utf-8') as f: f.write(fonts)
            print("  Fuentes descargadas y embebidas.")
        except Exception as e:
            print(f"  No se pudieron descargar fuentes: {e}")

    return chartjs, fonts

print("Preparando modo offline...")
chartjs, fonts = descargar_recursos()

# ── Reemplazar recursos externos ─────────────────────────────────────────────
if fonts:
    tmpl_before = tmpl_before.replace(
        '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&family=Roboto+Mono:wght@400;500&display=swap" rel="stylesheet"/>',
        f'<style>{fonts}</style>'
    )

if chartjs:
    tmpl_before = tmpl_before.replace(
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>',
        f'<script>{chartjs}</script>'
    )

# ── Generar index.html ────────────────────────────────────────────────────────
print("Generando index.html...")
html = tmpl_before + DATA_JSON + tmpl_after

output_path = os.path.join(HERE, 'index.html')
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)

print("")
print("  ================================================================")
print(f"  Listo! index.html generado con {len(records)} registros.")
if chartjs and fonts:
    print("  Modo OFFLINE activado: funciona sin internet.")
else:
    print("  AVISO: modo online. Necesitas internet para ver las graficas.")
print("  Sube el index.html a GitHub para actualizar el dashboard.")
print("  ================================================================")
