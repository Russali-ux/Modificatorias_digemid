"""
run_scraper.py — Orquestador para GitHub Actions.
Ejecuta el scraper de modificatorias, filtra las categorías relevantes para
farmacovigilancia y escribe los outputs al GITHUB_OUTPUT para el job send_email.

Los textos estandarizados (Acción Requerida / Indicador de Tiempos / Base Legal)
y la cabecera ConkoSafe del Excel viven en scripts/scraper.py.
"""
import os, sys
from datetime import datetime

exec(open('scripts/scraper.py').read())

max_pag   = int(os.environ.get('MAX_PAGINAS', '2'))
MAX_FILAS = int(os.environ.get('MAX_FILAS', '10'))   # antes .head(10) fijo
fecha     = datetime.now().strftime('%Y%m%d_%H%M')
nombre    = f"modificatorias_digemid_{fecha}.xlsx"

# Carpeta relativa dentro del checkout — portable en Windows y Linux.
out_dir = os.path.join(os.getcwd(), 'output')
os.makedirs(out_dir, exist_ok=True)
ruta = os.path.join(out_dir, nombre)

# ── Verificación real del motor Claude ─────────────────────────────────────
# Las keys de Anthropic empiezan con 'sk-ant-', no con 'sk-'. Además hay que
# comprobar que la librería esté instalada: si falta, _analizar_claude()
# lanza ImportError, lo captura su propio except y cae al heurístico en
# silencio — que es exactamente lo que pasó en la corrida del 19/07.
api_key = os.environ.get('ANTHROPIC_API_KEY', '')
api_ok  = api_key.startswith('sk-ant-')
try:
    import anthropic  # noqa: F401
    lib_ok = True
except ImportError:
    lib_ok = False

if not api_key:
    print("AVISO: ANTHROPIC_API_KEY no definida -> motor heuristico")
elif not api_ok:
    print("AVISO: ANTHROPIC_API_KEY con formato inesperado (no empieza con 'sk-ant-')")
if not lib_ok:
    print("AVISO: libreria 'anthropic' no instalada -> motor heuristico")

motor_esperado = 'Claude API' if (api_ok and lib_ok) else 'Heuristico'
print(f"Motor de analisis esperado: {motor_esperado} "
      f"(modelo: {os.environ.get('CLAUDE_MODEL', 'claude-sonnet-4-6')})")

# ── Scrapear ───────────────────────────────────────────────────────────────
df_completo = scrapear_modificaciones(max_paginas=max_pag, analizar_pdfs=True, solo_hoy=False)

FALLOS_PDF = {"Sin PDF", "Error descarga PDF", "Error PDF", "PDF escaneado"}
total_scrapeadas = len(df_completo)
sin_pdf = int(df_completo['motor_analisis'].isin(FALLOS_PDF).sum()) if 'motor_analisis' in df_completo.columns else 0
if sin_pdf > 0:
    print(f"ADVERTENCIA: {sin_pdf}/{total_scrapeadas} modificatorias sin texto de PDF disponible")

# ── Auditoría: log completo de clasificación ANTES de filtrar ──────────────
if 'tipo_modificacion' in df_completo.columns:
    print("\n--- Clasificacion completa (antes de filtrar) ---")
    cols_debug = [c for c in ['n_modificacion', 'tipo_modificacion', 'urgencia', 'motor_analisis']
                  if c in df_completo.columns]
    print(df_completo[cols_debug].to_string())
    print()

# ── Filtrar categorías relevantes para farmacovigilancia ───────────────────
# TIPOS_SEGURIDAD viene de scraper.py: es la misma lista que decide los textos
# estandarizados, para que filtro y textos no puedan desincronizarse.
df = df_completo.copy()
if 'tipo_modificacion' in df.columns:
    excluidas = df[~df['tipo_modificacion'].isin(TIPOS_SEGURIDAD)]
    if len(excluidas):
        print(f"Excluidas ({len(excluidas)}) por tipo no relevante: "
              f"{', '.join(excluidas['n_modificacion'].astype(str).tolist())}")
    df = df[df['tipo_modificacion'].isin(TIPOS_SEGURIDAD)]

    # Ordenar por fecha ANTES de recortar: el .head(10) anterior recortaba en
    # el orden de scraping, por lo que un dia con muchas publicaciones podia
    # descartar las mas recientes sin dejar rastro.
    if 'fecha_publicacion' in df.columns:
        df = df.sort_values('fecha_publicacion', ascending=False)
    if len(df) > MAX_FILAS:
        print(f"ADVERTENCIA: {len(df)} relevantes -> se recortan a las {MAX_FILAS} mas recientes")
    df = df.head(MAX_FILAS).reset_index(drop=True)

print(f"Filtradas: {len(df)} modificatorias relevantes "
      f"(categorias: {', '.join(sorted(TIPOS_SEGURIDAD))})")

exportar_excel(df, ruta)

# ── Outputs para el job de correo ──────────────────────────────────────────
total       = len(df)
inmediatas  = int((df['urgencia'] == 'INMEDIATA').sum())  if 'urgencia' in df.columns else 0
preventivas = int((df['urgencia'] == 'PREVENTIVA').sum()) if 'urgencia' in df.columns else 0
fecha_fmt   = datetime.now().strftime('%d/%m/%Y %H:%M')

# Motor REAL segun lo que quedo en el DataFrame, no lo que se esperaba.
if 'motor_analisis' in df.columns and len(df):
    motor = ' / '.join(sorted(set(df['motor_analisis'].dropna().astype(str))))
else:
    motor = motor_esperado
if motor_esperado == 'Claude API' and 'Claude API' not in motor:
    print("ADVERTENCIA: se esperaba Claude API pero el analisis cayo al motor heuristico "
          "(revisar el log por lineas '[Claude API] ...')")

github_output = os.environ.get('GITHUB_OUTPUT')
if github_output:
    with open(github_output, 'a', encoding='utf-8') as fh:
        fh.write(f"excel_name={nombre}\n")
        fh.write(f"total_modif={total}\n")
        fh.write(f"modif_inmediatas={inmediatas}\n")
        fh.write(f"modif_preventivas={preventivas}\n")
        fh.write(f"fecha_reporte={fecha_fmt}\n")
        fh.write(f"motor_analisis={motor}\n")
        fh.write(f"sin_pdf_count={sin_pdf}\n")
        fh.write(f"total_scrapeadas={total_scrapeadas}\n")
else:
    print("GITHUB_OUTPUT no definido (ejecucion local): se omiten los outputs")

print(f"Excel: {ruta} | Total: {total} | Inmediatas: {inmediatas} | "
      f"Preventivas: {preventivas} | Motor: {motor} | Sin PDF: {sin_pdf}/{total_scrapeadas}")
