"""
run_scraper.py — Orquestador para GitHub Actions.
Ejecuta el scraper de modificatorias, filtra ACTUALIZACION SEGURIDAD (top 10)
y escribe los outputs al GITHUB_OUTPUT para el job send_email.
"""
import os, sys
from datetime import datetime

exec(open('scripts/scraper.py').read())

max_pag = int(os.environ.get('MAX_PAGINAS', '2'))
fecha   = datetime.now().strftime('%Y%m%d_%H%M')
nombre  = f"modificatorias_digemid_{fecha}.xlsx"

# Carpeta relativa dentro del checkout — portable en Windows y Linux.
# (antes era /tmp/... hardcodeado, que en el runner Windows self-hosted
# no existe por defecto y hace que el guardado del Excel falle)
out_dir = os.path.join(os.getcwd(), 'output')
os.makedirs(out_dir, exist_ok=True)
ruta = os.path.join(out_dir, nombre)

api_key = os.environ.get('ANTHROPIC_API_KEY', '')
motor   = 'Claude API' if (api_key and api_key.startswith('sk-')) else 'Heuristico'
print(f"Motor de analisis: {motor}")

# Scrapear — df_completo conserva TODO lo scrapeado, antes de filtrar
# por tipo, para poder medir cuántas publicaciones no se pudieron leer bien.
df_completo = scrapear_modificaciones(max_paginas=max_pag, analizar_pdfs=True, solo_hoy=False)

FALLOS_PDF = {"Sin PDF", "Error descarga PDF", "PDF escaneado"}
total_scrapeadas = len(df_completo)
sin_pdf = int(df_completo['motor_analisis'].isin(FALLOS_PDF).sum()) if 'motor_analisis' in df_completo.columns else 0
if sin_pdf > 0:
    print(f"ADVERTENCIA: {sin_pdf}/{total_scrapeadas} modificatorias sin texto de PDF disponible")

# Filtrar solo ACTUALIZACION SEGURIDAD — las 10 mas recientes
df = df_completo.copy()
if 'tipo_modificacion' in df.columns:
    df = df[df['tipo_modificacion'] == 'ACTUALIZACION SEGURIDAD'].head(10).reset_index(drop=True)
print(f"Filtradas: {len(df)} modificatorias de tipo ACTUALIZACION SEGURIDAD")

exportar_excel(df, ruta)

total       = len(df)
inmediatas  = int((df['urgencia'] == 'INMEDIATA').sum())  if 'urgencia' in df.columns else 0
preventivas = int((df['urgencia'] == 'PREVENTIVA').sum()) if 'urgencia' in df.columns else 0
fecha_fmt   = datetime.now().strftime('%d/%m/%Y %H:%M')

with open(os.environ['GITHUB_OUTPUT'], 'a', encoding='utf-8') as fh:
    fh.write(f"excel_name={nombre}\n")
    fh.write(f"total_modif={total}\n")
    fh.write(f"modif_inmediatas={inmediatas}\n")
    fh.write(f"modif_preventivas={preventivas}\n")
    fh.write(f"fecha_reporte={fecha_fmt}\n")
    fh.write(f"motor_analisis={motor}\n")
    fh.write(f"sin_pdf_count={sin_pdf}\n")
    fh.write(f"total_scrapeadas={total_scrapeadas}\n")

print(f"Excel: {ruta} | Total: {total} | Inmediatas: {inmediatas} | "
      f"Preventivas: {preventivas} | Sin PDF: {sin_pdf}/{total_scrapeadas}")
