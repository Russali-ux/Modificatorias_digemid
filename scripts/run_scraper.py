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
ruta    = f"/tmp/{nombre}"

api_key = os.environ.get('ANTHROPIC_API_KEY', '')
motor   = 'Claude API' if (api_key and api_key.startswith('sk-')) else 'Heuristico'
print(f"Motor de analisis: {motor}")

# Scrapear
df = scrapear_modificaciones(max_paginas=max_pag, analizar_pdfs=True, solo_hoy=False)

# Filtrar solo ACTUALIZACION SEGURIDAD — las 10 mas recientes
if 'tipo_modificacion' in df.columns:
    df = df[df['tipo_modificacion'] == 'ACTUALIZACION SEGURIDAD'].head(10).reset_index(drop=True)
print(f"Filtradas: {len(df)} modificatorias de tipo ACTUALIZACION SEGURIDAD")

exportar_excel(df, ruta)

total       = len(df)
inmediatas  = int((df['urgencia'] == 'INMEDIATA').sum())  if 'urgencia' in df.columns else 0
preventivas = int((df['urgencia'] == 'PREVENTIVA').sum()) if 'urgencia' in df.columns else 0
fecha_fmt   = datetime.now().strftime('%d/%m/%Y %H:%M')

with open(os.environ['GITHUB_OUTPUT'], 'a') as fh:
    fh.write(f"excel_name={nombre}\n")
    fh.write(f"total_modif={total}\n")
    fh.write(f"modif_inmediatas={inmediatas}\n")
    fh.write(f"modif_preventivas={preventivas}\n")
    fh.write(f"fecha_reporte={fecha_fmt}\n")
    fh.write(f"motor_analisis={motor}\n")

print(f"Excel: {ruta} | Total: {total} | Inmediatas: {inmediatas} | Preventivas: {preventivas}")
