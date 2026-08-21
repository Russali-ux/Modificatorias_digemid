import os, sys
from datetime import datetime

exec(open('scripts/scraper.py').read())

max_pag = int(os.environ.get('MAX_PAGINAS', '1'))
fecha   = datetime.now().strftime('%Y%m%d_%H%M')
nombre  = f"alertas_digemid_{fecha}.xlsx"
ruta    = f"/tmp/{nombre}"

api_key = os.environ.get('ANTHROPIC_API_KEY', '')
motor   = 'Claude API' if (api_key and api_key.startswith('sk-')) else 'Heuristico'
print(f"Motor de analisis: {motor}")

df = scrapear_alertas(max_paginas=max_pag, analizar_acciones=True)
exportar_excel(df, ruta)

total       = len(df)
inmediatas  = int((df['urgencia'] == 'INMEDIATA').sum())  if 'urgencia' in df.columns else 0
preventivas = int((df['urgencia'] == 'PREVENTIVA').sum()) if 'urgencia' in df.columns else 0
fecha_fmt   = datetime.now().strftime('%d/%m/%Y %H:%M')

with open(os.environ['GITHUB_OUTPUT'], 'a') as fh:
    fh.write(f"excel_name={nombre}\n")
    fh.write(f"total_alertas={total}\n")
    fh.write(f"alertas_inmediatas={inmediatas}\n")
    fh.write(f"alertas_preventivas={preventivas}\n")
    fh.write(f"fecha_reporte={fecha_fmt}\n")
    fh.write(f"motor_analisis={motor}\n")

print(f"Excel: {ruta} | Total: {total} | Inmediatas: {inmediatas} | Preventivas: {preventivas}")
