"""
exportadores.py
---------------
Convierte el DataFrame de modificatorias (el mismo que arma scraper.py y que
exportar_excel() vuelca al Excel) a los dos formatos que se commitean al repo
como historial versionado y que alimentan a Supabase / al resumen legible:

    data/modificatorias_{YYYYMMDD}.json   <- snapshot del día (histórico)
    data/modificatorias_latest.json       <- snapshot más reciente (fuente para supabase_sync.py)
    summaries/resumen_{YYYY-MM-DD}.md     <- resumen legible para revisión humana / auditoría

El Excel (raw_data/modificatorias_digemid_{fecha}.xlsx) sigue siendo el
archivo histórico "crudo" tal como sale del scraper; JSON y MD son
derivados de esa misma corrida, no reemplazos.
"""

import json
import os
from datetime import datetime

import pandas as pd

EMOJI_URGENCIA = {"INMEDIATA": "🔴", "PREVENTIVA": "🟡", "INFORMATIVA": "🔵"}


def _serializable(valor):
    """Convierte tipos de pandas/datetime a algo que json.dumps entienda."""
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(valor, "isoformat"):
        return valor.isoformat()
    return valor


def exportar_json(df: pd.DataFrame, ruta_dia: str, ruta_latest: str) -> list:
    """Escribe data/modificatorias_{fecha}.json y data/modificatorias_latest.json.
    Devuelve la lista de registros (dict) tal como quedó escrita, para que
    run_scraper.py pueda pasársela directo a subir_a_supabase() sin releer el
    archivo."""
    os.makedirs(os.path.dirname(ruta_dia), exist_ok=True)
    registros = [
        {k: _serializable(v) for k, v in fila.items()}
        for fila in df.to_dict("records")
    ]
    for ruta in (ruta_dia, ruta_latest):
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(registros, f, ensure_ascii=False, indent=2)
    print(f"💾 JSON guardado: {ruta_dia}")
    print(f"💾 JSON guardado: {ruta_latest}")
    return registros


def exportar_resumen_md(df: pd.DataFrame, ruta: str, fecha_reporte: str = None) -> None:
    """Escribe summaries/resumen_{fecha}.md — un resumen legible para revisión
    humana, mismo estilo que los resumenes de Alertas-analyzer."""
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    fecha_reporte = fecha_reporte or datetime.now().strftime("%Y/%m/%d")
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")

    lineas = [
        f"# Resumen Modificatorias DIGEMID — {fecha_reporte}",
        "",
        f"**Total modificatorias:** {len(df)}  ",
        f"**Generado:** {ahora} (Lima, PE)",
        "",
    ]

    if "urgencia" in df.columns and len(df):
        lineas.append("## Por urgencia")
        lineas.append("| Urgencia | N° |")
        lineas.append("|----------|-----|")
        for urg, n in df["urgencia"].value_counts().items():
            emoji = EMOJI_URGENCIA.get(urg, "⚪")
            lineas.append(f"| {emoji} {urg} | {n} |")
        lineas.append("")

    lineas.append("## Modificatorias del reporte")
    lineas.append("")

    for _, fila in df.iterrows():
        titulo = fila.get("n_modificacion", "—")
        pdf_url = fila.get("pdf_url")
        encabezado = f"### {titulo}"
        if pdf_url:
            encabezado += f" — [📄 PDF]({pdf_url})"
        lineas.append(encabezado)
        if fila.get("producto"):
            lineas.append(f"- **Producto/IFA:** {fila.get('producto')}")
        if fila.get("titular_rs"):
            lineas.append(f"- **Titular RS:** {fila.get('titular_rs')}")
        lineas.append(f"- **Tipo:** {fila.get('tipo_modificacion', '—')}")
        lineas.append(f"- **Urgencia:** {fila.get('urgencia', '—')}")
        if fila.get("accion_requerida"):
            lineas.append(f"- **Acción Requerida:** {fila.get('accion_requerida')}")
        if fila.get("resumen"):
            lineas.append(f"- **Resumen IA:** {fila.get('resumen')}")
        lineas.append("")

    with open(ruta, "w", encoding="utf-8") as f:
        f.write("\n".join(lineas))
    print(f"📝 Resumen MD guardado: {ruta}")
