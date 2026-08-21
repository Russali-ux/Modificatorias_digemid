"""
supabase_sync.py
-----------------
Sube (upsert) las modificatorias DIGEMID ya procesadas por scraper.py
hacia la tabla `modificatorias_digemid` en Supabase (mismo proyecto que
Alertas-analyzer: ggbnfdaxtsngsjssrwrl).

Se asume que `df` es el DataFrame final que arma scrapear_y_filtrar_modificatorias()
en scraper.py, con columnas (snake_case):
    n_modificacion, producto, principio_activo, titular_rs, fecha_publicacion,
    tipo_modificacion, urgencia, accion_requerida, indicador_tiempos, base_legal,
    resumen, motor_analisis, url, pdf_url, fecha_captura

Requiere:
    pip install supabase --break-system-packages

Variables de entorno requeridas (agregar como Secrets en GitHub Actions):
    SUPABASE_URL              -> https://ggbnfdaxtsngsjssrwrl.supabase.co
    SUPABASE_SERVICE_ROLE_KEY -> Settings > API > service_role (NUNCA la anon key aquí)
"""

import os
import json
import pandas as pd
from supabase import create_client, Client


def _get_client() -> Client:
    url = os.environ["SUPABASE_URL"]
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    return create_client(url, key)


def _limpio(valor):
    """Convierte NaN/NaT de pandas a None; deja el resto intacto."""
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    return valor


def _fila_a_registro(fila: pd.Series) -> dict:
    """Convierte una fila del DataFrame (columnas reales de scraper.py) al formato
    de la tabla modificatorias_digemid."""
    fecha_pub = _limpio(fila.get("fecha_publicacion"))
    if fecha_pub is not None:
        fecha_pub = pd.to_datetime(fecha_pub).strftime("%Y-%m-%d")

    fecha_cap = _limpio(fila.get("fecha_captura"))
    if fecha_cap is not None:
        fecha_cap = pd.to_datetime(fecha_cap).isoformat()

    return {
        "n_modificacion":      _limpio(fila.get("n_modificacion")),
        "producto":            _limpio(fila.get("producto")),
        "principio_activo":    _limpio(fila.get("principio_activo")),
        "titular_rs":          _limpio(fila.get("titular_rs")),
        "fecha_publicacion":   fecha_pub,
        "tipo_modificacion":   _limpio(fila.get("tipo_modificacion")),
        "urgencia":            _limpio(fila.get("urgencia")),
        "accion_requerida":    _limpio(fila.get("accion_requerida")),
        "indicador_tiempos":   _limpio(fila.get("indicador_tiempos")),
        "base_legal":          _limpio(fila.get("base_legal")),
        "resumen":             _limpio(fila.get("resumen")),
        "motor_analisis":      _limpio(fila.get("motor_analisis")),
        "url_pagina_digemid":  _limpio(fila.get("url")),
        "url_pdf_digemid":     _limpio(fila.get("pdf_url")),
        "fecha_captura":       fecha_cap,
    }


def subir_a_supabase(df: pd.DataFrame) -> int:
    """Sube todas las filas del DataFrame a Supabase con upsert por
    (n_modificacion, fecha_publicacion). Devuelve la cantidad de filas subidas."""
    if df is None or df.empty:
        print("⚠️  DataFrame vacío, no hay nada que subir a Supabase.")
        return 0

    if "SUPABASE_URL" not in os.environ or "SUPABASE_SERVICE_ROLE_KEY" not in os.environ:
        print("⚠️  SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no configuradas; se omite la sincronización.")
        return 0

    supabase = _get_client()
    registros = [_fila_a_registro(fila) for _, fila in df.iterrows()]
    return _upsert(supabase, registros)


def subir_desde_json(ruta_json: str) -> int:
    """Sube a Supabase directamente desde un archivo JSON ya escrito por
    exportadores.exportar_json() (data/modificatorias_{fecha}.json o
    data/modificatorias_latest.json). Es la vía preferida en el pipeline:
    Excel (archivo crudo) -> JSON (commiteado, versionado) -> Supabase."""
    with open(ruta_json, encoding="utf-8") as f:
        registros_json = json.load(f)
    if not registros_json:
        print(f"⚠️  {ruta_json} está vacío, no hay nada que subir a Supabase.")
        return 0

    if "SUPABASE_URL" not in os.environ or "SUPABASE_SERVICE_ROLE_KEY" not in os.environ:
        print("⚠️  SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no configuradas; se omite la sincronización.")
        return 0

    # Las claves del JSON ya son las del DataFrame (n_modificacion, producto, ...);
    # se pasan por pd.Series para reusar _fila_a_registro tal cual.
    registros = [_fila_a_registro(pd.Series(r)) for r in registros_json]
    supabase = _get_client()
    return _upsert(supabase, registros)


def _upsert(supabase: Client, registros: list) -> int:
    registros_validos = [r for r in registros if r.get("n_modificacion") and r.get("fecha_publicacion")]
    omitidos = len(registros) - len(registros_validos)
    if omitidos:
        print(f"⚠️  {omitidos} fila(s) sin n° de modificación o fecha; se omiten de la subida a Supabase.")

    if not registros_validos:
        print("⚠️  No hay registros válidos para subir a Supabase.")
        return 0

    resultado = (
        supabase.table("modificatorias_digemid")
        .upsert(registros_validos, on_conflict="n_modificacion,fecha_publicacion")
        .execute()
    )
    n = len(resultado.data)
    print(f"✅ Supabase: {n} registros insertados/actualizados en modificatorias_digemid.")
    return n


if __name__ == "__main__":
    # Prueba manual: python scripts/supabase_sync.py output/modificatorias_digemid_YYYYMMDD.xlsx
    import sys

    ruta = sys.argv[1] if len(sys.argv) > 1 else None
    if not ruta:
        print("Uso: python scripts/supabase_sync.py <ruta_al_excel_o_csv>")
        sys.exit(1)

    if ruta.endswith(".xlsx"):
        df_prueba = pd.read_excel(ruta, sheet_name="Modificatorias DIGEMID")
    else:
        df_prueba = pd.read_csv(ruta)
    subir_a_supabase(df_prueba)
