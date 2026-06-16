# Monitor DIGEMID — Modificatorias al Registro Sanitario

Ejecuta el scraper de **Actualizaciones de Seguridad DIGEMID** tres veces por semana
(Lun / Mié / Vie a las 8:00 AM hora Lima), extrae las **10 modificatorias más recientes**
de tipo `ACTUALIZACION SEGURIDAD` y envía un correo HTML con el Excel adjunto.

## Flujo

```
Lun/Mié/Vie — 8:00 AM Lima
        |
        v
[1] Scraping DIGEMID (2 páginas ~20 posts)
        |
        v
[Filtro] ACTUALIZACION SEGURIDAD → top 10
        |
        v
[2] Generar Excel  →  modificatorias_digemid_YYYYMMDD_HHMM.xlsx
        |
        v
[3] Correo HTML con Excel adjunto
    De:   conkosafe.ai@gmail.com
    Para: july.maita / finanzas / alex.rodriguez
          angelica.aguero / giancarlos.chafloque
        |
        v
[4] Si falla → correo de error automático
```

## Estructura del repositorio

```
modificatorias_digemid/
 .gitattributes
 .github/
   workflows/
     monitor_modificatorias.yml     ← workflow principal
 scripts/
   scraper.py                       ← scraper + exportador Excel
   run_scraper.py                   ← orquestador GitHub Actions
   send_email.py                    ← correo HTML con adjunto
   notify_failure.py                ← alerta de fallo
 README.md
```

## Secrets requeridos en GitHub

**Settings → Secrets and variables → Actions → New repository secret**

| Secret | Descripción |
|--------|-------------|
| `SMTP_HOST` | Servidor SMTP — ej: `smtp.gmail.com` |
| `SMTP_PORT` | Puerto SMTP — ej: `587` |
| `SMTP_USER` | Cuenta que envía (`conkosafe.ai@gmail.com`) |
| `SMTP_PASS` | App Password de Gmail (16 caracteres) |
| `ANTHROPIC_API_KEY` | (Recomendado) Activa análisis semántico con Claude |

## Valores fijos (no necesitan secret)

| Parámetro | Valor |
|-----------|-------|
| Remitente | `conkosafe.ai@gmail.com` |
| Destinatarios | `july.maita`, `finanzas`, `alex.rodriguez`, `angelica.aguero`, `giancarlos.chafloque` @conkomerco.com |
| Frecuencia | Lun/Mié/Vie a las 8:00 AM hora Lima |
| Filtro | Solo `ACTUALIZACION SEGURIDAD` — top 10 más recientes |

## Excel generado

El Excel contiene **2 hojas**:

| Hoja | Contenido |
|------|-----------|
| `Modificatorias DIGEMID` | Tabla principal con colores por urgencia, indicador de tiempos (verde), hipervínculos a PDF |
| `Resumen Diario` | Conteos por urgencia, tipo y productos más frecuentes |

### Columnas principales

| Columna | Descripción |
|---------|-------------|
| N° Modificación | Ej: `MODIFICACIONES N° 08 – 2026` |
| Producto / IFA | Nombre del producto o IFA afectado |
| Principio Activo | IFA principal extraído del PDF |
| Titular RS | Laboratorio titular |
| Fecha Publicación | Fecha oficial DIGEMID |
| Tipo de Modificación | `ACTUALIZACION SEGURIDAD` (filtrado) |
| Urgencia | 🔴 INMEDIATA / 🟡 PREVENTIVA / 🔵 INFORMATIVA |
| Acción Requerida | Qué debe hacer el Titular del RS |
| ⏱ Indicador Tiempos | Plazo regulatorio D.S. 016-2011-SA |
| Resumen IA | Síntesis del PDF (con API key) |
| URL PDF | Hipervínculo directo al documento |

## Motor de análisis

| Campo | Sin `ANTHROPIC_API_KEY` | Con `ANTHROPIC_API_KEY` |
|-------|------------------------|------------------------|
| Tipo Modificación | Palabras clave del título | Análisis semántico del PDF |
| Principio Activo | Regex en PDF | Extraído por Claude |
| Titular RS | Regex en PDF | Extraído por Claude |
| Resumen IA | Vacío | Párrafo por modificatoria |
| Motor | `Heurístico` | `Claude API` |

## Plazo regulatorio

> Las Actualizaciones de Seguridad requieren evaluación e informe a DIGEMID
> en **15 días hábiles** desde la publicación.
> **Base legal:** Art. 55 D.S. 016-2011-SA / ICH E2C(R2).

## Probar manualmente

**Actions → Monitor DIGEMID - Modificatorias Actualizacion Seguridad → Run workflow**

Parámetros opcionales:
- `max_paginas`: `2` (default) — páginas del listado DIGEMID (~10 posts/página)
- `dry_run`: `true` — solo scrapea, no envía correo

## Troubleshooting

| Error | Solución |
|-------|----------|
| 403 en scraping | Normal; el script reintenta con backoff automático |
| `SMTPAuthenticationError` | Regenerar App Password en Google (Seguridad → Contraseñas de apps) |
| PDF escaneado | El PDF es imagen; análisis por título (heurístico) |
| Workflow no corre a las 8 AM | GitHub Actions puede tener delay de hasta 15 min |
