# Monitor DIGEMID — Modificatorias al Registro Sanitario

Ejecuta el scraper de **Actualizaciones de Seguridad DIGEMID** una vez por semana
(**lunes a las 8:00 AM hora Lima**), filtra las modificatorias de seguridad, las
sube a Supabase y deja en Gmail un **borrador** de correo HTML con el Excel de las
10 más recientes adjunto (no se envía solo: hay que abrirlo y pulsar "Enviar").

## Flujo

```
Lunes — 8:00 AM Lima
        |
        v
[1] Scraping DIGEMID (2 páginas ~20 posts) + análisis de PDF (Claude o heurístico)
        |
        v
[Filtro] Tipos de seguridad → todas a JSON/MD/Supabase, top 10 al Excel
        |
        v
[2] Excel raw_data/ + JSON data/ + resumen summaries/ (commit al repo)
        |
        v
[3] Borrador en Gmail (IMAP) con el Excel adjunto
    De:  conkosafe.ai@gmail.com
    Bcc: lista del secret EMAIL_TO
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
   exportadores.py                  ← JSON (data/) y resumen MD (summaries/)
   supabase_sync.py                 ← upsert a la tabla modificatorias_digemid
   draft_email.py                   ← borrador HTML en Gmail con el Excel adjunto
   notify_failure.py                ← alerta de fallo
 raw_data/  data/  summaries/       ← histórico generado por el workflow
 README.md
```

## Secrets requeridos en GitHub

**Settings → Secrets and variables → Actions → New repository secret**

| Secret | Descripción |
|--------|-------------|
| `SMTP_HOST` | Servidor SMTP — ej: `smtp.gmail.com` |
| `SMTP_PORT` | Puerto SMTP — ej: `587` |
| `SMTP_USER` | Cuenta que envía (`conkosafe.ai@gmail.com`) |
| `SMTP_PASS` | App Password de Gmail (16 caracteres; sirve también para IMAP) |
| `EMAIL_TO` | Destinatarios del borrador, separados por comas **en una sola línea**. Es secret porque incluye correos de clientes y el repo es público |
| `ANTHROPIC_API_KEY` | (Recomendado) Activa análisis semántico con Claude |
| `SUPABASE_URL` / `SUPABASE_SERVICE_ROLE_KEY` | Sincronización con la tabla `modificatorias_digemid` |

Para cambiar destinatarios: **Settings → Secrets and variables → Actions → `EMAIL_TO` → Update**.
El script descarta y avisa (sin mostrarlas en el log) las direcciones con formato inválido.

## Valores fijos (no necesitan secret)

| Parámetro | Valor |
|-----------|-------|
| Remitente | `conkosafe.ai@gmail.com` |
| Frecuencia | Lunes a las 8:00 AM hora Lima (`cron: 0 13 * * 1`) |
| Filtro | Tipos de seguridad (`TIPOS_SEGURIDAD` en `scraper.py`) — top 10 más recientes en el Excel |

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
- `max_filas`: `10` (default) — filas del Excel adjunto
- `dry_run`: `true` — solo scrapea, no crea el borrador

## Troubleshooting

| Error | Solución |
|-------|----------|
| 403 en scraping | Normal; el script reintenta con backoff automático |
| `SMTPAuthenticationError` / fallo de login IMAP | Regenerar App Password en Google (Seguridad → Contraseñas de apps) y verificar IMAP habilitado |
| `EMAIL_TO vacio` | Falta el secret `EMAIL_TO` |
| PDF escaneado | El PDF es imagen; análisis por título (heurístico) |
| Workflow no corre a las 8 AM | GitHub Actions puede tener delay de hasta 15 min |
