"""
send_email.py — Envía correo HTML con Excel de modificatorias DIGEMID adjunto.
Remitente: conkosafe.ai@gmail.com
"""
import os, glob, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

EMAIL_FROM = "conkosafe.ai@gmail.com"
EMAIL_TO   = (
    "finanzas@conkomerco.com,"
    "conkosafe.ai@gmail.com,"
    "july,maita@conkomerco.com,"
    "alex.rodriguez@conkomerco.com"
    "angelica.aguero@conkomerco.com"
    "giancarlos.chafloque@conkomerco.com"
)

smtp_host   = os.environ['SMTP_HOST']
smtp_port   = int(os.environ.get('SMTP_PORT', '587'))
smtp_user   = os.environ['SMTP_USER']
smtp_pass   = os.environ['SMTP_PASS']
total       = os.environ.get('TOTAL', '0')
inmediatas  = os.environ.get('INMEDIATAS', '0')
preventivas = os.environ.get('PREVENTIVAS', '0')
fecha       = os.environ.get('FECHA', '')
excel_name  = os.environ.get('EXCEL_NAME', 'modificatorias_digemid.xlsx')
motor       = os.environ.get('MOTOR', 'Heuristico')

n_inm = int(inmediatas)
n_pre = int(preventivas)
n_tot = int(total)

# Badge de estado
if n_inm > 0:
    badge_color = "#C00000"
    badge_texto = f"ATENCION: {n_inm} modificatoria(s) con accion INMEDIATA requerida"
elif n_pre > 0:
    badge_color = "#ED7D31"
    badge_texto = f"{n_pre} modificatoria(s) PREVENTIVA(S) — Evaluar en 15 dias habiles"
else:
    badge_color = "#2E75B6"
    badge_texto = f"{n_tot} actualizaciones de seguridad — Sin urgencias criticas"

motor_txt = "Claude API" if motor == "Claude API" else "Motor Heuristico"

html = (
    '<!DOCTYPE html><html lang="es"><head>'
    '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    '</head><body style="margin:0;padding:0;background:#F0F2F5;font-family:Arial,sans-serif;">'

    # Contenedor externo
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#F0F2F5;padding:28px 0;">'
    '<tr><td align="center">'
    '<table width="640" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;'
    'overflow:hidden;box-shadow:0 3px 16px rgba(0,0,0,.12);">'

    # Header
    '<tr><td style="background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);padding:26px 32px;">'
    '<table width="100%" cellpadding="0" cellspacing="0"><tr><td>'
    '<p style="margin:0;color:#BDD7EE;font-size:11px;text-transform:uppercase;letter-spacing:1.2px;">'
    'CONKOSAFE IA &mdash; PV Intelligence</p>'
    '<h1 style="margin:6px 0 0;color:#fff;font-size:21px;font-weight:bold;line-height:1.3;">'
    'Modificaciones por seguridad Registro Sanitario</h1>'
    '<p style="margin:4px 0 0;color:#BDD7EE;font-size:12px;">'
    f'Actualizaciones de Seguridad DIGEMID &mdash; {fecha} hora Lima</p>'
    '</td>'
    f'<td align="right" valign="middle" style="padding-left:16px;">'
    f'<span style="background:rgba(255,255,255,.15);color:#fff;font-size:11px;'
    f'padding:5px 12px;border-radius:20px;white-space:nowrap;">{motor_txt}</span>'
    '</td></tr></table></td></tr>'

    # Badge estado
    f'<tr><td style="padding:20px 32px 0;">'
    f'<div style="background:{badge_color};color:#fff;padding:13px 20px;border-radius:8px;'
    f'font-weight:bold;font-size:14px;text-align:center;letter-spacing:.3px;">'
    f'{badge_texto}</div></td></tr>'

    # Tarjetas KPI
    '<tr><td style="padding:18px 32px;">'
    '<table width="100%" cellspacing="10" cellpadding="0"><tr>'

    # KPI Total
    '<td width="33%" style="background:#F0F5FF;border-radius:10px;padding:16px 10px;'
    'text-align:center;border-top:4px solid #1F4E79;">'
    f'<div style="font-size:36px;font-weight:bold;color:#1F4E79;">{total}</div>'
    '<div style="font-size:12px;color:#555;margin-top:5px;">Actualizaciones<br>de seguridad</div></td>'

    # KPI Preventivas
    '<td width="33%" style="background:#FFFBF0;border-radius:10px;padding:16px 10px;'
    'text-align:center;border-top:4px solid #ED7D31;">'
    f'<div style="font-size:36px;font-weight:bold;color:#ED7D31;">{preventivas}</div>'
    '<div style="font-size:12px;color:#555;margin-top:5px;">Preventivas<br>(15 d&iacute;as h&aacute;biles)</div></td>'

    # KPI Inmediatas
    '<td width="33%" style="background:#FFF5F5;border-radius:10px;padding:16px 10px;'
    'text-align:center;border-top:4px solid #C00000;">'
    f'<div style="font-size:36px;font-weight:bold;color:#C00000;">{inmediatas}</div>'
    '<div style="font-size:12px;color:#555;margin-top:5px;">Inmediatas<br>(acci&oacute;n urgente)</div></td>'

    '</tr></table></td></tr>'

    # Excel adjunto
    '<tr><td style="padding:4px 32px 18px;">'
    '<div style="background:#EBF3FB;border-left:4px solid #2E75B6;padding:14px 18px;'
    'border-radius:0 8px 8px 0;">'
    '<p style="margin:0 0 4px;font-size:13px;color:#1F4E79;font-weight:bold;">'
    '&#128206; Excel adjunto a este correo</p>'
    f'<p style="margin:0;font-size:12px;color:#555;">{excel_name}</p>'
    '<p style="margin:6px 0 0;font-size:11px;color:#888;">'
    'Contiene las 10 actualizaciones de seguridad m&aacute;s recientes publicadas por DIGEMID'
    '</p></div></td></tr>'

    # Pasos de revisión
    '<tr><td style="padding:0 32px 22px;">'
    '<div style="background:#F4FBF0;border-left:4px solid #375623;padding:14px 18px;'
    'border-radius:0 8px 8px 0;">'
    '<p style="margin:0 0 8px;font-size:13px;color:#375623;font-weight:bold;">'
    '&#9989; Pasos de revisi&oacute;n recomendados</p>'
    '<ol style="margin:0;padding-left:18px;font-size:13px;color:#333;line-height:2;">'
    '<li>Abrir el Excel adjunto &mdash; columna <strong>Urgencia</strong> ya viene coloreada</li>'
    '<li>Revisar <strong>Acci&oacute;n Requerida</strong> y <strong>Indicador Tiempos</strong> (verde)</li>'
    '<li>Identificar los productos/IFAs de tu portafolio en columna <strong>Producto / IFA</strong></li>'
    '<li>Comunicar a Direcci&oacute;n T&eacute;cnica / Asuntos Regulatorios / Calidad seg&uacute;n corresponda</li>'
    '<li>Registrar evaluaci&oacute;n en el sistema de farmacovigilancia si aplica (D.S. 016-2011-SA Art. 55)</li>'
    '</ol></div></td></tr>'

    # Nota regulatoria
    '<tr><td style="padding:0 32px 22px;">'
    '<div style="background:#FFFBF0;border-left:4px solid #ED7D31;padding:12px 18px;'
    'border-radius:0 8px 8px 0;">'
    '<p style="margin:0;font-size:12px;color:#7F4B00;">'
    '<strong>&#9201; Plazo regulatorio:</strong> Las actualizaciones de seguridad requieren evaluaci&oacute;n '
    'e informe a DIGEMID en <strong>15 d&iacute;as h&aacute;biles</strong> desde la publicaci&oacute;n '
    '(Art. 55 D.S. 016-2011-SA / ICH E2C).'
    '</p></div></td></tr>'

    # Footer
    '<tr><td style="background:#F4F6F9;padding:14px 32px;border-top:1px solid #E5E8ED;">'
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
    '<td style="font-size:11px;color:#999;">'
    'Enviado autom&aacute;ticamente por Monitor DIGEMID &mdash; CONKOMERCO S.A.C. &mdash; via GitHub Actions'
    '</td>'
    '<td align="right">'
    '<a href="https://www.digemid.minsa.gob.pe/webDigemid/publicaciones/alertas-modificaciones/modificaciones/" '
    'style="color:#1F4E79;text-decoration:none;font-size:11px;">digemid.minsa.gob.pe</a>'
    '</td></tr></table></td></tr>'

    '</table></td></tr></table>'
    '</body></html>'
)

# Asunto
fecha_corta = fecha[:10] if fecha else ""
if n_inm > 0:
    asunto = f"[DIGEMID MOD {fecha_corta}] {n_inm} modificatoria(s) INMEDIATA(S) — Accion requerida"
elif n_pre > 0:
    asunto = f"[DIGEMID MOD {fecha_corta}] {n_pre} actualizacion(es) de seguridad — Evaluar en 15 dias"
else:
    asunto = f"[DIGEMID MOD {fecha_corta}] {n_tot} actualizaciones de seguridad — Sin urgencias"

# Buscar Excel generado
archivos = glob.glob('/tmp/modificatorias_digemid_*.xlsx')
if not archivos:
    raise FileNotFoundError("No se encontro el Excel de modificatorias en /tmp/")
ruta_excel = sorted(archivos)[-1]

# Construir mensaje
msg = MIMEMultipart('mixed')
msg['Subject']  = asunto
msg['From']     = f"Monitor DIGEMID CONKOMERCO <{EMAIL_FROM}>"
msg['To']       = EMAIL_TO
msg['Reply-To'] = EMAIL_FROM
msg.attach(MIMEText(html, 'html'))

# Adjuntar Excel
with open(ruta_excel, 'rb') as f:
    adjunto = MIMEBase('application',
                       'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    adjunto.set_payload(f.read())
encoders.encode_base64(adjunto)
adjunto.add_header('Content-Disposition', f'attachment; filename="{excel_name}"')
msg.attach(adjunto)

destinatarios = [e.strip() for e in EMAIL_TO.split(',')]

with smtplib.SMTP(smtp_host, smtp_port) as server:
    server.ehlo()
    server.starttls()
    server.login(smtp_user, smtp_pass)
    server.sendmail(EMAIL_FROM, destinatarios, msg.as_string())

print(f"Correo enviado: {EMAIL_FROM} -> {EMAIL_TO}")
print(f"Asunto: {asunto}")
