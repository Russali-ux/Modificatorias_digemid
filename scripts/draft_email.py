"""
draft_email.py — Deja el correo HTML de modificatorias DIGEMID como BORRADOR
en Gmail (carpeta Drafts) con el Excel adjunto. NO lo envía — hay que abrirlo
en Gmail y darle "Enviar" manualmente.

Remitente: conkosafe.ai@gmail.com
Destinatarios: van en header Bcc real (oculto entre sí al enviar desde Gmail).
Requiere que la cuenta tenga IMAP habilitado: Gmail → Configuración →
Reenvío y correo POP/IMAP → Habilitar IMAP.
"""
import os, glob, imaplib, re, time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formatdate, make_msgid
from email import encoders

# ✅ Se leen desde variables de entorno definidas en el YML
EMAIL_FROM = os.environ.get('EMAIL_FROM', 'conkosafe.ai@gmail.com')
EMAIL_TO   = os.environ['EMAIL_TO']   # lista completa separada por comas

# Mismas credenciales que usabas para SMTP — un app password de Gmail sirve
# tanto para SMTP como para IMAP (solo hay que tener IMAP habilitado en la
# cuenta). Si prefieres separarlas, agrega IMAP_USER/IMAP_PASS como secrets.
imap_host   = os.environ.get('IMAP_HOST', 'imap.gmail.com')
imap_port   = int(os.environ.get('IMAP_PORT', '993'))
imap_user   = os.environ.get('IMAP_USER', os.environ.get('SMTP_USER', EMAIL_FROM))
imap_pass   = os.environ.get('IMAP_PASS', os.environ.get('SMTP_PASS', ''))
total       = os.environ.get('TOTAL', '0')
inmediatas  = os.environ.get('INMEDIATAS', '0')
preventivas = os.environ.get('PREVENTIVAS', '0')
fecha       = os.environ.get('FECHA', '')
excel_name  = os.environ.get('EXCEL_NAME', 'modificatorias_digemid.xlsx')
motor       = os.environ.get('MOTOR', 'Heuristico')
sin_pdf     = int(os.environ.get('SIN_PDF_COUNT', '0') or '0')
total_scrapeadas = int(os.environ.get('TOTAL_SCRAPEADAS', '0') or '0')

n_inm = int(inmediatas)
n_pre = int(preventivas)
n_tot = int(total)

# Badge de estado
if n_inm > 0:
    badge_color = "#C00000"
    badge_texto = f"ATENCION: {n_inm} modificatoria(s) con accion INMEDIATA requerida"
elif n_pre > 0:
    badge_color = "#ED7D31"
    badge_texto = f"{n_pre} modificatoria(s) PREVENTIVA(S) — Evaluar en 15 dias según procedimiento de gestión de riesgos"
else:
    badge_color = "#2E75B6"
    badge_texto = f"{n_tot} actualizaciones de seguridad — Sin urgencias criticas"

motor_txt = "Claude API" if motor == "Claude API" else "Motor Heuristico"

# Aviso de fiabilidad: si alguna publicación no se pudo leer completa
# (PDF caído/bloqueado), que quede visible en el correo en vez de perderse
# en silencio del reporte.
aviso_html = ""
if sin_pdf > 0:
    aviso_html = (
        '<tr><td style="padding:0 32px 14px;">'
        '<div style="background:#FDECEC;border-left:4px solid #C00000;padding:12px 18px;'
        'border-radius:0 8px 8px 0;">'
        '<p style="margin:0;font-size:12px;color:#8B0000;">'
        f'<strong>&#9888; Aviso de fiabilidad:</strong> {sin_pdf} de {total_scrapeadas} publicaci&oacute;n(es) '
        'escaneadas no se pudieron leer completas (PDF no disponible o bloqueado por DIGEMID) y se '
        'clasificaron solo por t&iacute;tulo. Es posible que falte alguna modificatoria de seguridad en '
        'este reporte &mdash; revisar manualmente en '
        '<a href="https://www.digemid.minsa.gob.pe/webDigemid/publicaciones/alertas-modificaciones/modificaciones/" '
        'style="color:#8B0000;">digemid.minsa.gob.pe</a> si es cr&iacute;tico.'
        '</p></div></td></tr>'
    )

html = (
    '<!DOCTYPE html><html lang="es"><head>'
    '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    '</head><body style="margin:0;padding:0;background:#F0F2F5;font-family:Arial,sans-serif;">'
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#F0F2F5;padding:28px 0;">'
    '<tr><td align="center">'
    '<table width="640" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;'
    'overflow:hidden;box-shadow:0 3px 16px rgba(0,0,0,.12);">'
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
    f'<tr><td style="padding:20px 32px 0;">'
    f'<div style="background:{badge_color};color:#fff;padding:13px 20px;border-radius:8px;'
    f'font-weight:bold;font-size:14px;text-align:center;letter-spacing:.3px;">'
    f'{badge_texto}</div></td></tr>'
    f'{aviso_html}'
    '<tr><td style="padding:18px 32px;">'
    '<table width="100%" cellspacing="10" cellpadding="0"><tr>'
    '<td width="33%" style="background:#F0F5FF;border-radius:10px;padding:16px 10px;'
    'text-align:center;border-top:4px solid #1F4E79;">'
    f'<div style="font-size:36px;font-weight:bold;color:#1F4E79;">{total}</div>'
    '<div style="font-size:12px;color:#555;margin-top:5px;">Actualizaciones<br>de seguridad</div></td>'
    '<td width="33%" style="background:#FFFBF0;border-radius:10px;padding:16px 10px;'
    'text-align:center;border-top:4px solid #ED7D31;">'
    f'<div style="font-size:36px;font-weight:bold;color:#ED7D31;">{preventivas}</div>'
    '<div style="font-size:12px;color:#555;margin-top:5px;">Preventivas<br>(15 d&iacute;as h&aacute;biles)</div></td>'
    '<td width="33%" style="background:#FFF5F5;border-radius:10px;padding:16px 10px;'
    'text-align:center;border-top:4px solid #C00000;">'
    f'<div style="font-size:36px;font-weight:bold;color:#C00000;">{inmediatas}</div>'
    '<div style="font-size:12px;color:#555;margin-top:5px;">Inmediatas<br>(acci&oacute;n urgente)</div></td>'
    '</tr></table></td></tr>'
    '<tr><td style="padding:4px 32px 18px;">'
    '<div style="background:#EBF3FB;border-left:4px solid #2E75B6;padding:14px 18px;'
    'border-radius:0 8px 8px 0;">'
    '<p style="margin:0 0 4px;font-size:13px;color:#1F4E79;font-weight:bold;">'
    '&#128206; Excel adjunto a este correo</p>'
    f'<p style="margin:0;font-size:12px;color:#555;">{excel_name}</p>'
    '<p style="margin:6px 0 0;font-size:11px;color:#888;">'
    'Contiene las 10 actualizaciones de seguridad m&aacute;s recientes publicadas por DIGEMID'
    '</p></div></td></tr>'
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
    '<tr><td style="padding:0 32px 22px;">'
    '<div style="background:#FFFBF0;border-left:4px solid #ED7D31;padding:12px 18px;'
    'border-radius:0 8px 8px 0;">'
    '<p style="margin:0;font-size:12px;color:#7F4B00;">'
    '<strong>&#9201; Plazo regulatorio:</strong> Las actualizaciones de seguridad requieren evaluaci&oacute;n '
    'de gesti&oacute;n de riesgos en <strong>15 d&iacute;as h&aacute;biles</strong> y realizar las acciones '
    'mandatorios DIGEMID seg&uacute;n plazos establecidos en la resoluci&oacute;n.'
    '</p></div></td></tr>'
    '<tr><td style="background:#F4F6F9;padding:14px 32px;border-top:1px solid #E5E8ED;">'
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
    '<td style="font-size:11px;color:#999;">'
    'Enviado autom&aacute;ticamente por CONKOSAFE IA - PV INTELLIGENCE, CONKOMERCO S.A.C.'
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
asunto = f"CONKOMERCO Modificatorias ({fecha_corta})"

# Buscar Excel generado
archivos = glob.glob('/tmp/modificatorias_digemid_*.xlsx')
if not archivos:
    raise FileNotFoundError("No se encontro el Excel de modificatorias en /tmp/")
ruta_excel = sorted(archivos)[-1]

# Construir mensaje
# El "To" sigue mostrando solo el remitente. Como esto ya NO se envía por
# SMTP (no hay envelope), los destinatarios reales van en un header Bcc
# real — Gmail lo respeta y los oculta entre sí al momento en que TÚ le
# des "Enviar" manualmente sobre este borrador.
destinatarios = [e.strip() for e in EMAIL_TO.split(',') if e.strip()]

msg = MIMEMultipart('mixed')
msg['Subject']    = asunto
msg['From']       = f"Monitor DIGEMID CONKOMERCO <{EMAIL_FROM}>"
msg['To']         = EMAIL_FROM
msg['Bcc']        = ', '.join(destinatarios)
msg['Reply-To']   = EMAIL_FROM
msg['Date']       = formatdate(localtime=True)
msg['Message-ID'] = make_msgid()
msg.attach(MIMEText(html, 'html'))

# Adjuntar Excel
with open(ruta_excel, 'rb') as f:
    adjunto = MIMEBase('application',
                       'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    adjunto.set_payload(f.read())
encoders.encode_base64(adjunto)
adjunto.add_header('Content-Disposition', f'attachment; filename="{excel_name}"')
msg.attach(adjunto)


def _encontrar_carpeta_borradores(imap: imaplib.IMAP4_SSL) -> str:
    """Busca la carpeta con el atributo especial \\Drafts (RFC 6154) en vez de
    asumir el nombre en inglés — si la cuenta de Gmail está en español, la
    carpeta real puede llamarse '[Gmail]/Borradores' y no '[Gmail]/Drafts'."""
    typ, folders = imap.list()
    if typ == 'OK' and folders:
        for f in folders:
            linea = f.decode(errors='ignore') if isinstance(f, bytes) else f
            if '\\Drafts' in linea:
                m = re.search(r'"([^"]+)"\s*$', linea) or re.search(r'(\S+)\s*$', linea)
                if m:
                    return m.group(1).strip('"')
    return '[Gmail]/Drafts'  # fallback si no se pudo detectar


with imaplib.IMAP4_SSL(imap_host, imap_port) as imap:
    imap.login(imap_user, imap_pass)
    carpeta = _encontrar_carpeta_borradores(imap)
    typ, _ = imap.append(
        carpeta,
        '\\Draft',
        imaplib.Time2Internaldate(time.time()),
        msg.as_bytes(),
    )
    imap.logout()

if typ != 'OK':
    raise RuntimeError(f"No se pudo guardar el borrador en '{carpeta}' (respuesta IMAP: {typ})")

print(f"Borrador guardado en '{carpeta}': {EMAIL_FROM} (Bcc oculto: {len(destinatarios)} destinatarios)")
print(f"Asunto: {asunto}")
print("⚠️  El correo NO fue enviado — ábrelo en Gmail y dale Enviar manualmente.")