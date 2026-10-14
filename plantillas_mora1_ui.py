"""Editor visual para las cinco plantillas de Mora 1. Nunca envia correos."""
import html
import json
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import streamlit as st
import streamlit.components.v1 as components
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "1ELFN7ONft65RWeeHW5j7e5kvIyzOiG0UFWVeHo9PTLg"
SHEET = "Plantillas_Mora1"
HEADERS = ["CODIGO", "ASUNTO", "HTML", "ACTUALIZADO"]
TZ = ZoneInfo("America/Bogota")
CODES = ["MORA101", "MORA102", "MORA103", "MORA104", "MORA105"]
DAYS = {"MORA101": 4, "MORA102": 10, "MORA103": 15, "MORA104": 19, "MORA105": 24}

SUBJECTS = {
    "MORA101": "Recordatorio sobre tu mensualidad pendiente",
    "MORA102": "Seguimiento a tu mensualidad",
    "MORA103": "Queremos ayudarte con tu pago pendiente",
    "MORA104": "Seguimos atentos a tu mensualidad",
    "MORA105": "Comunícate con nosotros sobre tu pago",
}
MESSAGES = {
    "MORA101": "Queremos recordarte que, de acuerdo con nuestros registros, tienes una mensualidad pendiente. Si ya realizaste el pago, puedes responder a este mensaje para que revisemos tu información.",
    "MORA102": "Te escribimos para hacer seguimiento a la mensualidad que continúa pendiente en nuestros registros. Sabemos que pueden presentarse novedades y queremos conocer tu situación.",
    "MORA103": "Hemos intentado mantener el seguimiento de tu mensualidad y queremos encontrar contigo la mejor manera de resolver cualquier inconveniente.",
    "MORA104": "Tu mensualidad sigue apareciendo pendiente en nuestros registros. Nuestro equipo está disponible para revisar tu caso y resolver cualquier novedad.",
    "MORA105": "Queremos dar continuidad al seguimiento de tu mensualidad pendiente y conocer si necesitas apoyo para regularizar tu situación.",
}
CLOSINGS = {
    "MORA101": "Estamos disponibles para ayudarte si tienes alguna inquietud.",
    "MORA102": "Si ya realizaste el pago, por favor confírmanoslo respondiendo este mensaje.",
    "MORA103": "Si necesitas orientación o ya realizaste el pago, cuéntanos por este medio.",
    "MORA104": "Si tienes una fecha de pago acordada o ya pagaste, puedes informarnos para revisar tu gestión.",
    "MORA105": "Por favor responde a este mensaje o comunícate por nuestros canales habituales.",
}

def default_html(code):
    """HTML de correo con estilos inline para compatibilidad con clientes de correo."""
    message = html.escape(MESSAGES[code])
    closing = html.escape(CLOSINGS[code])
    subject = html.escape(SUBJECTS[code])
    day = DAYS[code]
    return f'''<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:24px;background-color:#edf2f7;font-family:Arial,Helvetica,sans-serif;color:#24364b">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:620px;margin:0 auto;background:#ffffff;border:1px solid #dfe6ed;border-collapse:separate">
<tr><td style="background:#123456;color:#ffffff;padding:28px 32px">
<span style="font-size:29px;font-weight:800;letter-spacing:2px">BRAVO</span><br>
<span style="font-size:11px;letter-spacing:1px;color:#d7e7f3">ACOMPAÑAMIENTO A TU PROCESO</span>
</td></tr>
<tr><td style="background:#16b4a0;height:5px;font-size:1px">&nbsp;</td></tr>
<tr><td style="padding:32px">
<p style="font-size:11px;letter-spacing:1px;color:#687a8a;margin-top:0">SEGUIMIENTO DE MENSUALIDAD · DÍA {day}</p>
<h1 style="font-size:24px;line-height:1.3;color:#123456;margin:0 0 25px">{subject}</h1>
<p style="font-size:15px;line-height:1.7">Hola, <strong>{{{{nombre}}}}</strong>:</p>
<p style="font-size:15px;line-height:1.7">{message}</p>
<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f3f7f8;border-left:4px solid #16b4a0;margin:22px 0"><tr><td style="padding:15px 18px">
<span style="font-size:11px;color:#667888">REFERENCIA DEL CLIENTE</span><br>
<strong style="font-size:16px">{{{{referencia}}}}</strong>
</td></tr></table>
<p style="font-size:15px;line-height:1.7">{closing}</p>
<p style="font-size:15px;line-height:1.7">Gracias por mantenernos informados.</p>
<p style="font-size:15px;line-height:1.7;margin-top:30px">Cordialmente,<br><strong>Equipo BRAVO</strong></p>
</td></tr>
<tr><td style="background:#f2f5f7;padding:18px 30px;color:#647581;font-size:12px;line-height:1.5">
<strong>BRAVO</strong> · Comunicación de seguimiento.<br>
Si realizaste el pago recientemente, es posible que todavía no se refleje en nuestros registros. Este mensaje no constituye una confirmación de saldo.
</td></tr>
</table></body></html>'''

def secret_value(name):
    if os.getenv(name):
        return os.environ[name]
    try:
        return st.secrets.get(name)
    except Exception:
        return None

def open_book():
    raw = secret_value("MI_JSON")
    if not raw:
        raise RuntimeError("No se configuró MI_JSON en los secretos de Streamlit.")
    info = json.loads(raw) if isinstance(raw, str) else dict(raw)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

def load_saved():
    try:
        book = open_book()
        ws = book.worksheet(SHEET)
        data = ws.get_all_values()
        if not data:
            return {}, None
        h = {n.strip().upper(): i for i, n in enumerate(data[0])}
        if not {"CODIGO", "ASUNTO", "HTML"}.issubset(h):
            raise RuntimeError("La hoja Plantillas_Mora1 tiene encabezados incompatibles.")
        saved = {}
        for r in data[1:]:
            def get(n):
                i = h[n]
                return r[i] if i < len(r) else ""
            code = get("CODIGO").strip().upper()
            if code in CODES:
                saved[code] = {"subject": get("ASUNTO"), "body": get("HTML")}
        return saved, None
    except gspread.WorksheetNotFound:
        return {}, None
    except Exception as exc:
        return {}, str(exc)

def save_template(code, subject, body):
    book = open_book()
    try:
        ws = book.worksheet(SHEET)
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(title=SHEET, rows=100, cols=4)
        ws.update(range_name="A1:D1", values=[HEADERS])
    existing = ws.get_all_values()
    if existing:
        header = [x.strip().upper() for x in existing[0]]
        if header[:4] != HEADERS:
            raise RuntimeError("Encabezados inesperados: se evitó sobrescribir la hoja.")
    else:
        ws.update(range_name="A1:D1", values=[HEADERS])
    target = next(
        (i for i, row in enumerate(existing[1:], start=2)
         if row and row[0].strip().upper() == code),
        len(existing) + 1 if existing else 2,
    )
    value = [code, subject, body, datetime.now(TZ).isoformat(timespec="seconds")]
    ws.update(range_name=f"A{target}:D{target}", values=[value], value_input_option="RAW")

def preview_html(source):
    samples = {
        "{{nombre}}": "Cliente de prueba",
        "{{referencia}}": "REF-EJEMPLO",
        "{{fecha_pago}}": "Fecha por verificar",
        "{{valor_pendiente}}": "Valor por verificar",
    }
    for key, sample in samples.items():
        source = source.replace(key, html.escape(sample))
    return source

def render_plantillas_mora1():
    """Panel integrado en el menú Plantillas de Masivos Correos."""
    st.title("✉️ Plantillas de correos — Mora 1")
    st.info("Aquí solo se editan y guardan plantillas. Esta página NO envía correos.")
    st.caption("Marcadores disponibles: {{nombre}}, {{referencia}}, {{fecha_pago}}, {{valor_pendiente}}. "
               "Solo se deben utilizar datos verificados antes de enviar.")

    saved, error = load_saved()
    if error:
        st.warning("No se pudieron leer las plantillas guardadas. Se muestran borradores locales. "
                   "Revisa MI_JSON y el acceso del servicio a Google Sheets.")
    code = st.selectbox("Selecciona la plantilla", CODES, format_func=lambda c: f"{c} — Día {DAYS[c]}")
    st.caption(f"Campaña especial del día {DAYS[code]} de Mora 1")
    subject = st.text_input("Asunto", value=saved.get(code, {}).get("subject") or SUBJECTS[code], key=f"subject_{code}")
    body = st.text_area("Contenido HTML editable", value=saved.get(code, {}).get("body") or default_html(code),
                        height=320, key=f"body_{code}")
    st.subheader("Vista previa (datos ficticios)")
    components.html(preview_html(body), height=440, scrolling=True)
    if st.button("Guardar esta plantilla en Google Sheets", type="primary"):
        if not subject.strip() or not body.strip():
            st.error("El asunto y el contenido no pueden estar vacíos.")
        else:
            try:
                save_template(code, subject.strip(), body)
                st.success(f"{code} guardada correctamente en la hoja {SHEET}.")
            except Exception:
                st.error("No se pudo guardar. Comprueba MI_JSON, los permisos de edición y la hoja Plantillas_Mora1.")
    st.caption("No habilites envíos en producción hasta verificar destinatarios, exclusiones e historial de mensajes.")
