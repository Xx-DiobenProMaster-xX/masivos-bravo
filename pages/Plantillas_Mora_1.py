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
    "MORA101": "Te contactamos porque, según nuestro registro, hay una mensualidad pendiente. Si ya realizaste tu pago, puedes omitir este mensaje e informarnos para actualizar tu gestión.",
    "MORA102": "Queremos dar seguimiento a la mensualidad que aparece pendiente en nuestro registro. Si presentaste alguna novedad o ya pagaste, por favor cuéntanos para poder orientarte.",
    "MORA103": "Nos gustaría conocer si has tenido algún inconveniente con tu mensualidad pendiente. Estamos disponibles para revisar tu caso y orientarte con los pasos a seguir.",
    "MORA104": "Continuamos con el seguimiento de tu mensualidad. Si ya pagaste o acordaste una fecha con nuestro equipo, responde a este correo para revisar la información.",
    "MORA105": "Nos interesa revisar contigo el estado de la mensualidad pendiente. Por favor comunícate con nuestro equipo para resolver cualquier novedad o confirmar tu pago.",
}

def default_html(code):
    return (
        '<div style="font-family:Arial,sans-serif;max-width:620px;margin:auto;color:#263238;'
        'line-height:1.6;padding:26px;border:1px solid #e3e8ec;border-radius:12px">'
        '<h2 style="color:#173c64;margin-top:0">Hola, {{nombre}}</h2>'
        '<p>' + MESSAGES[code] + '</p>'
        '<p><strong>Referencia:</strong> {{referencia}}</p>'
        '<p>Si necesitas ayuda, escríbenos a través de nuestros canales habituales.</p>'
        '<p style="margin-top:30px">Cordialmente,<br><strong>Equipo de atención</strong></p>'
        '<hr style="border:0;border-top:1px solid #eee">'
        '<p style="font-size:12px;color:#69757d">Mensaje informativo. '
        'Si ya realizaste el pago, agradecemos tu confirmación.</p></div>'
    )

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

st.set_page_config(page_title="Plantillas Mora 1", page_icon="✉️", layout="wide")
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
