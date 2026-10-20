"""Editor visual para las cinco plantillas de Mora 1. Nunca envia correos."""
import html
from mora1_html import html_mora1
import json
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import streamlit as st
import streamlit.components.v1 as components
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
SHEET = "PLANTILLAS"
REQUIRED = {"ID_PLANTILLA", "ASUNTO", "CUERPO", "ESTADO"}
TZ = ZoneInfo("America/Bogota")
CODES = ["T001", "MORA101", "MORA102", "MORA103", "MORA104", "MORA105"]
DAYS = {"T001": "pares", "MORA101": 4, "MORA102": 10, "MORA103": 15, "MORA104": 19, "MORA105": 24}

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

BRAVO_LOGO_URL = "https://drive.google.com/uc?export=view&id=13kK3v4FiyXFa4UzM_au3TllhOhwjvWb7"

def default_html(code):
    """Plantilla BRAVO morado/turquesa; conserva mensaje y cierre."""
    if code == "T001":
        return html_mora1("""Estimado(a) {{NOMBRE}}:

Nos permitimos informarle que su obligación actualmente presenta una mora. Queremos acompañarle para regularizar su situación.

**Línea de atención:** WhatsApp 301 241 1885.

Atentamente, **Bravo S.A.S.**""")
    subject = html.escape(SUBJECTS[code])
    message = html.escape(MESSAGES[code])
    closing = html.escape(CLOSINGS[code])
    day = DAYS[code]
    return f'''<!doctype html><html lang="es"><head><meta charset="utf-8"></head>
<body style="margin:0;background:#f4f5f9;font-family:Arial,sans-serif;color:#525b82">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr><td align="center" style="padding:20px">
<table role="presentation" width="660" cellspacing="0" cellpadding="0" style="max-width:660px;width:100%;background:#fff;border-top:6px solid #38278f">
<tr><td style="padding:28px 38px"><img src="{BRAVO_LOGO_URL}" alt="Bravo" width="150" style="display:block;border:0;color:#38278f;font-size:24px;font-weight:bold">
<p style="font-size:12px;color:#737e9c">Soluciones financieras para un mejor futuro</p></td></tr>
<tr><td style="padding:8px 38px"><small style="color:#747f9e">SEGUIMIENTO DE MENSUALIDAD · DÍA {day}</small>
<h1 style="color:#35238f;font-size:34px;margin:18px 0 25px">Tu acuerdo<br><span style="color:#08b9bd">necesita atención</span></h1>
<p style="font-size:17px">Hola <strong style="color:#35238f">{{{{NOMBRE}}}}</strong>,</p>
<p style="line-height:1.7;font-size:16px">{message}</p></td></tr>
<tr><td style="padding:12px 38px"><div style="background:#f3f0fc;padding:20px;border-radius:12px"><small>REFERENCIA DEL CLIENTE</small><div style="color:#35238f;font-size:20px;font-weight:bold">{{{{REFERENCIA}}}}</div></div></td></tr>
<tr><td style="padding:12px 38px"><div style="background:#e9fbfc;padding:20px;border-radius:12px"><strong style="color:#08aeb5;font-size:21px">Aún estás a tiempo</strong><p style="line-height:1.7">{closing}</p></div></td></tr>
<tr><td style="padding:16px 38px 30px"><p>Gracias por mantenernos informados.</p><p>Cordialmente,<br><strong style="color:#35238f">Equipo BRAVO</strong></p><p style="font-size:12px;color:#8087a0">Si ya realizaste el pago, es posible que todavía no se refleje en nuestros registros. Este mensaje no constituye confirmación de saldo.</p></td></tr>
<tr><td style="background:#f6f5fb;border-bottom:5px solid #10bcc6;padding:18px 38px;font-size:12px"><strong>Bravo S.A.S.</strong> · Comunicación de seguimiento</td></tr>
</table></td></tr></table></body></html>'''

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

def _plantillas_sheet():
    book = open_book()
    # Nunca crear otra pestaña: el modulo usa exclusivamente PLANTILLAS.
    ws = book.worksheet(SHEET)
    data = ws.get_all_values()
    if not data:
        raise ValueError("PLANTILLAS no tiene encabezados.")
    headers = [str(x).strip().upper() for x in data[0]]
    if len(set(headers)) != len(headers):
        raise ValueError("PLANTILLAS contiene encabezados duplicados.")
    missing = REQUIRED - set(headers)
    if missing:
        raise ValueError("Faltan columnas en PLANTILLAS: " + ", ".join(sorted(missing)))
    return ws, data, headers

def load_saved():
    try:
        _, data, headers = _plantillas_sheet()
        h = {name: i for i, name in enumerate(headers)}
        saved = {}
        for row in data[1:]:
            def get(name):
                i = h[name]
                return row[i] if i < len(row) else ""
            code = get("ID_PLANTILLA").strip().upper()
            if code in CODES:
                saved[code] = {"subject": get("ASUNTO"), "body": get("CUERPO")}
        return saved, None
    except Exception as exc:
        return {}, f"{type(exc).__name__}: {exc}"

def save_template(code, subject, body):
    if code not in CODES:
        raise ValueError("Código de Mora 1 no permitido.")
    ws, data, headers = _plantillas_sheet()
    index = {col: i for i, col in enumerate(headers)}
    matches = [
        i for i, row in enumerate(data[1:], start=2)
        if len(row) > index["ID_PLANTILLA"]
        and row[index["ID_PLANTILLA"]].strip().upper() == code
    ]
    if len(matches) > 1:
        raise ValueError(f"Hay varias filas con {code}; revisa duplicados antes de guardar.")
    target = matches[0] if matches else len(data) + 1
    current = list(data[target-1]) if matches else []
    current += [""] * (len(headers) - len(current))
    current[index["ID_PLANTILLA"]] = code
    current[index["ASUNTO"]] = subject
    current[index["CUERPO"]] = body
    current[index["ESTADO"]] = "ACTIVA"
    for field, label in (
        ("NOMBRE_PLANTILLA", f"Mora 1 · {code} · Día {DAYS[code]}"),
        ("NOMBRE", f"Mora 1 · {code} · Día {DAYS[code]}"),
    ):
        if field in index and not current[index[field]]:
            current[index[field]] = label
    # Solo modifica la fila de esta plantilla; preserva otras columnas y filas.
    from gspread.utils import rowcol_to_a1
    last_cell = rowcol_to_a1(target, len(headers))
    ws.update(range_name=f"A{target}:{last_cell}", values=[current], value_input_option="RAW")
    readback = ws.row_values(target)
    if len(readback) <= index["ID_PLANTILLA"] or readback[index["ID_PLANTILLA"]].strip().upper() != code:
        raise RuntimeError("No se pudo verificar el guardado en PLANTILLAS.")
    return target

def preview_html(source):
    samples = {
        "{{NOMBRE}}": "Cliente de prueba",
        "{{REFERENCIA}}": "REF-EJEMPLO",
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
    st.caption("Variables compatibles con Masivos Correos: {{NOMBRE}} y {{REFERENCIA}}.")

    saved, error = load_saved()
    if error:
        st.error(f"No fue posible leer PLANTILLAS: {error}")
    code = st.selectbox("Selecciona la plantilla", CODES, format_func=lambda c: f"{c} — Día {DAYS[c]}")
    st.caption("Recordatorio habitual en días pares de Mora 1" if code=="T001" else f"Campaña especial del día {DAYS[code]} de Mora 1")
    subject = st.text_input("Asunto", value=saved.get(code, {}).get("subject") or SUBJECTS.get(code, "Apartado Mensual Bravo"), key=f"subject_{code}")
    def aplicar_diseno_bravo():
        st.session_state[f"body_{code}"] = html_mora1(saved.get(code, {}).get("body") or default_html(code)) if code=="T001" else default_html(code)
    st.button("Aplicar diseño corporativo BRAVO a esta plantilla",
              key=f"diseno_bravo_{code}", on_click=aplicar_diseno_bravo)
    body = st.text_area("Contenido HTML editable", value=saved.get(code, {}).get("body") or default_html(code),
                        height=320, key=f"body_{code}")
    st.subheader("Vista previa (datos ficticios)")
    components.html(preview_html(html_mora1(body)), height=440, scrolling=True)
    if st.button("Guardar esta plantilla en Google Sheets", type="primary"):
        if not subject.strip() or not body.strip():
            st.error("El asunto y el contenido no pueden estar vacíos.")
        else:
            try:
                row_number = save_template(code, subject.strip(), html_mora1(body))
                st.success(f"{code} guardada y verificada en PLANTILLAS, fila {row_number}.")
                st.cache_data.clear()
            except Exception as exc:
                st.error(f"No se pudo guardar: {type(exc).__name__}: {exc}")
    st.caption("No habilites envíos en producción hasta verificar destinatarios, exclusiones e historial de mensajes.")
