import os
import json
import re
import base64
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

import gspread
from gspread.exceptions import APIError
import random
from time import sleep
from google.oauth2.service_account import Credentials as ServiceCredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from googleapiclient.discovery import build



# ============================================================
# RETRY GLOBAL PARA GSPREAD
# Todas las llamadas de Client / Spreadsheet / Worksheet pasan
# automáticamente por _retry, incluyendo open_by_key(),
# worksheet(), get(), get_all_values(), update_cell(),
# append_row(s), delete_rows(), batch_update(), etc.
# ============================================================

_RETRIABLE_HTTP_CODES = {429, 500, 502, 503, 504}


def _retry(fn, label="", tries=10, base_sleep=1.5, jitter=0.6, max_sleep=45):
    last_err = None

    for i in range(tries):
        try:
            return fn()
        except APIError as e:
            last_err = e

            response = getattr(e, "response", None)
            status = getattr(response, "status_code", None)

            # Fallback para versiones de gspread donde el status no viene expuesto.
            msg = str(e)
            retriable = (
                status in _RETRIABLE_HTTP_CODES
                or any(f"[{code}]" in msg for code in _RETRIABLE_HTTP_CODES)
            )

            if not retriable:
                raise

            sleep_s = min(
                base_sleep * (2 ** i) + random.uniform(0, jitter),
                max_sleep
            )

            print(
                f"[GSPREAD RETRY {i + 1}/{tries}] "
                f"{label or 'request'} -> "
                f"{'HTTP ' + str(status) if status else msg[:100]} | "
                f"sleep {sleep_s:.1f}s"
            )
            sleep(sleep_s)

    raise last_err


def _wrap_gspread_result(value):
    """
    Envuelve únicamente objetos operativos de gspread.

    IMPORTANTE: gspread devuelve algunos resultados de lectura (por ejemplo
    ValueRange) como subclases de list cuyo módulo también empieza por
    ``gspread``. Esos resultados SON datos y deben seguir siendo indexables,
    iterables y compatibles con slices como rows[1:].
    """
    if isinstance(value, _GSpreadRetryProxy):
        return value

    # Nunca envolver datos/resultados. ValueRange hereda de list.
    if isinstance(value, (list, tuple, dict, set, str, bytes, int, float, bool, type(None))):
        return value

    module = getattr(value.__class__, "__module__", "")
    class_name = value.__class__.__name__

    # Solo los objetos sobre los que luego hacemos nuevas requests HTTP.
    if module.startswith("gspread") and class_name in {
        "Client",
        "Spreadsheet",
        "Worksheet",
    }:
        return _GSpreadRetryProxy(value)

    return value


class _GSpreadRetryProxy:
    """
    Proxy transparente: cualquier método de gspread ejecutado sobre Client,
    Spreadsheet o Worksheet usa _retry. Los objetos gspread devueltos por
    una llamada también quedan envueltos automáticamente.
    """

    __slots__ = ("_obj",)

    def __init__(self, obj):
        object.__setattr__(self, "_obj", obj)

    def __getattr__(self, name):
        attr = getattr(self._obj, name)

        if not callable(attr):
            return _wrap_gspread_result(attr)

        def wrapped(*args, **kwargs):
            obj_name = self._obj.__class__.__name__
            result = _retry(
                lambda: attr(*args, **kwargs),
                label=f"{obj_name}.{name}"
            )
            return _wrap_gspread_result(result)

        return wrapped

    def __setattr__(self, name, value):
        setattr(self._obj, name, value)

    def __repr__(self):
        return repr(self._obj)


def _gspread_client_with_retry(credentials):
    # authorize() crea el cliente localmente. Desde aquí, todas sus
    # operaciones HTTP quedan protegidas por el proxy.
    return _GSpreadRetryProxy(gspread.authorize(credentials))


TZ = ZoneInfo("America/Bogota")
MASIVOS_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
CARTERA_BEREX_ID = "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
EXCLUSIONES_ID = "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"

GMAIL_FROM = "acuerdosRTD@resuelvetudeuda.com"
GMAIL_SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/drive.file",
]


def txt(v):
    return "" if v is None else str(v).strip()


def ref(v):
    v = txt(v)
    if re.fullmatch(r"\d+\.0+", v):
        v = v.split(".")[0]
    return re.sub(r"\D", "", v)


def headers(row):
    return {txt(v): i for i, v in enumerate(row)}


def getv(row, h, name):
    i = h.get(name)
    return txt(row[i]) if i is not None and len(row) > i else ""


def fecha(v):
    v = txt(v)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(v[:10], fmt).date()
        except ValueError:
            pass
    return None


def es_mora_180(v):
    v = re.sub(r"\s+", " ", txt(v).upper().replace("_", " ")).strip()
    return v in {"MORA 180", "180"}


def moneda(v):
    s = txt(v).replace("$", "").replace("COP", "").replace(" ", "")
    if "." in s and "," not in s:
        p = s.split(".")
        if len(p) > 1 and all(len(x) == 3 for x in p[1:]):
            s = "".join(p)
    elif "," in s and "." not in s:
        s = s.replace(",", "")
    elif "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        n = float(s)
    except Exception:
        n = 0
    return "$" + f"{n:,.0f}".replace(",", ".")


BRAVO_LOGO_URL = "https://drive.google.com/uc?export=view&id=13kK3v4FiyXFa4UzM_au3TllhOhwjvWb7"
BRAVO_WHATSAPP = "573012411885"
BRAVO_WHATSAPP_DISPLAY = "301 241 1885"

def fecha_larga_es(v):
    d=fecha(v)
    if not d: return txt(v)
    meses=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
    return f"{d.day} de {meses[d.month-1]} de {d.year}"

def html_pab(nombre, referencia, fecha_pab, valor_pab, days, cuerpo_base=""):
    import html as _html
    nombre = _html.escape(str(nombre or "Cliente").strip() or "Cliente")
    fecha_txt = _html.escape(fecha_larga_es(fecha_pab))
    valor_txt = _html.escape(moneda(valor_pab))
    if int(days) == 0:
        titulo_1, titulo_2, badge = "Tu pago a banco", "es hoy", "Pago programado para hoy"
        intro = "Te recordamos que hoy tienes un pago a banco programado. Te compartimos los detalles de tu pago:"
    else:
        titulo_1, titulo_2, badge = "Tu próximo pago", "está cerca", "Faltan 3 días"
        intro = "Queremos recordarte que tienes un pago a banco programado para los próximos días. Te compartimos los detalles de tu próximo pago:"
    return f'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f3f5f9;font-family:Arial,Helvetica,sans-serif;color:#27304f;"><table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;background:#f3f5f9;"><tr><td align="center" style="padding:24px 10px;"><table role="presentation" width="600" cellspacing="0" cellpadding="0" border="0" style="width:600px;max-width:600px;background:#fff;border-radius:14px;overflow:hidden;">
<tr><td style="padding:28px 38px 12px;"><img src="{BRAVO_LOGO_URL}" width="190" alt="Bravo" style="display:block;width:190px;max-width:55%;height:auto;border:0;"></td></tr>
<tr><td style="padding:8px 38px 0;font-size:42px;line-height:43px;font-weight:800;color:#261269;">{titulo_1}<br><span style="color:#19b9c7;">{titulo_2}</span></td></tr>
<tr><td style="padding:26px 38px 10px;font-size:18px;line-height:27px;">Hola, <b>{nombre}:</b></td></tr><tr><td style="padding:0 38px 22px;font-size:16px;line-height:25px;">{intro}</td></tr>
<tr><td style="padding:0 38px 18px;"><table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#f5f9ff;border:1px solid #dce7f3;border-radius:12px;"><tr><td width="74" align="center" style="padding:20px 0 20px 20px;border-bottom:1px solid #dce7f3;"><div style="width:50px;height:50px;line-height:50px;border-radius:50%;background:#d8f5fb;font-size:25px;">📅</div></td><td style="padding:20px 22px;border-bottom:1px solid #dce7f3;"><div style="font-size:14px;">Fecha de pago</div><div style="font-size:25px;line-height:32px;font-weight:800;color:#111a43;">{fecha_txt}</div><span style="display:inline-block;margin-top:5px;background:#c9f3ff;color:#087eaa;padding:5px 13px;border-radius:8px;font-size:13px;font-weight:700;">{badge}</span></td></tr><tr><td width="74" align="center" style="padding:20px 0 20px 20px;"><div style="width:50px;height:50px;line-height:50px;border-radius:50%;background:#d8f5fb;font-size:25px;">$</div></td><td style="padding:20px 22px;"><div style="font-size:14px;">Valor del pago</div><div style="font-size:29px;line-height:36px;font-weight:800;color:#111a43;">{valor_txt}</div></td></tr></table></td></tr>
<tr><td style="padding:0 38px 18px;"><div style="background:#f1f7fc;border-radius:10px;padding:15px 17px;font-size:15px;line-height:22px;color:#59657c;">ⓘ &nbsp; Te recomendamos tener presente esta fecha para continuar con normalidad tu proceso.</div></td></tr>
<tr><td align="center" style="padding:2px 38px 25px;"><a href="https://wa.me/{BRAVO_WHATSAPP}" style="display:block;background:#19b9c7;color:#fff;text-decoration:none;font-size:17px;font-weight:700;padding:15px 20px;border-radius:28px;">WhatsApp &nbsp; Consultar por WhatsApp</a></td></tr>
<tr><td style="padding:18px 28px;background:#f4f8fc;"><table role="presentation" width="100%"><tr><td align="center" width="33%" style="font-size:11px;line-height:16px;color:#59657c;">◇<br>Continuamos<br>con tu proceso</td><td align="center" width="34%" style="font-size:11px;line-height:16px;color:#59657c;">♙<br>Estamos aquí<br>para apoyarte</td><td align="center" width="33%" style="font-size:11px;line-height:16px;color:#59657c;">✉<br>Escríbenos si<br>tienes dudas</td></tr></table></td></tr>
<tr><td style="padding:18px 28px 20px;border-bottom:5px solid #20bfd0;font-size:11px;line-height:17px;color:#59657c;"><table role="presentation" width="100%"><tr><td><b style="color:#261269;font-size:13px;">Bravo S.A.S.</b><br>Tu tranquilidad, nuestra prioridad.</td><td align="right">WhatsApp: {BRAVO_WHATSAPP_DISPLAY}<br>Lunes a viernes, 8:00 a.m. - 6:00 p.m.</td></tr></table></td></tr></table></td></tr></table></body></html>'''

def gc():
    raw = os.environ.get("MI_JSON", "").strip()
    if not raw:
        raise RuntimeError("Falta GitHub Secret MI_JSON.")
    info = json.loads(raw)
    creds = ServiceCredentials.from_service_account_info(
        info,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return _gspread_client_with_retry(creds)


def gmail_service():
    client_id = os.environ.get("GMAIL_CLIENT_ID", "").strip()
    client_secret = os.environ.get("GMAIL_CLIENT_SECRET", "").strip()
    refresh_token = os.environ.get("GMAIL_REFRESH_TOKEN", "").strip()
    if not all([client_id, client_secret, refresh_token]):
        raise RuntimeError(
            "Faltan GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET o GMAIL_REFRESH_TOKEN."
        )
    creds = OAuthCredentials(
        token=None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=GMAIL_SCOPES,
    )
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def send_gmail(service, to, subject, html):
    msg = EmailMessage()
    msg["To"] = to
    msg["From"] = f"Bravo S.A.S. <{GMAIL_FROM}>"
    msg["Reply-To"] = GMAIL_FROM
    msg["Subject"] = subject
    plain = re.sub(r"<[^>]+>", " ", html)
    plain = re.sub(r"\s+", " ", plain).strip() or "Bravo S.A.S."
    msg.set_content(plain)
    msg.add_alternative(html, subtype="html")
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    return service.users().messages().send(userId="me", body={"raw": raw}).execute()


def row_by_headers(hlist, values):
    return [values.get(name, "") for name in hlist]


def update_cell_by_header(ws, row_num, hlist, name, value):
    if name in hlist:
        ws.update_cell(row_num, hlist.index(name) + 1, value)


def load_client_map(g):
    result = {}

    # CLIENTES first
    ws = g.open_by_key(MASIVOS_ID).worksheet("CLIENTES")
    data = ws.get_all_values()
    if len(data) > 1:
        h = headers(data[0])
        for r in data[1:]:
            rr = ref(getv(r, h, "Referencia"))
            if rr:
                result[rr] = {
                    "nombre": getv(r, h, "Nombre"),
                    "email": getv(r, h, "Email"),
                    "mora": getv(r, h, "Mora"),
                }

    # Info_Clientes_V2 fills blanks
    book = g.open_by_key(CARTERA_BEREX_ID)
    ws2 = None
    for name in ("Info_Clientes_V2", "Hoja Info_Clientes_V2", ". Hoja Info_Clientes_V2"):
        try:
            ws2 = book.worksheet(name)
            break
        except gspread.WorksheetNotFound:
            pass
    if ws2:
        rows = ws2.get("C:F")
        for r in rows[1:]:
            rr = ref(r[0] if len(r) > 0 else "")
            if not rr:
                continue
            nombre = txt(r[2] if len(r) > 2 else "")
            email = txt(r[3] if len(r) > 3 else "")
            result.setdefault(rr, {"nombre": "", "email": "", "mora": ""})
            if not result[rr]["nombre"] and nombre:
                result[rr]["nombre"] = nombre
            if not result[rr]["email"] and email:
                result[rr]["email"] = email
    return result


def load_exclusions(g):
    ws = g.open_by_key(EXCLUSIONES_ID).worksheet("Excluir_correo")
    rows = ws.get_all_values()
    return {ref(r[0]) for r in rows[1:] if r and ref(r[0])}


def load_templates(g):
    ws = g.open_by_key(MASIVOS_ID).worksheet("PLANTILLAS")
    rows = ws.get_all_values()
    if len(rows) <= 1:
        raise RuntimeError("PLANTILLAS está vacía.")
    h = headers(rows[0])
    out = {}
    for r in rows[1:]:
        pid = (getv(r, h, "ID_PLANTILLA") or getv(r, h, "PLANTILLA")).upper()
        if not pid:
            continue
        estado = getv(r, h, "ESTADO").upper() or "ACTIVA"
        if estado == "ACTIVA":
            out[pid] = {
                "asunto": getv(r, h, "ASUNTO"),
                "cuerpo": getv(r, h, "CUERPO"),
            }
    return out


def render(template, values):
    s = txt(template)
    for k, v in values.items():
        s = s.replace(k, str(v))
    return s


def ensure_campaign(book, campaign_id, now):
    ws = book.worksheet("CAMPAÑAS")
    rows = ws.get_all_values()
    hlist = [txt(x) for x in rows[0]]
    h = headers(hlist)
    for n, r in enumerate(rows[1:], start=2):
        if getv(r, h, "ID_CAMPAÑA") == campaign_id:
            return n, ws, hlist
    values = {
        "ID_CAMPAÑA": campaign_id,
        "NOMBRE_CAMPAÑA": f"Recordatorios PaB {now.strftime('%d/%m/%Y')}",
        "PLANTILLA": "PAB",
        "FILTRO": "PAB",
        "FECHA_ENVIO": now.strftime("%d/%m/%Y"),
        "HORA_ENVIO": "08:00",
        "ESTADO": "EN PROCESO",
        "TOTAL_CLIENTES": 0,
        "ENVIADOS": 0,
        "PENDIENTES": 0,
        "ERRORES": 0,
        "FECHA_CREACIÓN": now.strftime("%d/%m/%Y %H:%M:%S"),
        "COMENTARIOS": "Campaña automática GitHub Actions 08:00 America/Bogota",
    }
    ws.append_row(row_by_headers(hlist, values), value_input_option="USER_ENTERED")
    return len(rows) + 1, ws, hlist


def recalc_campaign(book, campaign_id):
    q = book.worksheet("COLA_ENVIO")
    qr = q.get_all_values()
    qh = headers(qr[0])
    states = []
    for r in qr[1:]:
        if getv(r, qh, "ID_CAMPAÑA") == campaign_id:
            states.append(getv(r, qh, "ESTADO").upper())
    total = len(states)
    sent = sum(x == "ENVIADO" for x in states)
    errors = sum(x in {"ERROR", "BLOQUEADO"} for x in states)
    pending = sum(x in {"BORRADOR", "PENDIENTE", "ENVIANDO"} for x in states)

    c = book.worksheet("CAMPAÑAS")
    cr = c.get_all_values()
    chlist = [txt(x) for x in cr[0]]
    ch = headers(chlist)
    for n, r in enumerate(cr[1:], start=2):
        if getv(r, ch, "ID_CAMPAÑA") == campaign_id:
            update_cell_by_header(c, n, chlist, "TOTAL_CLIENTES", total)
            update_cell_by_header(c, n, chlist, "ENVIADOS", sent)
            update_cell_by_header(c, n, chlist, "PENDIENTES", pending)
            update_cell_by_header(c, n, chlist, "ERRORES", errors)
            state = "FINALIZADA" if pending == 0 and errors == 0 else (
                "FINALIZADA CON ERRORES" if pending == 0 else "EN PROCESO"
            )
            update_cell_by_header(c, n, chlist, "ESTADO", state)
            break


def main():
    now = datetime.now(TZ)
    today = now.date()
    g = gc()
    gmail = gmail_service()
    book = g.open_by_key(MASIVOS_ID)

    pab_ws = book.worksheet("PAB_PROXIMOS")
    pab_rows = pab_ws.get_all_values()
    if len(pab_rows) <= 1:
        print("Sin PAB_PROXIMOS.")
        return
    phlist = [txt(x) for x in pab_rows[0]]
    ph = headers(phlist)

    queue_ws = book.worksheet("COLA_ENVIO")
    queue_rows = queue_ws.get_all_values()
    qhlist = [txt(x) for x in queue_rows[0]]
    qh = headers(qhlist)

    clients = load_client_map(g)
    exclusions = load_exclusions(g)
    templates = load_templates(g)
    campaign_id = f"PAB-{today.strftime('%Y%m%d')}"
    ensure_campaign(book, campaign_id, now)

    # Existing queue by deterministic ID
    existing = {}
    for n, r in enumerate(queue_rows[1:], start=2):
        eid = getv(r, qh, "ID_ENVIO")
        if eid:
            existing[eid] = (n, r)

    sent = errors = skipped = 0

    for pab_row_num, r in enumerate(pab_rows[1:], start=2):
        payment_date = fecha(getv(r, ph, "FECHA_PAB"))
        if payment_date is None:
            continue
        days = (payment_date - today).days
        if days not in {0, 3}:
            continue

        pid = "PAB000" if days == 0 else "PAB003"
        flag = "AVISO_HOY" if days == 0 else "AVISO_3_DIAS"
        if getv(r, ph, flag):
            skipped += 1
            continue

        rr = ref(getv(r, ph, "REFERENCIA"))
        if not rr or rr in exclusions:
            skipped += 1
            continue

        client = clients.get(rr, {})
        nombre = getv(r, ph, "NOMBRE") or client.get("nombre", "")
        email = getv(r, ph, "EMAIL") or client.get("email", "")
        mora = getv(r, ph, "MORA") or client.get("mora", "")
        if not email or es_mora_180(mora):
            skipped += 1
            continue

        tpl = templates.get(pid)
        if not tpl:
            print(f"ERROR {rr}: falta plantilla {pid}")
            errors += 1
            continue

        values = {
            "{{NOMBRE}}": nombre or "Cliente",
            "{{REFERENCIA}}": rr,
            "{{FECHA_PAB}}": getv(r, ph, "FECHA_PAB"),
            "{{VALOR_PAB}}": moneda(getv(r, ph, "VALOR_PAB")),
            "{{TIPO_AVISO}}": "Pago programado para hoy" if days == 0 else "Recordatorio 3 días antes",
        }
        subject = render(tpl["asunto"], values)
        cuerpo_base = render(tpl["cuerpo"], values)
        body = html_pab(nombre, rr, getv(r, ph, "FECHA_PAB"), getv(r, ph, "VALOR_PAB"), days, cuerpo_base)
        eid = f"ENV-PAB-{today.strftime('%Y%m%d')}-{rr}-{pid}"

        if eid in existing:
            row_num, qr = existing[eid]
            state = getv(qr, qh, "ESTADO").upper()
            if state == "ENVIADO":
                update_cell_by_header(pab_ws, pab_row_num, phlist, flag, "ENVIADO")
                skipped += 1
                continue
            if state not in {"BORRADOR", "PENDIENTE"}:
                print(f"OMITIDO {eid}: estado {state}")
                skipped += 1
                continue
        else:
            data = {
                "ID_ENVIO": eid,
                "ID_CAMPAÑA": campaign_id,
                "REFERENCIA": rr,
                "NOMBRE": nombre,
                "EMAIL": email,
                "PLANTILLA": pid,
                "ASUNTO": subject,
                "ESTADO": "BORRADOR",
                "FECHA_PROG": now.strftime("%d/%m/%Y %H:%M"),
                "FECHA_ENVIO": "",
                "INTENTOS": 0,
                "ERROR": "",
                "ID_MENSAJE": "",
                "CUERPO": body,
                "ENCARGADO": getv(r, ph, "ENCARGADO") or "Camila",
            }
            queue_ws.append_row(row_by_headers(qhlist, data), value_input_option="USER_ENTERED")
            row_num = len(queue_rows) + 1
            queue_rows.append(row_by_headers(qhlist, data))
            existing[eid] = (row_num, queue_rows[-1])

        # Reserve before Gmail call
        update_cell_by_header(queue_ws, row_num, qhlist, "ESTADO", "ENVIANDO")
        try:
            current_attempts = 0
            if eid in existing:
                try:
                    current_attempts = int(float(getv(existing[eid][1], qh, "INTENTOS") or 0))
                except Exception:
                    current_attempts = 0
            update_cell_by_header(queue_ws, row_num, qhlist, "INTENTOS", current_attempts + 1)

            result = send_gmail(gmail, email, subject, body)
            gmail_id = txt(result.get("id"))
            update_cell_by_header(queue_ws, row_num, qhlist, "FECHA_ENVIO", datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S"))
            update_cell_by_header(queue_ws, row_num, qhlist, "ID_MENSAJE", gmail_id)
            update_cell_by_header(queue_ws, row_num, qhlist, "ERROR", "")
            update_cell_by_header(queue_ws, row_num, qhlist, "ESTADO", "ENVIADO")
            update_cell_by_header(pab_ws, pab_row_num, phlist, flag, "ENVIADO")
            sent += 1
            print(f"ENVIADO {pid} {rr} -> {email}")
        except Exception as e:
            update_cell_by_header(queue_ws, row_num, qhlist, "ERROR", str(e)[:500])
            update_cell_by_header(queue_ws, row_num, qhlist, "ESTADO", "ERROR")
            errors += 1
            print(f"ERROR {pid} {rr}: {e}")

    recalc_campaign(book, campaign_id)
    print(json.dumps({
        "fecha_bogota": today.isoformat(),
        "enviados": sent,
        "errores": errors,
        "omitidos": skipped,
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
