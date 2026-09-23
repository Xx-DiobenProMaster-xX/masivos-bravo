import os, re, json, unicodedata
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import gspread
import io
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.credentials import Credentials as UserCredentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import base64

TZ = ZoneInfo("America/Bogota")

PIPELINE_SPREADSHEET_ID = "1H3sYEtkeu47POnu8xZMaMtID1Vj53YIcWblWeZ8d0rc"
PIPELINE_SHEET = "BD 2026"
PIPELINE_SHEET_MES = "BD del mes"

MASIVOS_SPREADSHEET_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
CARTERA_SPREADSHEET_ID = "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
ESTRUCTURADOS_SPREADSHEET_ID = "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"

HOJAS_INFO_CLIENTES = ["Info_Clientes_V2", "Hoja Info_Clientes_V2", ". Hoja Info_Clientes_V2"]
HOJA_CARTERA_BEREX = "2. Cartera Berex"  # legado; no usado aquí
ASIGNACIONES_FOLDER_ID = "1cf2p3R7iM0xowAt4muEruDwxZoZqD_jB"
MESES_ES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
            7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
HOJA_EXCLUIR = "Excluir_correo"
PLANTILLA_ID = "ESTBIENV003"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def norm_txt(v):
    s = unicodedata.normalize("NFD", str(v or "").strip())
    return "".join(c for c in s if unicodedata.category(c) != "Mn").upper()

def norm_ref(v):
    s = str(v or "").strip().replace("\xa0", "").replace(" ", "")
    if not s or s.upper() in {"NAN","NONE","NULL"}: return ""
    if re.fullmatch(r"[+-]?\d+\.0+", s): return s.split(".")[0].lstrip("+")
    if re.fullmatch(r"[+-]?\d+", s): return s.lstrip("+")
    try:
        n = float(s.replace(",", ""))
        if n.is_integer(): return str(int(n))
    except Exception:
        pass
    return re.sub(r"\D", "", s)

def parse_fecha(v):
    s = str(v or "").strip()
    for f in ("%d/%m/%Y","%Y-%m-%d","%d/%m/%Y %H:%M:%S","%Y-%m-%d %H:%M:%S","%d-%m-%Y"):
        try: return datetime.strptime(s, f).date()
        except Exception: pass
    return None

def es_true(v):
    if v is True: return True
    return norm_txt(v) in {"TRUE","VERDADERO","SI","1","X","CHECKED"}

def correo_valido(v):
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(v or "").strip().lower()))

def ocultar_email(v):
    s = str(v or "").strip()
    if "@" not in s: return s
    u, d = s.split("@", 1)
    return f"{u[:2]}***@{d}"

def sheets():
    raw = os.environ.get("MI_JSON", "").strip()
    if not raw: raise RuntimeError("No encontré el Secret MI_JSON.")
    cred = Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    return gspread.authorize(cred)

def maestro_clientes(gc):
    libro = gc.open_by_key(CARTERA_SPREADSHEET_ID)
    hoja = None
    encontrada = ""
    for nombre in HOJAS_INFO_CLIENTES:
        try:
            hoja = libro.worksheet(nombre)
            encontrada = nombre
            break
        except Exception:
            pass
    if hoja is None:
        raise RuntimeError("No encontré Info_Clientes_V2.")
    datos = {}
    for f in hoja.get("C:F")[1:]:
        ref = norm_ref(f[0] if len(f)>0 else "")
        if not ref: continue
        nombre = str(f[2] if len(f)>2 else "").strip()
        email = str(f[3] if len(f)>3 else "").strip().lower()
        actual = datos.get(ref, {"NOMBRE":"","EMAIL":""})
        if nombre and not actual["NOMBRE"]: actual["NOMBRE"] = nombre
        if email and not actual["EMAIL"]: actual["EMAIL"] = email
        datos[ref] = actual
    print(f"Fuente clientes encontrada: {encontrada}")
    print(f"Referencias cargadas desde Info_Clientes_V2: {len(datos)}")
    return datos



def _credenciales_service_account():
    raw = os.environ.get("MI_JSON", "").strip()
    if not raw:
        raise RuntimeError("Falta el secret MI_JSON.")
    info = json.loads(raw)
    return Credentials.from_service_account_info(info, scopes=SCOPES)


def _descargar_archivo_drive(drive, file_id):
    request = drive.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    terminado = False
    while not terminado:
        _, terminado = downloader.next_chunk()
    buffer.seek(0)
    return buffer


def _mes_desde_nombre_hoja(nombre):
    partes = str(nombre or "").strip().split()
    if len(partes) != 2 or not partes[1].isdigit():
        return None
    mes_num = next(
        (n for n, v in MESES_ES.items() if norm_txt(v) == norm_txt(partes[0])),
        None
    )
    if mes_num is None:
        return None
    return int(partes[1]), mes_num


def maestro_asignaciones_vigentes(gc, hoy):
    """Busca en Drive el XLSX de Asignaciones más reciente y lee A/C/E con pandas."""
    creds = _credenciales_service_account()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    # Busca archivos dentro de la carpeta compartida. Soporta XLSX y Excel antiguo.
    q = (
        f"'{ASIGNACIONES_FOLDER_ID}' in parents and trashed = false "
        "and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
        "or mimeType = 'application/vnd.ms-excel')"
    )
    resp = drive.files().list(
        q=q,
        fields="files(id,name,mimeType,modifiedTime,createdTime)",
        orderBy="modifiedTime desc",
        pageSize=100,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    archivos = resp.get("files", [])

    archivos = [
        f for f in archivos
        if "ASIGNACIONES DE CARTERA" in norm_txt(f.get("name", ""))
    ]
    if not archivos:
        raise RuntimeError(
            "No encontré archivos Excel 'Asignaciones de Cartera' dentro de la carpeta. "
            "Comparte la carpeta/archivo con el correo de la service account de MI_JSON."
        )

    # Priorizamos el libro cuyo nombre cubra el mes/año actual; si no, el más recientemente modificado.
    meses_abrev = {
        1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",
        7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"
    }
    yy = str(hoy.year)[-2:]
    esperado_token = f"{meses_abrev[hoy.month]}{yy}"
    candidatos_mes = [f for f in archivos if esperado_token in norm_txt(f.get("name","")).replace(" ","")]
    archivo = candidatos_mes[0] if candidatos_mes else archivos[0]

    print(f"Archivo de asignaciones seleccionado: {archivo['name']}")
    print(f"Última modificación Drive: {archivo.get('modifiedTime','')}")

    contenido = _descargar_archivo_drive(drive, archivo["id"])
    excel = pd.ExcelFile(contenido, engine="openpyxl")
    hojas = excel.sheet_names

    esperado = f"{MESES_ES[hoy.month]} {hoy.year}"
    hoja = next((h for h in hojas if norm_txt(h) == norm_txt(esperado)), None)

    if hoja is None:
        candidatas = []
        for h in hojas:
            parsed = _mes_desde_nombre_hoja(h)
            if parsed:
                candidatas.append((parsed[0], parsed[1], h))
        if not candidatas:
            raise RuntimeError(
                f"El archivo {archivo['name']} no contiene hojas mensuales reconocibles. "
                f"Hojas encontradas: {hojas}"
            )
        candidatas.sort(key=lambda x:(x[0],x[1]), reverse=True)
        hoja = candidatas[0][2]

    print(f"Hoja mensual seleccionada: {hoja}")

    # Volvemos a descargar porque ExcelFile ya consumió el buffer en algunas versiones.
    contenido = _descargar_archivo_drive(drive, archivo["id"])
    df = pd.read_excel(
        contenido,
        sheet_name=hoja,
        usecols="A:E",
        dtype=str,
        engine="openpyxl"
    ).fillna("")

    datos = {}
    for _, fila in df.iterrows():
        ref = norm_ref(fila.iloc[0] if len(fila)>0 else "")
        if not ref:
            continue
        nombre = str(fila.iloc[2] if len(fila)>2 else "").strip()
        email = str(fila.iloc[4] if len(fila)>4 else "").strip().lower()
        actual = datos.get(ref, {"NOMBRE":"","EMAIL":""})
        if nombre and not actual["NOMBRE"]:
            actual["NOMBRE"] = nombre
        if correo_valido(email) and not correo_valido(actual["EMAIL"]):
            actual["EMAIL"] = email
        datos[ref] = actual

    print(f"Referencias cargadas desde {archivo['name']} / {hoja}: {len(datos)}")
    return datos

def exclusiones(gc):
    vals = gc.open_by_key(ESTRUCTURADOS_SPREADSHEET_ID).worksheet(HOJA_EXCLUIR).col_values(1)
    return {norm_ref(x) for x in vals[1:] if norm_ref(x)}

def ids_existentes(gc):
    vals = gc.open_by_key(MASIVOS_SPREADSHEET_ID).worksheet("COLA_ENVIO").get_all_values()
    if not vals: return set()
    cab = [str(x).strip() for x in vals[0]]
    if "ID_ENVIO" not in cab: raise RuntimeError("COLA_ENVIO no contiene ID_ENVIO.")
    i = cab.index("ID_ENVIO")
    return {str(f[i]).strip() for f in vals[1:] if len(f)>i and str(f[i]).strip()}


GMAIL_FROM = "estructurados@gobravo.com.co"
GMAIL_REPLY_TO = "estructurados@gobravo.com.co"
WHATSAPP_URL = "https://wa.me/573012411885"
LOGO_URL = "https://drive.google.com/uc?export=view&id=13kK3v4FiyXFa4UzM_au3TllhOhwjvWb7"
ASUNTO_BIENVENIDA = "¡Bienvenido al área de estructurados! | Bravo"

PLANTILLA_DESCUENTO_INCOBRABLE = "DESCINC001"
PLANTILLA_ALTERNATIVAS_PAGO = "ALTPAGO001"
ASUNTO_DESCUENTO_INCOBRABLE = "Tenemos un beneficio especial para ti | Bravo"
ASUNTO_ALTERNATIVAS_PAGO = "Tenemos alternativas de pago para ti | Bravo"
WHATSAPP_BRAVO = "https://wa.me/573012411885"


def gmail_service():
    client_id = os.environ.get("GMAIL_CLIENT_ID", "").strip()
    client_secret = os.environ.get("GMAIL_CLIENT_SECRET", "").strip()
    refresh_token = os.environ.get("GMAIL_REFRESH_TOKEN", "").strip()
    if not all([client_id, client_secret, refresh_token]):
        raise RuntimeError("Faltan GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET o GMAIL_REFRESH_TOKEN.")
    creds = UserCredentials(
        token=None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=["https://www.googleapis.com/auth/gmail.send"],
    )
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def html_bienvenida_estructurados(nombre):
    nombre = str(nombre or "").strip()
    saludo = f"Hola, {nombre}" if nombre else "Hola"
    return f"""<!doctype html>
<html>
<body style="margin:0;padding:0;background:#f4f5fb;font-family:Arial,Helvetica,sans-serif;color:#17145f;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f4f5fb;padding:24px 0;">
<tr><td align="center">
<table role="presentation" width="600" cellspacing="0" cellpadding="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:18px;overflow:hidden;">
<tr><td align="center" style="padding:28px 28px 12px;">
<img src="{LOGO_URL}" alt="Bravo" width="150" style="display:block;max-width:150px;height:auto;">
</td></tr>
<tr><td align="center" style="padding:10px 34px 8px;">
<div style="font-size:30px;font-weight:800;line-height:1.1;">¡Bienvenido al área<br>de <em>estructurados!</em></div>
<div style="margin:20px auto 0;background:#eee9ff;border-radius:24px;padding:10px 18px;font-size:13px;max-width:390px;">
Estás cada vez más cerca de liberarte de las deudas.
</div>
</td></tr>
<tr><td style="padding:20px 42px 8px;font-size:15px;line-height:1.55;">
<strong>{saludo}</strong><br><br>
Acabas de iniciar tu proceso de <strong>liquidación estructurada</strong>, un beneficio exclusivo para clientes con un excelente hábito de pago.
</td></tr>
<tr><td style="padding:20px 34px 10px;background:#8269df;text-align:center;color:white;">
<div style="font-size:24px;font-weight:800;">Para tener éxito<br><span style="color:#35f0ef;">en esta etapa recuerda:</span></div>
</td></tr>
<tr><td style="padding:8px 42px 30px;background:#8269df;">
<div style="background:#342b86;color:white;border-radius:14px;padding:18px;margin:12px 0;font-size:13px;line-height:1.45;">
<strong style="font-size:18px;color:#35f0ef;">1</strong><br>
Si tienes comisiones diferidas, firmaste un pagaré que garantiza el pago total, incluso si decides no continuar con el programa.
</div>
<div style="background:#342b86;color:white;border-radius:14px;padding:18px;margin:12px 0;font-size:13px;line-height:1.45;">
<strong style="font-size:18px;color:#35f0ef;">2</strong><br>
Es importante cumplir puntualmente con los acuerdos de pago pendientes con las entidades financieras. En caso de incumplimiento tus aportes se destinarán primero a intereses de mora, luego a gastos de cobranza, interés corriente y, por último, a capital.
</div>
<div style="background:#342b86;color:white;border-radius:14px;padding:18px;margin:12px 0;font-size:13px;line-height:1.45;">
<strong style="font-size:18px;color:#35f0ef;">3</strong><br>
Si realizas más de un pago bancario, te informaremos oportunamente cómo se aplican.
</div>
</td></tr>
<tr><td style="padding:26px 42px;text-align:center;font-size:14px;line-height:1.5;">
Si necesitas apoyo, nuestro equipo está disponible para ayudarte.<br><br>
<a href="{WHATSAPP_URL}" style="display:block;background:#15c576;color:white;text-decoration:none;font-weight:700;border-radius:12px;padding:14px 18px;">Hablar con Bravo por WhatsApp</a>
</td></tr>
<tr><td style="background:#eee9ff;text-align:center;padding:20px 30px;">
<strong style="font-size:16px;">Disciplina hoy, tranquilidad mañana.</strong><br>
<span style="font-size:12px;">Estamos contigo durante esta nueva etapa.</span>
</td></tr>
<tr><td style="background:#3b2a98;text-align:center;color:white;padding:24px 30px;font-size:12px;">
<strong>¡Gracias por confiar en nosotros!</strong><br>Equipo Bravo
</td></tr>
</table>
</td></tr></table>
</body></html>"""



def _html_base_beneficio_bravo(nombre, tipo):
    """Base visual para las plantillas DESCINC001 y ALTPAGO001."""
    from urllib.parse import quote

    nombre = str(nombre or "").strip()
    saludo = f"¡Hola, {nombre}! 👋" if nombre else "¡Hola! 👋"

    if str(tipo).strip().upper() == PLANTILLA_DESCUENTO_INCOBRABLE:
        etiqueta = "TENEMOS UN BENEFICIO PARA TI"
        intro = (
            "Queremos ayudarte a finalizar tus compromisos pendientes con Bravo. "
            "Por eso, tenemos un <strong>beneficio especial sobre tu comisión de éxito.</strong>"
        )
        titulo = "Descuento en tu comisión de éxito"
        detalle = (
            "Aprovecha este beneficio para finalizar tu comisión de éxito pendiente "
            "y quedar a paz y salvo con Bravo."
        )
        bloque_titulo = "Ponte al día con Bravo"
        bloque_texto = (
            "Aprovecha este beneficio para finalizar tu comisión de éxito pendiente "
            "y quedar a paz y salvo con Bravo."
        )
        boton = "Quiero aprovechar mi descuento"
        mensaje_wa = (
            "Hola, quiero conocer el beneficio disponible sobre mi comisión de éxito con Bravo."
        )
        cierre = "Una oportunidad para cerrar esta etapa."
        subcierre = "Estamos contigo para ayudarte a finalizar tu proceso."
    else:
        etiqueta = "TENEMOS ALTERNATIVAS PARA TI"
        intro = (
            "Queremos ayudarte a encontrar una solución para tus compromisos pendientes con Bravo."
        )
        titulo = "Tenemos alternativas de pago para ti"
        detalle = (
            "Conoce las opciones disponibles para ponerte al día "
            "o avanzar hacia tu paz y salvo con Bravo."
        )
        bloque_titulo = "Encuentra una alternativa para ti"
        bloque_texto = (
            "Nuestro equipo puede revisar contigo las opciones disponibles "
            "para ayudarte a ponerte al día o quedar a paz y salvo con Bravo."
        )
        boton = "Conocer mis alternativas de pago"
        mensaje_wa = (
            "Hola, quiero conocer las alternativas de pago que tengo disponibles con Bravo."
        )
        cierre = "Siempre hay un siguiente paso."
        subcierre = "Estamos contigo para ayudarte a encontrarlo."

    whatsapp_url = f"{WHATSAPP_BRAVO}?text={quote(mensaje_wa)}"

    return f"""<!doctype html>
<html>
<body style="margin:0;padding:0;background:#f4f3f9;font-family:Arial,Helvetica,sans-serif;color:#17145c;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f4f3f9;padding:24px 0;">
<tr><td align="center">
<table role="presentation" width="600" cellspacing="0" cellpadding="0"
       style="max-width:600px;width:100%;background:#ffffff;border-radius:18px;overflow:hidden;">

<tr>
<td align="center" style="background:#3e2c96;padding:26px 20px;">
<img src="{LOGO_URL}" alt="Bravo" width="150" style="display:block;max-width:150px;height:auto;">
</td>
</tr>

<tr>
<td align="center" style="padding:34px 40px 10px;">
<span style="display:inline-block;background:#eee9ff;color:#4934a5;font-size:12px;font-weight:bold;
             padding:9px 18px;border-radius:20px;">{etiqueta}</span>
<div style="margin-top:20px;font-size:29px;line-height:35px;font-weight:800;color:#17145c;">
{saludo}
</div>
<div style="margin-top:16px;font-size:15px;line-height:24px;color:#4d4969;">
{intro}
</div>
</td>
</tr>

<tr>
<td style="padding:22px 40px 10px;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0"
       style="background:#f1edff;border-radius:18px;">
<tr><td align="center" style="padding:30px 25px;">
<div style="font-size:13px;font-weight:bold;color:#6156a6;letter-spacing:.5px;">
QUEREMOS AYUDARTE
</div>
<div style="margin-top:11px;font-size:27px;line-height:33px;font-weight:800;color:#3e2c96;">
{titulo}
</div>
<div style="margin-top:10px;font-size:15px;line-height:23px;color:#514c72;">
{detalle}
</div>
</td></tr>
</table>
</td>
</tr>

<tr>
<td style="padding:20px 40px 4px;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0"
       style="background:#f8f7fc;border-radius:14px;">
<tr>
<td align="center" style="padding:23px 22px;">
<div style="font-size:22px;color:#3e2c96;">✓</div>
<div style="margin-top:6px;font-size:16px;font-weight:bold;color:#17145c;">
{bloque_titulo}
</div>
<div style="margin-top:7px;font-size:13px;line-height:20px;color:#625e7a;">
{bloque_texto}
</div>
</td>
</tr>
</table>
</td>
</tr>

<tr>
<td align="center" style="padding:28px 40px 34px;">
<div style="margin-bottom:17px;font-size:14px;line-height:22px;color:#5e5a78;">
Si quieres conocer las condiciones u opciones disponibles, nuestro equipo está listo para ayudarte.
</div>

<a href="{whatsapp_url}" target="_blank"
   style="display:block;background:#16c477;color:#ffffff;text-decoration:none;font-weight:700;
          border-radius:12px;padding:16px 18px;font-size:16px;">
💬 &nbsp; {boton}
</a>

<div style="margin-top:14px;font-size:12px;line-height:18px;color:#77738d;">
También puedes escribirnos por WhatsApp al <strong>301 241 1885</strong>
</div>
</td>
</tr>

<tr>
<td align="center" style="background:#eeeaff;padding:22px 30px;">
<strong style="font-size:16px;color:#241778;">{cierre}</strong><br>
<span style="font-size:12px;line-height:18px;color:#625d83;">{subcierre}</span>
</td>
</tr>

<tr>
<td align="center" style="background:#3e2c96;color:#ffffff;padding:24px 30px;font-size:12px;">
<strong>¡Gracias por confiar en nosotros!</strong><br>Equipo Bravo
</td>
</tr>

</table>
</td></tr>
</table>
</body>
</html>"""


def html_descuento_incobrable(nombre):
    """
    DESCINC001:
    beneficio sobre comisión de éxito para clientes con status INCOBRABLE.
    No muestra tipo de mora, días de mora, comisión pendiente ni valor con beneficio.
    """
    return _html_base_beneficio_bravo(nombre, PLANTILLA_DESCUENTO_INCOBRABLE)


def html_alternativas_pago(nombre):
    """
    ALTPAGO001:
    comunica alternativas de pago para ponerse al día o avanzar hacia paz y salvo.
    """
    return _html_base_beneficio_bravo(nombre, PLANTILLA_ALTERNATIVAS_PAGO)


def enviar_gmail(service, destino, asunto, html):
    msg = MIMEMultipart("alternative")
    msg["To"] = destino
    msg["From"] = GMAIL_FROM
    msg["Reply-To"] = GMAIL_REPLY_TO
    msg["Subject"] = asunto
    msg.attach(MIMEText(html, "html", "utf-8"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    return service.users().messages().send(userId="me", body={"raw": raw}).execute()


def agregar_cola(ws, id_envio, ref, nombre, email, fecha_liquidacion, estado, message_id="", error=""):
    ahora = datetime.now(TZ)
    # A:O según estructura existente de COLA_ENVIO.
    fila = [
        id_envio, "AUTO_BIENVENIDA_ESTRUCTURADOS", ref, nombre, email,
        PLANTILLA_ID, ASUNTO_BIENVENIDA, estado,
        ahora.strftime("%Y-%m-%d %H:%M:%S"),
        ahora.strftime("%Y-%m-%d %H:%M:%S") if estado == "ENVIADO" else "",
        1, error, message_id,
        f"Bienvenida estructurados | liquidación {fecha_liquidacion.strftime('%d/%m/%Y')}",
        "AUTO",
    ]
    ws.append_row(fila, value_input_option="USER_ENTERED")


def main():
    ahora = datetime.now(TZ)
    hoy = ahora.date()
    limite = hoy - timedelta(days=3)
    inicio_mes = hoy.replace(day=1)

    print("=" * 76)
    print("BIENVENIDA A ESTRUCTURADOS - PRODUCCIÓN")
    print(f"Fecha/hora Colombia: {ahora.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"Primera puesta al día: liquidaciones {inicio_mes.strftime('%d/%m/%Y')} a {limite.strftime('%d/%m/%Y')}")
    print("=" * 76)

    raw = os.environ.get("MI_JSON", "").strip()
    if not raw:
        raise RuntimeError("Falta MI_JSON.")
    sa = json.loads(raw)
    creds = Credentials.from_service_account_info(sa, scopes=SCOPES)
    gc = gspread.authorize(creds)

    libro_pipeline = gc.open_by_key(PIPELINE_SPREADSHEET_ID)
    filas_mes = libro_pipeline.worksheet(PIPELINE_SHEET_MES).get("A:P")
    filas_hist = libro_pipeline.worksheet(PIPELINE_SHEET).get("A:P")
    fuentes = [("BD del mes", f) for f in filas_mes[1:]] + [("BD 2026", f) for f in filas_hist[1:]]

    maestro = maestro_clientes(gc)
    asignaciones = maestro_asignaciones_vigentes(gc, hoy)
    excluir = exclusiones(gc)
    existentes = ids_existentes(gc)
    cola = gc.open_by_key(MASIVOS_SPREADSHEET_ID).worksheet("COLA_ENVIO")

    # Un evento por referencia + fecha. P=TRUE confirma estructurado.
    eventos = {}
    for fuente, f in fuentes:
        fecha = parse_fecha(f[2] if len(f)>2 else "")
        ref = norm_ref(f[7] if len(f)>7 else "")
        est = es_true(f[15] if len(f)>15 else "")
        if not fecha or not ref or not est:
            continue
        if not (inicio_mes <= fecha <= limite):
            continue
        eventos[(ref, fecha)] = {"REF": ref, "FECHA": fecha, "FUENTE": fuente}

    print(f"Estructurados elegibles encontrados: {len(eventos)}")

    gmail = gmail_service()
    enviados = duplicados = excluidos = sin_correo = errores = 0

    for _, e in sorted(eventos.items(), key=lambda kv:(kv[1]["FECHA"], kv[1]["REF"])):
        ref, fecha = e["REF"], e["FECHA"]
        id_envio = f"ENV-ESTBIENV-{fecha.strftime('%Y%m%d')}-{ref}-{PLANTILLA_ID}"

        if id_envio in existentes:
            duplicados += 1
            continue
        if ref in excluir:
            excluidos += 1
            continue

        a = maestro.get(ref, {})
        b = asignaciones.get(ref, {})
        nombre = str(a.get("NOMBRE","")).strip() or str(b.get("NOMBRE","")).strip()
        email_a = str(a.get("EMAIL","")).strip().lower()
        email_b = str(b.get("EMAIL","")).strip().lower()
        email = email_a if correo_valido(email_a) else email_b if correo_valido(email_b) else ""

        if not correo_valido(email):
            sin_correo += 1
            print(f"SIN_CORREO | {ref} | {fecha.strftime('%d/%m/%Y')}")
            continue

        try:
            html = html_bienvenida_estructurados(nombre)
            resp = enviar_gmail(gmail, email, ASUNTO_BIENVENIDA, html)
            mid = str(resp.get("id",""))
            agregar_cola(cola, id_envio, ref, nombre, email, fecha, "ENVIADO", mid)
            existentes.add(id_envio)
            enviados += 1
            print(f"ENVIADO | {ref} | {fecha.strftime('%d/%m/%Y')} | {ocultar_email(email)}")
        except Exception as exc:
            errores += 1
            err = str(exc)[:450]
            try:
                agregar_cola(cola, id_envio, ref, nombre, email, fecha, "ERROR", "", err)
                existentes.add(id_envio)
            except Exception:
                pass
            print(f"ERROR | {ref} | {fecha.strftime('%d/%m/%Y')} | {err}")

    print("=" * 76)
    print(f"ENVIADOS: {enviados}")
    print(f"YA REGISTRADOS / DUPLICADOS: {duplicados}")
    print(f"EXCLUIDOS: {excluidos}")
    print(f"SIN CORREO: {sin_correo}")
    print(f"ERRORES: {errores}")
    print("=" * 76)

    if errores:
        raise RuntimeError(f"Finalizó con {errores} error(es) de envío.")


if __name__ == "__main__":
    main()
