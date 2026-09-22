import os
import re
import json
import base64
from email.message import EmailMessage
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ============================================================
# MODO DIAGNÓSTICO - NO ENVÍA NI MODIFICA NADA
# ============================================================

TZ = ZoneInfo("America/Bogota")

MASIVOS_SPREADSHEET_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
UNIDOS_EST_SPREADSHEET_ID = "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"
HOJA_UNIDOS_EST = "Unidos_Est"
HOJA_EXCLUIR = "Excluir_correo"

CARTERA_SPREADSHEET_ID = "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
HOJA_CARTERA = "2. Cartera Berex"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

GMAIL_FROM = "estructurados@gobravo.com.co"
GMAIL_REPLY_TO = "estructurados@gobravo.com.co"
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


def normalizar(texto):
    import unicodedata
    t = unicodedata.normalize("NFD", str(texto or "").strip())
    return "".join(c for c in t if unicodedata.category(c) != "Mn").upper()


def normalizar_referencia(valor):
    if valor is None:
        return ""
    texto = str(valor).strip().replace("\xa0", "").replace(" ", "")
    if not texto or texto.upper() in {"NAN", "NONE", "NULL"}:
        return ""
    if re.fullmatch(r"[+-]?\d+\.0+", texto):
        return texto.split(".")[0].lstrip("+")
    if re.fullmatch(r"[+-]?\d{1,3}([.,]\d{3})+", texto):
        return re.sub(r"[.,]", "", texto).lstrip("+")
    if re.fullmatch(r"[+-]?\d+", texto):
        return texto.lstrip("+")
    try:
        numero = float(texto.replace(",", ""))
        if numero.is_integer():
            return str(int(numero))
    except Exception:
        pass
    return re.sub(r"\D", "", texto)


def parsear_fecha(valor):
    texto = str(valor or "").strip()
    if not texto:
        return None
    formatos = (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y",
    )
    for formato in formatos:
        try:
            return datetime.strptime(texto, formato).date()
        except Exception:
            pass
    return None


def correo_valido(valor):
    correo = str(valor or "").strip().lower()
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", correo))


def cliente_sheets():
    if not os.environ.get("MI_JSON"):
        raise RuntimeError("No encontré el Secret MI_JSON.")
    info = json.loads(os.environ["MI_JSON"])
    credenciales = ServiceAccountCredentials.from_service_account_info(
        info,
        scopes=SCOPES,
    )
    return gspread.authorize(credenciales)



def cliente_gmail():
    faltan = [
        k for k in ("GOOGLE_CLIENT_ID", "GOOGLE_CLIENT_SECRET", "GOOGLE_REFRESH_TOKEN")
        if not os.environ.get(k)
    ]
    if faltan:
        raise RuntimeError("Faltan variables Gmail: " + ", ".join(faltan))
    cred = OAuthCredentials(
        token=None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        scopes=GMAIL_SCOPES,
    )
    cred.refresh(Request())
    return build("gmail", "v1", credentials=cred, cache_discovery=False)


def fecha_larga(d):
    meses = ["enero","febrero","marzo","abril","mayo","junio",
             "julio","agosto","septiembre","octubre","noviembre","diciembre"]
    return f"{d.day} de {meses[d.month-1]} de {d.year}"


def html_al_dia(nombre, d, dias):
    import html as html_lib
    nombre = html_lib.escape(str(nombre or "Cliente"))
    ft = fecha_larga(d)
    if dias == 0:
        intro = "Te recordamos que hoy es la fecha de tu apartado mensual en Bravo."
        destacado = "Hoy es la fecha de tu apartado mensual"
        fondo = "#e9f7ff"
        acento = "#147fd1"
    else:
        intro = "Queremos recordarte que se acerca la fecha de tu apartado mensual en Bravo."
        destacado = "Faltan 3 días para la fecha de tu apartado mensual"
        fondo = "#f1edff"
        acento = "#5b45c6"
    return f"""<!doctype html><html><body style="margin:0;background:#f4f5f9;font-family:Arial;color:#525b82">
<table width="100%"><tr><td align="center"><table width="700" style="max-width:700px;background:#fff;border-top:6px solid #38278f">
<tr><td style="padding:28px 48px"><img src="https://drive.google.com/uc?export=view&id=13kK3v4FiyXFa4UzM_au3TllhOhwjvWb7" width="170"></td></tr>
<tr><td style="padding:5px 48px;font-size:30px;font-weight:800;color:#11183f">Hola, <span style="color:#35238f">{nombre}</span> 👋</td></tr>
<tr><td style="padding:12px 48px 24px;font-size:18px;line-height:28px">{intro}</td></tr>
<tr><td style="padding:0 48px 24px"><table width="100%" style="background:{fondo};border-radius:18px"><tr>
<td style="padding:28px;font-size:55px">📅</td><td style="padding:28px 28px 28px 0">
<div style="font-size:14px;font-weight:800;color:{acento};text-transform:uppercase">Fecha de tu apartado mensual</div>
<div style="font-size:29px;font-weight:800;color:#11183f;margin-top:6px">{ft}</div>
<div style="background:#fff;border-radius:12px;padding:13px 16px;margin-top:18px;font-size:18px;font-weight:800;color:{acento}">{destacado}</div>
</td></tr></table></td></tr>
<tr><td style="padding:0 48px 20px;font-size:17px;line-height:27px">Tener presente la fecha de tu apartado mensual te ayuda a mantener tu programa al día y continuar avanzando en tu proceso.</td></tr>
<tr><td style="padding:0 48px 28px"><a href="https://wa.me/573012411885" style="display:block;background:#12bd70;color:white;text-align:center;padding:18px;border-radius:14px;text-decoration:none;font-size:19px;font-weight:800">Hablar con Bravo por WhatsApp</a></td></tr>
<tr><td align="center" style="padding:22px;border-top:1px solid #ddd"><b>¡Gracias por ser parte de Bravo!</b><br>Equipo Bravo</td></tr>
</table></td></tr></table></body></html>"""


def enviar_gmail(svc, destinatario, asunto, cuerpo):
    msg = EmailMessage()
    msg["To"] = destinatario
    msg["From"] = f"Bravo S.A.S. <{GMAIL_FROM}>"
    msg["Reply-To"] = GMAIL_REPLY_TO
    msg["Subject"] = asunto
    msg.set_content("Recordatorio de tu apartado mensual en Bravo.")
    msg.add_alternative(cuerpo, subtype="html")
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    return svc.users().messages().send(userId="me", body={"raw": raw}).execute()


def cargar_maestro_clientes(gc):
    valores = (
        gc.open_by_key(CARTERA_SPREADSHEET_ID)
        .worksheet(HOJA_CARTERA)
        .get("B:G")
    )
    if len(valores) <= 1:
        return {}

    encabezados = [str(x).strip() for x in valores[0]]
    pos = {c: i for i, c in enumerate(encabezados)}

    requeridas = {"Nombre_Cliente", "Email"}
    faltan = requeridas - set(encabezados)
    if faltan:
        raise RuntimeError(
            "Faltan columnas en 2. Cartera Berex: " + ", ".join(sorted(faltan))
        )

    llaves = [
        c for c in ("Referencia", "Referencia_Berex", "Numero")
        if c in pos
    ]
    maestro = {}

    for fila in valores[1:]:
        def valor(col):
            i = pos[col]
            return fila[i] if len(fila) > i else ""

        nombre = str(valor("Nombre_Cliente")).strip()
        email = str(valor("Email")).strip().lower()

        for columna in llaves:
            referencia = normalizar_referencia(valor(columna))
            if not referencia:
                continue

            actual = maestro.get(
                referencia,
                {"NOMBRE": "", "EMAIL": ""}
            )
            if nombre:
                actual["NOMBRE"] = nombre
            if email:
                actual["EMAIL"] = email
            maestro[referencia] = actual

    return maestro


def cargar_ids_existentes(gc):
    valores = (
        gc.open_by_key(MASIVOS_SPREADSHEET_ID)
        .worksheet("COLA_ENVIO")
        .get_all_values()
    )
    if len(valores) <= 1:
        return set()

    encabezados = [str(x).strip() for x in valores[0]]
    if "ID_ENVIO" not in encabezados:
        raise RuntimeError("COLA_ENVIO no tiene la columna ID_ENVIO.")

    i_id = encabezados.index("ID_ENVIO")
    return {
        str(fila[i_id]).strip()
        for fila in valores[1:]
        if len(fila) > i_id and str(fila[i_id]).strip()
    }


def main():
    ahora = datetime.now(TZ)
    hoy = ahora.date()
    print("=" * 72)
    print("RECORDATORIOS DE APARTADO MENSUAL - PRODUCCIÓN")
    print(f"Fecha/hora Colombia: {ahora.strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 72)

    gc = cliente_sheets()
    fuente = gc.open_by_key(UNIDOS_EST_SPREADSHEET_ID)
    unidos = fuente.worksheet(HOJA_UNIDOS_EST).get("A:I")
    exclusiones = {
        normalizar_referencia(x)
        for x in fuente.worksheet(HOJA_EXCLUIR).col_values(1)[1:]
        if normalizar_referencia(x)
    }
    maestro = cargar_maestro_clientes(gc)

    libro = gc.open_by_key(MASIVOS_SPREADSHEET_ID)
    cola = libro.worksheet("COLA_ENVIO")
    valores_cola = cola.get_all_values()
    cab = [str(x).strip() for x in valores_cola[0]]
    idx = {c: i for i, c in enumerate(cab)}
    requeridas = {"ID_ENVIO","ID_CAMPAÑA","REFERENCIA","NOMBRE","EMAIL","PLANTILLA",
                  "ASUNTO","ESTADO","FECHA_PROG","FECHA_ENVIO","INTENTOS","ERROR",
                  "ID_MENSAJE","CUERPO","ENCARGADO"}
    faltan = requeridas - set(cab)
    if faltan:
        raise RuntimeError("Faltan columnas en COLA_ENVIO: " + ", ".join(sorted(faltan)))
    existentes = {
        str(f[idx["ID_ENVIO"]]).strip()
        for f in valores_cola[1:]
        if len(f) > idx["ID_ENVIO"] and str(f[idx["ID_ENVIO"]]).strip()
    }

    candidatos = {}
    for fila in unidos[1:]:
        referencia = normalizar_referencia(fila[0] if len(fila) > 0 else "")
        status = str(fila[2] if len(fila) > 2 else "").strip()
        d = parsear_fecha(fila[7] if len(fila) > 7 else "")
        if not referencia or normalizar(status) != "AL DIA" or not d:
            continue
        dias = (d - hoy).days
        if dias not in (0, 3):
            continue
        plantilla = "ALDIA000" if dias == 0 else "ALDIA003"
        candidatos[(referencia, d.isoformat(), plantilla)] = (referencia, d, dias, plantilla)

    listos = []
    excl = dup = sinmail = 0
    for referencia, d, dias, plantilla in candidatos.values():
        ide = f"ENV-ALDIA-{d.strftime('%Y%m%d')}-{referencia}-{plantilla}"
        if referencia in exclusiones:
            excl += 1; continue
        if ide in existentes:
            dup += 1; continue
        cli = maestro.get(referencia, {})
        nombre = str(cli.get("NOMBRE", "")).strip() or "Cliente"
        email = str(cli.get("EMAIL", "")).strip().lower()
        if not correo_valido(email):
            sinmail += 1; continue
        asunto = ("Hoy es la fecha de tu apartado mensual | Bravo"
                  if dias == 0 else
                  "Se acerca la fecha de tu apartado mensual | Bravo")
        listos.append({
            "ID_ENVIO": ide, "REFERENCIA": referencia, "FECHA": d, "DIAS": dias,
            "PLANTILLA": plantilla, "NOMBRE": nombre, "EMAIL": email,
            "ASUNTO": asunto, "CUERPO": html_al_dia(nombre, d, dias)
        })

    print(f"Candidatos por fecha/status: {len(candidatos)}")
    print(f"Excluidos: {excl} | Duplicados: {dup} | Sin correo: {sinmail}")
    print(f"LISTOS PARA ENVIAR: {len(listos)}")
    if not listos:
        print("No hay envíos nuevos.")
        return

    svc = cliente_gmail()
    enviados = errores = 0
    camp = f"ALDIA-AUTO-{hoy.strftime('%Y%m%d')}"

    for item in listos:
        datos = {
            "ID_ENVIO": item["ID_ENVIO"], "ID_CAMPAÑA": camp,
            "REFERENCIA": item["REFERENCIA"], "NOMBRE": item["NOMBRE"],
            "EMAIL": item["EMAIL"], "PLANTILLA": item["PLANTILLA"],
            "ASUNTO": item["ASUNTO"], "ESTADO": "ENVIANDO",
            "FECHA_PROG": ahora.strftime("%d/%m/%Y %H:%M:%S"),
            "FECHA_ENVIO": "", "INTENTOS": 1, "ERROR": "",
            "ID_MENSAJE": "", "CUERPO": item["CUERPO"], "ENCARGADO": ""
        }
        cola.append_row([datos.get(c, "") for c in cab], value_input_option="USER_ENTERED")
        fila_sheet = len(cola.get_all_values())
        try:
            resp = enviar_gmail(svc, item["EMAIL"], item["ASUNTO"], item["CUERPO"])
            cola.update_cell(fila_sheet, idx["FECHA_ENVIO"] + 1, datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S"))
            cola.update_cell(fila_sheet, idx["ID_MENSAJE"] + 1, str(resp.get("id", "")))
            cola.update_cell(fila_sheet, idx["ESTADO"] + 1, "ENVIADO")
            enviados += 1
            print(f"ENVIADO | {item['PLANTILLA']} | {item['REFERENCIA']} | {item['FECHA'].strftime('%d/%m/%Y')}")
        except Exception as e:
            cola.update_cell(fila_sheet, idx["ERROR"] + 1, str(e)[:500])
            cola.update_cell(fila_sheet, idx["ESTADO"] + 1, "ERROR")
            errores += 1
            print(f"ERROR | {item['PLANTILLA']} | {item['REFERENCIA']} | {e}")

    print("=" * 72)
    print(f"ENVIADOS: {enviados}")
    print(f"ERRORES: {errores}")
    print("=" * 72)


if __name__ == "__main__":
    main()
