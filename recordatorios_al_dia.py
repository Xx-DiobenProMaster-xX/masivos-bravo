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
# MODO DIAGNÓSTICO
# NO ENVÍA CORREOS
# NO MODIFICA COLA_ENVIO
# ============================================================

TZ = ZoneInfo("America/Bogota")


# ============================================================
# CONFIGURACIÓN
# ============================================================

MASIVOS_SPREADSHEET_ID = (
    "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
)

UNIDOS_EST_SPREADSHEET_ID = (
    "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"
)

HOJA_UNIDOS_EST = "Unidos_Est"
HOJA_EXCLUIR = "Excluir_correo"


# ------------------------------------------------------------
# CARTERA
# ------------------------------------------------------------

CARTERA_SPREADSHEET_ID = (
    "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
)

HOJA_CARTERA = "Cartera"

HOJAS_INFO_CLIENTES_V2 = [
    "Info_Clientes_V2",
    "Hoja Info_Clientes_V2",
    ". Hoja Info_Clientes_V2",
]


# ------------------------------------------------------------
# DF_MORA_ESTADOS
# ------------------------------------------------------------

DF_MORA_ESTADOS_ID = (
    "1jcPPhtF2YK3Kr7P_A0Mgh2OqhOfnVWB2to3UPoSH5tE"
)

HOJA_COMISION = "Comisión"


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ------------------------------------------------------------
# Gmail
#
# Se conservan las funciones porque luego volveremos a
# activar producción, pero main() NO las utiliza.
# ------------------------------------------------------------

GMAIL_FROM = "estructurados@gobravo.com.co"
GMAIL_REPLY_TO = "estructurados@gobravo.com.co"

GMAIL_SCOPES = [
    "https://www.googleapis.com/auth/gmail.send"
]


# ============================================================
# NORMALIZACIÓN
# ============================================================

def normalizar(texto):

    import unicodedata

    t = unicodedata.normalize(
        "NFD",
        str(texto or "").strip()
    )

    return "".join(
        c
        for c in t
        if unicodedata.category(c) != "Mn"
    ).upper()


def normalizar_referencia(valor):

    if valor is None:
        return ""

    texto = (
        str(valor)
        .strip()
        .replace("\xa0", "")
        .replace(" ", "")
    )

    if not texto:
        return ""

    if texto.upper() in {
        "NAN",
        "NONE",
        "NULL",
    }:
        return ""

    # Ejemplo:
    # 123456.0
    if re.fullmatch(
        r"[+-]?\d+\.0+",
        texto
    ):
        return (
            texto
            .split(".")[0]
            .lstrip("+")
        )

    # Ejemplo:
    # 1.234.567
    # 1,234,567
    if re.fullmatch(
        r"[+-]?\d{1,3}([.,]\d{3})+",
        texto
    ):
        return (
            re.sub(
                r"[.,]",
                "",
                texto
            )
            .lstrip("+")
        )

    if re.fullmatch(
        r"[+-]?\d+",
        texto
    ):
        return texto.lstrip("+")

    try:

        numero = float(
            texto.replace(",", "")
        )

        if numero.is_integer():
            return str(int(numero))

    except Exception:
        pass

    return re.sub(
        r"\D",
        "",
        texto
    )


# ============================================================
# FECHAS
# ============================================================

def parsear_fecha(valor):

    texto = str(
        valor or ""
    ).strip()

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

            return datetime.strptime(
                texto,
                formato
            ).date()

        except Exception:
            pass

    return None


# ============================================================
# NÚMEROS
# ============================================================

def parsear_numero(valor):

    texto = str(
        valor or ""
    ).strip()

    texto = (
        texto
        .replace("$", "")
        .replace(" ", "")
    )

    if not texto:
        return 0.0

    try:

        # Ejemplo:
        # 1,234.56
        # 1.234,56
        if "," in texto and "." in texto:

            if texto.rfind(".") > texto.rfind(","):

                texto = texto.replace(
                    ",",
                    ""
                )

            else:

                texto = (
                    texto
                    .replace(".", "")
                    .replace(",", ".")
                )

        elif "," in texto:

            partes = texto.split(",")

            if len(partes[-1]) in (1, 2):

                texto = (
                    texto
                    .replace(".", "")
                    .replace(",", ".")
                )

            else:

                texto = texto.replace(
                    ",",
                    ""
                )

        return float(texto)

    except Exception:
        return 0.0


# ============================================================
# VALIDACIÓN CORREO
# ============================================================

def correo_valido(valor):

    correo = str(
        valor or ""
    ).strip().lower()

    return bool(
        re.fullmatch(
            r"[^@\s]+@[^@\s]+\.[^@\s]+",
            correo
        )
    )


# ============================================================
# GOOGLE SHEETS
# ============================================================

def cliente_sheets():

    if not os.environ.get("MI_JSON"):

        raise RuntimeError(
            "No encontré el Secret MI_JSON."
        )

    info = json.loads(
        os.environ["MI_JSON"]
    )

    credenciales = (
        ServiceAccountCredentials
        .from_service_account_info(
            info,
            scopes=SCOPES,
        )
    )

    return gspread.authorize(
        credenciales
    )


# ============================================================
# GMAIL
#
# NO SE USA EN MODO DIAGNÓSTICO
# ============================================================

def cliente_gmail():

    faltan = [
        k
        for k in (
            "GOOGLE_CLIENT_ID",
            "GOOGLE_CLIENT_SECRET",
            "GOOGLE_REFRESH_TOKEN",
        )
        if not os.environ.get(k)
    ]

    if faltan:

        raise RuntimeError(
            "Faltan variables Gmail: "
            + ", ".join(faltan)
        )

    cred = OAuthCredentials(

        token=None,

        refresh_token=os.environ[
            "GOOGLE_REFRESH_TOKEN"
        ],

        token_uri=(
            "https://oauth2.googleapis.com/token"
        ),

        client_id=os.environ[
            "GOOGLE_CLIENT_ID"
        ],

        client_secret=os.environ[
            "GOOGLE_CLIENT_SECRET"
        ],

        scopes=GMAIL_SCOPES,
    )

    cred.refresh(
        Request()
    )

    return build(
        "gmail",
        "v1",
        credentials=cred,
        cache_discovery=False,
    )


# ============================================================
# FORMATO FECHA
# ============================================================

def fecha_larga(d):

    meses = [
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre",
    ]

    return (
        f"{d.day} de "
        f"{meses[d.month - 1]} de "
        f"{d.year}"
    )


# ============================================================
# HTML
#
# SE CONSERVA SIN CAMBIOS PARA PRODUCCIÓN
# ============================================================

def html_al_dia(
    nombre,
    d,
    dias,
):

    import html as html_lib

    nombre = html_lib.escape(
        str(
            nombre
            or "Cliente"
        )
    )

    ft = fecha_larga(d)

    if dias == 0:

        intro = (
            "Te recordamos que hoy es la fecha "
            "de tu apartado mensual en Bravo."
        )

        destacado = (
            "Hoy es la fecha de tu "
            "apartado mensual"
        )

        fondo = "#e9f7ff"
        acento = "#147fd1"

    else:

        intro = (
            "Queremos recordarte que se acerca "
            "la fecha de tu apartado mensual "
            "en Bravo."
        )

        destacado = (
            "Faltan 3 días para la fecha "
            "de tu apartado mensual"
        )

        fondo = "#f1edff"
        acento = "#5b45c6"

    return f"""
<!doctype html>
<html>
<body
style="
margin:0;
background:#f4f5f9;
font-family:Arial;
color:#525b82
"
>

<table width="100%">

<tr>

<td align="center">

<table
width="700"
style="
max-width:700px;
background:#fff;
border-top:6px solid #38278f
"
>

<tr>

<td
style="
padding:28px 48px
"
>

<img
src="https://drive.google.com/uc?export=view&id=13kK3v4FiyXFa4UzM_au3TllhOhwjvWb7"
width="170"
>

</td>

</tr>


<tr>

<td
style="
padding:5px 48px;
font-size:30px;
font-weight:800;
color:#11183f
"
>

Hola,

<span
style="
color:#35238f
"
>

{nombre}

</span>

👋

</td>

</tr>


<tr>

<td
style="
padding:12px 48px 24px;
font-size:18px;
line-height:28px
"
>

{intro}

</td>

</tr>


<tr>

<td
style="
padding:0 48px 24px
"
>

<table
width="100%"
style="
background:{fondo};
border-radius:18px
"
>

<tr>

<td
style="
padding:28px;
font-size:55px
"
>

📅

</td>

<td
style="
padding:28px 28px 28px 0
"
>

<div
style="
font-size:14px;
font-weight:800;
color:{acento};
text-transform:uppercase
"
>

Fecha de tu apartado mensual

</div>


<div
style="
font-size:29px;
font-weight:800;
color:#11183f;
margin-top:6px
"
>

{ft}

</div>


<div
style="
background:#fff;
border-radius:12px;
padding:13px 16px;
margin-top:18px;
font-size:18px;
font-weight:800;
color:{acento}
"
>

{destacado}

</div>

</td>

</tr>

</table>

</td>

</tr>


<tr>

<td
style="
padding:0 48px 20px;
font-size:17px;
line-height:27px
"
>

Tener presente la fecha de tu apartado mensual
te ayuda a mantener tu programa al día y
continuar avanzando en tu proceso.

</td>

</tr>


<tr>

<td
style="
padding:0 48px 28px
"
>

<a
href="https://wa.me/573012411885"
style="
display:block;
background:#12bd70;
color:white;
text-align:center;
padding:18px;
border-radius:14px;
text-decoration:none;
font-size:19px;
font-weight:800
"
>

Hablar con Bravo por WhatsApp

</a>

</td>

</tr>


<tr>

<td
align="center"
style="
padding:22px;
border-top:1px solid #ddd
"
>

<b>
¡Gracias por ser parte de Bravo!
</b>

<br>

Equipo Bravo

</td>

</tr>

</table>

</td>

</tr>

</table>

</body>

</html>
"""


# ============================================================
# ENVÍO GMAIL
#
# SE CONSERVA PARA PRODUCCIÓN,
# PERO NO SE LLAMA EN main()
# ============================================================

def enviar_gmail(
    svc,
    destinatario,
    asunto,
    cuerpo,
):

    msg = EmailMessage()

    msg["To"] = destinatario

    msg["From"] = (
        f"Bravo S.A.S. "
        f"<{GMAIL_FROM}>"
    )

    msg["Reply-To"] = (
        GMAIL_REPLY_TO
    )

    msg["Subject"] = asunto

    msg.set_content(
        "Recordatorio de tu "
        "apartado mensual en Bravo."
    )

    msg.add_alternative(
        cuerpo,
        subtype="html",
    )

    raw = (
        base64
        .urlsafe_b64encode(
            msg.as_bytes()
        )
        .decode("utf-8")
    )

    return (
        svc
        .users()
        .messages()
        .send(
            userId="me",
            body={
                "raw": raw
            },
        )
        .execute()
    )


# ============================================================
# MAESTRO DE CLIENTES
# ============================================================

def cargar_maestro_clientes(gc):

    """
    Info_Clientes_V2

    C = Referencia
    E = Nombre cliente
    F = Email
    """

    archivo = gc.open_by_key(
        CARTERA_SPREADSHEET_ID
    )

    hoja = None
    nombre_encontrado = None

    for nombre_hoja in (
        HOJAS_INFO_CLIENTES_V2
    ):

        try:

            hoja = archivo.worksheet(
                nombre_hoja
            )

            nombre_encontrado = (
                nombre_hoja
            )

            break

        except Exception:
            continue

    if hoja is None:

        disponibles = [
            ws.title
            for ws in archivo.worksheets()
        ]

        raise RuntimeError(
            "No encontré Info_Clientes_V2. "
            "Probé: "
            + ", ".join(
                HOJAS_INFO_CLIENTES_V2
            )
            + ". Pestañas disponibles: "
            + ", ".join(disponibles)
        )

    valores = hoja.get(
        "C:F"
    )

    maestro = {}

    for fila in valores[1:]:

        referencia = (
            normalizar_referencia(
                fila[0]
                if len(fila) > 0
                else ""
            )
        )

        if not referencia:
            continue

        nombre = str(
            fila[2]
            if len(fila) > 2
            else ""
        ).strip()

        email = str(
            fila[3]
            if len(fila) > 3
            else ""
        ).strip().lower()

        actual = maestro.get(
            referencia,
            {
                "NOMBRE": "",
                "EMAIL": "",
            },
        )

        if (
            nombre
            and not actual["NOMBRE"]
        ):

            actual["NOMBRE"] = nombre

        if (
            email
            and not actual["EMAIL"]
        ):

            actual["EMAIL"] = email

        maestro[
            referencia
        ] = actual

    print(
        "Fuente clientes encontrada: "
        f"{nombre_encontrado}"
    )

    print(
        "Referencias cargadas desde "
        "Info_Clientes_V2: "
        f"{len(maestro)}"
    )

    return maestro


# ============================================================
# NUEVA LÓGICA
# CARTERA -> COMMISSION DEL MES ACTUAL
# ============================================================

def cargar_commission_mes_actual(
    gc,
    hoy,
):

    """
    Cartera

    A = reference
    F = destination
    G = payment_date

    REGLA:

    - destination = commission
    - payment_date debe pertenecer
      al MES ACTUAL
    - payment_date debe pertenecer
      al AÑO ACTUAL

    NO importa si es fin de mes.

    Ejemplo:

    Hoy = octubre 2026

    30/09/2026 -> NO
    09/10/2026 -> SÍ
    15/10/2026 -> SÍ
    31/10/2026 -> SÍ
    30/11/2026 -> NO
    """

    archivo = gc.open_by_key(
        CARTERA_SPREADSHEET_ID
    )

    hoja = archivo.worksheet(
        HOJA_CARTERA
    )

    valores = hoja.get(
        "A:G"
    )

    fechas_por_ref = {}

    for fila in valores[1:]:

        referencia = (
            normalizar_referencia(
                fila[0]
                if len(fila) > 0
                else ""
            )
        )

        destination = str(
            fila[5]
            if len(fila) > 5
            else ""
        ).strip().lower()

        fecha_pago = parsear_fecha(
            fila[6]
            if len(fila) > 6
            else ""
        )

        if not referencia:
            continue

        if (
            destination
            != "commission"
        ):
            continue

        if not fecha_pago:
            continue

        # ----------------------------------------
        # SOLO MES/AÑO ACTUAL
        # ----------------------------------------

        if (
            fecha_pago.year
            != hoy.year
        ):
            continue

        if (
            fecha_pago.month
            != hoy.month
        ):
            continue

        fechas_por_ref.setdefault(
            referencia,
            set(),
        ).add(
            fecha_pago
        )

    fechas_validas = {}
    anomalas = {}

    for (
        referencia,
        fechas_ref,
    ) in fechas_por_ref.items():

        ordenadas = sorted(
            fechas_ref
        )

        # Una única fecha
        # commission este mes
        if len(ordenadas) == 1:

            fechas_validas[
                referencia
            ] = ordenadas[0]

        # Más de una fecha diferente
        # en el mismo mes.
        #
        # Por seguridad NO decidimos
        # automáticamente cuál utilizar.
        else:

            anomalas[
                referencia
            ] = ordenadas

    return (
        fechas_validas,
        anomalas,
    )


# ============================================================
# VALIDACIÓN DE PAGOS
# DF_MORA_ESTADOS -> COMISIÓN
# ============================================================

def cargar_validacion_comision(
    gc,
    hoy,
):

    """
    Hoja Comisión

    A = REFERENCIA
    B = FECHA
    C = X_COBRAR
    D = PAGO
    I = FECHA_COBRO
    N = MORA_STATUS

    Solamente analizamos el mes/año actual.

    Una referencia se considera cubierta
    cuando:

    PAGO >= X_COBRAR - 100 COP

    La tolerancia de 100 COP evita que
    diferencias mínimas generen un correo
    incorrecto.
    """

    archivo = gc.open_by_key(
        DF_MORA_ESTADOS_ID
    )

    hoja = archivo.worksheet(
        HOJA_COMISION
    )

    valores = hoja.get(
        "A:N"
    )

    resumen = {}

    for fila in valores[1:]:

        referencia = (
            normalizar_referencia(
                fila[0]
                if len(fila) > 0
                else ""
            )
        )

        fecha_periodo = parsear_fecha(
            fila[1]
            if len(fila) > 1
            else ""
        )

        if not referencia:
            continue

        if not fecha_periodo:
            continue

        # Solo periodo actual
        if (
            fecha_periodo.year
            != hoy.year
        ):
            continue

        if (
            fecha_periodo.month
            != hoy.month
        ):
            continue

        x_cobrar = parsear_numero(
            fila[2]
            if len(fila) > 2
            else ""
        )

        pago = parsear_numero(
            fila[3]
            if len(fila) > 3
            else ""
        )

        fecha_cobro = parsear_fecha(
            fila[8]
            if len(fila) > 8
            else ""
        )

        mora_status = str(
            fila[13]
            if len(fila) > 13
            else ""
        ).strip()

        item = resumen.setdefault(
            referencia,
            {
                "X_COBRAR": 0.0,
                "PAGO": 0.0,
                "FECHA_COBRO": None,
                "MORA_STATUS": "",
            },
        )

        # Puede haber más de una fila
        # de la misma referencia.
        item[
            "X_COBRAR"
        ] += x_cobrar

        item[
            "PAGO"
        ] += pago

        # Conservamos la fecha
        # de cobro más reciente.
        if fecha_cobro:

            if (
                item["FECHA_COBRO"]
                is None
                or fecha_cobro
                > item["FECHA_COBRO"]
            ):

                item[
                    "FECHA_COBRO"
                ] = fecha_cobro

        if mora_status:

            item[
                "MORA_STATUS"
            ] = mora_status

    # ----------------------------------------
    # DEFINIR SI YA CUBRIÓ
    # ----------------------------------------

    for item in resumen.values():

        x_cobrar = item[
            "X_COBRAR"
        ]

        pago = item[
            "PAGO"
        ]

        item[
            "CUBIERTO"
        ] = (
            x_cobrar > 0
            and pago
            >= (
                x_cobrar
                - 100
            )
        )

    return resumen


# ============================================================
# IDS EXISTENTES
#
# Se conserva para producción.
# En diagnóstico NO se necesita modificar nada.
# ============================================================

def cargar_ids_existentes(gc):

    valores = (
        gc
        .open_by_key(
            MASIVOS_SPREADSHEET_ID
        )
        .worksheet(
            "COLA_ENVIO"
        )
        .get_all_values()
    )

    if len(valores) <= 1:
        return set()

    encabezados = [
        str(x).strip()
        for x in valores[0]
    ]

    if (
        "ID_ENVIO"
        not in encabezados
    ):

        raise RuntimeError(
            "COLA_ENVIO no tiene "
            "la columna ID_ENVIO."
        )

    i_id = encabezados.index(
        "ID_ENVIO"
    )

    return {
        str(
            fila[i_id]
        ).strip()

        for fila in valores[1:]

        if (
            len(fila) > i_id
            and str(
                fila[i_id]
            ).strip()
        )
    }


# ============================================================
# MAIN
# ============================================================

def main():

    ahora = datetime.now(
        TZ
    )

    hoy = ahora.date()

    print(
        "=" * 90
    )

    print(
        "RECORDATORIOS AL DÍA "
        "- MODO DIAGNÓSTICO"
    )

    print(
        "NO ENVÍA CORREOS "
        "Y NO MODIFICA COLA_ENVIO"
    )

    print(
        "Fecha/hora Colombia: "
        + ahora.strftime(
            "%d/%m/%Y %H:%M:%S"
        )
    )

    print(
        "=" * 90
    )


    # ========================================================
    # CONEXIÓN SHEETS
    # ========================================================

    gc = cliente_sheets()


    # ========================================================
    # EXCLUSIONES
    # ========================================================

    fuente = gc.open_by_key(
        UNIDOS_EST_SPREADSHEET_ID
    )

    exclusiones = {

        normalizar_referencia(x)

        for x in (
            fuente
            .worksheet(
                HOJA_EXCLUIR
            )
            .col_values(1)[1:]
        )

        if normalizar_referencia(x)
    }


    # ========================================================
    # DATOS CLIENTE
    # ========================================================

    maestro = (
        cargar_maestro_clientes(
            gc
        )
    )


    # ========================================================
    # FECHA COMMISSION
    # CARTERA
    # ========================================================

    (
        fechas_mes,
        anomalas,
    ) = (
        cargar_commission_mes_actual(
            gc,
            hoy,
        )
    )


    # ========================================================
    # PAGOS
    # DF_MORA_ESTADOS
    # ========================================================

    pagos_mes = (
        cargar_validacion_comision(
            gc,
            hoy,
        )
    )


    print()

    print(
        "Commission del mes actual: "
        f"{len(fechas_mes)}"
    )

    print(
        "Referencias con más de una "
        "fecha commission este mes: "
        f"{len(anomalas)}"
    )


    # ========================================================
    # ANOMALÍAS
    # ========================================================

    if anomalas:

        print()

        print(
            "=" * 90
        )

        print(
            "ANOMALÍAS DE FECHA"
        )

        print(
            "NO SE ENVIARÍAN "
            "AUTOMÁTICAMENTE"
        )

        print(
            "=" * 90
        )

        for (
            referencia,
            fechas,
        ) in list(
            anomalas.items()
        )[:50]:

            texto_fechas = (
                ", ".join(
                    d.strftime(
                        "%d/%m/%Y"
                    )
                    for d in fechas
                )
            )

            print(
                "ANOMALIA | "
                f"{referencia} | "
                f"{texto_fechas}"
            )


    # ========================================================
    # CANDIDATOS DEL DÍA
    # ========================================================

    print()

    print(
        "=" * 90
    )

    print(
        "CANDIDATOS DE HOY"
    )

    print(
        "=" * 90
    )


    candidatos_hoy = 0
    enviar = 0
    no_enviar = 0


    for (
        referencia,
        fecha_pago,
    ) in sorted(

        fechas_mes.items(),

        key=lambda x: (
            x[1],
            x[0],
        ),
    ):

        dias = (
            fecha_pago
            - hoy
        ).days


        # ====================================================
        # REGLA DE ENVÍO
        #
        # 3 días antes = ALDIA003
        # mismo día    = ALDIA000
        # ====================================================

        if dias not in (
            0,
            3,
        ):
            continue


        candidatos_hoy += 1


        plantilla = (

            "ALDIA000"

            if dias == 0

            else "ALDIA003"
        )


        # ====================================================
        # DATOS CLIENTE
        # ====================================================

        cli = maestro.get(
            referencia,
            {},
        )

        nombre = str(
            cli.get(
                "NOMBRE",
                "",
            )
        ).strip()

        if not nombre:
            nombre = "Cliente"


        email = str(
            cli.get(
                "EMAIL",
                "",
            )
        ).strip().lower()


        # ====================================================
        # INFORMACIÓN DE PAGO
        # ====================================================

        pago_info = pagos_mes.get(

            referencia,

            {
                "X_COBRAR": 0.0,
                "PAGO": 0.0,
                "FECHA_COBRO": None,
                "MORA_STATUS": "",
                "CUBIERTO": False,
            },
        )


        # ====================================================
        # DECISIÓN
        # ====================================================

        motivos = []


        # ----------------------------------------------------
        # Exclusión manual
        # ----------------------------------------------------

        if (
            referencia
            in exclusiones
        ):

            motivos.append(
                "EXCLUIR_CORREO"
            )


        # ----------------------------------------------------
        # Correo inválido
        # ----------------------------------------------------

        if not correo_valido(
            email
        ):

            motivos.append(
                "SIN_CORREO_VALIDO"
            )


        # ----------------------------------------------------
        # Ya pagó
        # ----------------------------------------------------

        if pago_info.get(
            "CUBIERTO",
            False,
        ):

            motivos.append(
                "COMPROMISO_YA_CUBIERTO"
            )


        # ----------------------------------------------------
        # Debe estar Al día
        # ----------------------------------------------------

        status = normalizar(
            pago_info.get(
                "MORA_STATUS",
                "",
            )
        )

        if status != "AL DIA":

            if status:

                motivos.append(
                    "STATUS_NO_AL_DIA"
                )

            else:

                motivos.append(
                    "SIN_STATUS_COMISION"
                )


        # ====================================================
        # RESULTADO
        # ====================================================

        if motivos:

            decision = (
                "NO ENVIAR"
            )

            no_enviar += 1

        else:

            decision = (
                "ENVIAR"
            )

            enviar += 1


        # ====================================================
        # FECHA COBRO
        # ====================================================

        fecha_cobro = (
            pago_info.get(
                "FECHA_COBRO"
            )
        )

        if fecha_cobro:

            fecha_cobro_txt = (
                fecha_cobro.strftime(
                    "%d/%m/%Y"
                )
            )

        else:

            fecha_cobro_txt = "-"


        # ====================================================
        # MOSTRAR DIAGNÓSTICO
        # ====================================================

        print(
            f"{decision} | "
            f"{plantilla} | "
            f"Ref {referencia} | "
            f"{nombre} | "
            f"Fecha "
            f"{fecha_pago.strftime('%d/%m/%Y')} | "
            f"X_COBRAR "
            f"{pago_info.get('X_COBRAR', 0):,.2f} | "
            f"PAGO "
            f"{pago_info.get('PAGO', 0):,.2f} | "
            f"FECHA_COBRO "
            f"{fecha_cobro_txt} | "
            f"STATUS "
            f"{pago_info.get('MORA_STATUS', '') or '-'} | "
            f"{', '.join(motivos) if motivos else 'OK'}"
        )


    # ========================================================
    # RESUMEN
    # ========================================================

    print()

    print(
        "=" * 90
    )

    print(
        "RESUMEN"
    )

    print(
        "=" * 90
    )

    print(
        "Candidatos por fecha hoy: "
        f"{candidatos_hoy}"
    )

    print(
        "Resultado ENVIAR: "
        f"{enviar}"
    )

    print(
        "Resultado NO ENVIAR: "
        f"{no_enviar}"
    )

    print()

    print(
        "MODO DIAGNÓSTICO:"
    )

    print(
        "0 correos enviados"
    )

    print(
        "0 filas modificadas"
    )

    print(
        "=" * 90
    )


# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":
    main()
