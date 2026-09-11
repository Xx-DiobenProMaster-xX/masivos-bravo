import os
import json
import re
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURACIÓN
# ============================================================

TZ = ZoneInfo("America/Bogota")

MASIVOS_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"

CARTERA_BEREX_ID = (
    "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
)

EXCLUSIONES_ID = (
    "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"
)

HOJA_PAB = "PAB_PROXIMOS"
HOJA_CLIENTES = "CLIENTES"
HOJA_PLANTILLAS = "PLANTILLAS"
HOJA_COLA = "COLA_ENVIO"
HOJA_CAMPANAS = "CAMPAÑAS"
HOJA_EXCLUIR = "Excluir_correo"

INFO_CLIENTES_V2 = [
    "Info_Clientes_V2",
    "Hoja Info_Clientes_V2",
    ". Hoja Info_Clientes_V2",
]

# IMPORTANTE:
# BORRADOR = se agrega a COLA_ENVIO pero NO debe ser enviado.
ESTADO_NUEVO = "BORRADOR"


# ============================================================
# GOOGLE
# ============================================================

def obtener_gc():

    secreto = os.environ.get(
        "MI_JSON",
        ""
    ).strip()

    if not secreto:
        raise RuntimeError(
            "No existe MI_JSON en GitHub Secrets."
        )

    info = json.loads(
        secreto
    )

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credenciales = (
        Credentials
        .from_service_account_info(
            info,
            scopes=scopes
        )
    )

    return gspread.authorize(
        credenciales
    )


def obtener_hoja(
    archivo_id,
    nombre_hoja
):

    gc = obtener_gc()

    return (
        gc
        .open_by_key(archivo_id)
        .worksheet(nombre_hoja)
    )


# ============================================================
# AUXILIARES
# ============================================================

def texto(valor):

    if valor is None:
        return ""

    return str(valor).strip()


def referencia(valor):

    valor = texto(valor)

    if re.fullmatch(
        r"\d+\.0+",
        valor
    ):
        valor = valor.split(".")[0]

    return re.sub(
        r"\D",
        "",
        valor
    )


def mapa_headers(headers):

    return {
        str(valor).strip(): i
        for i, valor in enumerate(headers)
    }


def convertir_fecha(valor):

    valor = texto(valor)

    if not valor:
        return None

    formatos = [
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%Y/%m/%d",
    ]

    for formato in formatos:

        try:

            return datetime.strptime(
                valor[:10],
                formato
            ).date()

        except ValueError:
            pass

    return None


def es_mora_180(valor):

    valor = (
        texto(valor)
        .upper()
        .replace("_", " ")
    )

    valor = re.sub(
        r"\s+",
        " ",
        valor
    ).strip()

    return valor in {
        "MORA 180",
        "180"
    }


def valor_fila(
    fila,
    headers,
    nombre
):

    if nombre not in headers:
        return ""

    posicion = headers[nombre]

    if len(fila) <= posicion:
        return ""

    return texto(
        fila[posicion]
    )


def reemplazar_variables(
    contenido,
    variables
):

    contenido = texto(
        contenido
    )

    for variable, valor in variables.items():

        contenido = contenido.replace(
            variable,
            str(valor)
        )

    return contenido


def moneda(valor):

    valor = texto(valor)

    if not valor:
        return "$0"

    limpio = (
        valor
        .replace("$", "")
        .replace("COP", "")
        .replace(" ", "")
    )

    # 8.000.000
    if (
        "." in limpio
        and "," not in limpio
    ):

        partes = limpio.split(".")

        if all(
            len(parte) == 3
            for parte in partes[1:]
        ):
            limpio = "".join(partes)

    # 8,000,000
    elif (
        "," in limpio
        and "." not in limpio
    ):

        limpio = limpio.replace(
            ",",
            ""
        )

    # 8.000.000,50
    elif (
        "." in limpio
        and "," in limpio
    ):

        limpio = (
            limpio
            .replace(".", "")
            .replace(",", ".")
        )

    limpio = re.sub(
        r"[^0-9.\-]",
        "",
        limpio
    )

    try:

        numero = float(
            limpio
        )

    except ValueError:

        numero = 0

    return (
        "$"
        + f"{numero:,.0f}".replace(
            ",",
            "."
        )
    )


# ============================================================
# INFO CLIENTES V2
# ============================================================

def cargar_info_clientes_v2():

    gc = obtener_gc()

    archivo = gc.open_by_key(
        CARTERA_BEREX_ID
    )

    hoja = None

    for nombre in INFO_CLIENTES_V2:

        try:

            hoja = archivo.worksheet(
                nombre
            )

            break

        except gspread.WorksheetNotFound:
            continue

    if hoja is None:

        raise RuntimeError(
            "No encontré Info_Clientes_V2."
        )

    # C = Referencia
    # E = Nombre
    # F = Email

    datos = hoja.get(
        "C:F"
    )

    resultado = {}

    for fila in datos[1:]:

        ref = referencia(
            fila[0]
            if len(fila) > 0
            else ""
        )

        if not ref:
            continue

        nombre = texto(
            fila[2]
            if len(fila) > 2
            else ""
        )

        email = texto(
            fila[3]
            if len(fila) > 3
            else ""
        )

        if ref not in resultado:

            resultado[ref] = {
                "nombre": "",
                "email": ""
            }

        if (
            not resultado[ref]["nombre"]
            and nombre
        ):

            resultado[ref][
                "nombre"
            ] = nombre

        if (
            not resultado[ref]["email"]
            and email
        ):

            resultado[ref][
                "email"
            ] = email

    return resultado


# ============================================================
# CLIENTES
# ============================================================

def cargar_clientes():

    hoja = obtener_hoja(
        MASIVOS_ID,
        HOJA_CLIENTES
    )

    datos = hoja.get_all_values()

    if len(datos) <= 1:
        return {}

    h = mapa_headers(
        datos[0]
    )

    resultado = {}

    for fila in datos[1:]:

        ref = referencia(
            valor_fila(
                fila,
                h,
                "Referencia"
            )
        )

        if not ref:
            continue

        resultado[ref] = {

            "nombre": valor_fila(
                fila,
                h,
                "Nombre"
            ),

            "email": valor_fila(
                fila,
                h,
                "Email"
            ),

            "mora": valor_fila(
                fila,
                h,
                "Mora"
            ),

            "encargado": valor_fila(
                fila,
                h,
                "Encargado"
            )
        }

    return resultado


# ============================================================
# EXCLUSIONES
# ============================================================

def cargar_exclusiones():

    hoja = obtener_hoja(
        EXCLUSIONES_ID,
        HOJA_EXCLUIR
    )

    datos = hoja.get_all_values()

    resultado = set()

    for fila in datos[1:]:

        if not fila:
            continue

        ref = referencia(
            fila[0]
        )

        if ref:
            resultado.add(
                ref
            )

    return resultado


# ============================================================
# PLANTILLAS
# ============================================================

def cargar_plantillas():

    hoja = obtener_hoja(
        MASIVOS_ID,
        HOJA_PLANTILLAS
    )

    datos = hoja.get_all_values()

    if len(datos) <= 1:
        return {}

    h = mapa_headers(
        datos[0]
    )

    resultado = {}

    for fila in datos[1:]:

        # Compatible con ID_PLANTILLA o PLANTILLA
        plantilla = ""

        if "ID_PLANTILLA" in h:

            plantilla = valor_fila(
                fila,
                h,
                "ID_PLANTILLA"
            )

        elif "PLANTILLA" in h:

            plantilla = valor_fila(
                fila,
                h,
                "PLANTILLA"
            )

        plantilla = plantilla.upper()

        if not plantilla:
            continue

        asunto = valor_fila(
            fila,
            h,
            "ASUNTO"
        )

        # Compatible con CUERPO o HTML
        cuerpo = ""

        if "CUERPO" in h:

            cuerpo = valor_fila(
                fila,
                h,
                "CUERPO"
            )

        elif "HTML" in h:

            cuerpo = valor_fila(
                fila,
                h,
                "HTML"
            )

        estado = valor_fila(
            fila,
            h,
            "ESTADO"
        ).upper()

        # Si no existe estado,
        # asumimos ACTIVA
        if not estado:
            estado = "ACTIVA"

        resultado[
            plantilla
        ] = {

            "asunto": asunto,

            "cuerpo": cuerpo,

            "estado": estado
        }

    return resultado


# ============================================================
# COLA EXISTENTE
# ============================================================

def cargar_ids_cola():

    hoja = obtener_hoja(
        MASIVOS_ID,
        HOJA_COLA
    )

    datos = hoja.get_all_values()

    if len(datos) <= 1:
        return set()

    h = mapa_headers(
        datos[0]
    )

    if "ID_ENVIO" not in h:
        return set()

    resultado = set()

    for fila in datos[1:]:

        id_envio = valor_fila(
            fila,
            h,
            "ID_ENVIO"
        )

        if id_envio:

            resultado.add(
                id_envio
            )

    return resultado


# ============================================================
# CREAR CAMPAÑA
# ============================================================

def asegurar_campana(
    id_campana,
    cantidad
):

    hoja = obtener_hoja(
        MASIVOS_ID,
        HOJA_CAMPANAS
    )

    datos = hoja.get_all_values()

    if not datos:

        raise RuntimeError(
            "CAMPAÑAS no tiene encabezados."
        )

    headers = datos[0]

    h = mapa_headers(
        headers
    )

    # Verificar si ya existe
    for fila in datos[1:]:

        actual = valor_fila(
            fila,
            h,
            "ID_CAMPAÑA"
        )

        if actual == id_campana:
            return

    nueva = [
        ""
    ] * len(headers)

    valores = {

        "ID_CAMPAÑA":
            id_campana,

        "NOMBRE_CAMPAÑA":
            "Recordatorios PaB automáticos",

        "PLANTILLA":
            "PAB",

        "FILTRO":
            "PAB",

        "FECHA_ENVIO":
            datetime.now(
                TZ
            ).strftime(
                "%d/%m/%Y"
            ),

        "HORA_ENVIO":
            datetime.now(
                TZ
            ).strftime(
                "%H:%M"
            ),

        "ESTADO":
            "BORRADOR",

        "TOTAL_CLIENTES":
            cantidad,

        "ENVIADOS":
            0,

        "PENDIENTES":
            0,

        "ERRORES":
            0,

        "FECHA_CREACIÓN":
            datetime.now(
                TZ
            ).strftime(
                "%d/%m/%Y %H:%M:%S"
            ),

        "COMENTARIOS":
            "Creada automáticamente por GitHub Actions"
    }

    for columna, valor in valores.items():

        if columna in h:

            nueva[
                h[columna]
            ] = valor

    hoja.append_row(
        nueva,
        value_input_option="USER_ENTERED"
    )


# ============================================================
# AUTOMATIZACIÓN PAB
# ============================================================

def generar_recordatorios():

    hoja_pab = obtener_hoja(
        MASIVOS_ID,
        HOJA_PAB
    )

    hoja_cola = obtener_hoja(
        MASIVOS_ID,
        HOJA_COLA
    )

    pab = hoja_pab.get_all_values()

    cola = hoja_cola.get_all_values()

    if len(pab) <= 1:

        print(
            "PAB_PROXIMOS no tiene registros."
        )

        return

    if not cola:

        raise RuntimeError(
            "COLA_ENVIO no tiene encabezados."
        )

    hp = mapa_headers(
        pab[0]
    )

    headers_cola = cola[0]

    hc = mapa_headers(
        headers_cola
    )

    clientes = cargar_clientes()

    info_v2 = (
        cargar_info_clientes_v2()
    )

    exclusiones = (
        cargar_exclusiones()
    )

    plantillas = (
        cargar_plantillas()
    )

    ids_cola = (
        cargar_ids_cola()
    )

    hoy = datetime.now(
        TZ
    ).date()

    id_campana = (
        "PAB-"
        + hoy.strftime(
            "%Y%m%d"
        )
    )

    resumen = {

        "candidatos": 0,

        "agregados_borrador": 0,

        "sin_email": 0,

        "mora_180": 0,

        "excluir_correo": 0,

        "duplicados": 0,

        "ejemplos": []
    }

    filas_nuevas = []

    for fila in pab[1:]:

        ref = referencia(
            valor_fila(
                fila,
                hp,
                "REFERENCIA"
            )
        )

        fecha = convertir_fecha(
            valor_fila(
                fila,
                hp,
                "FECHA_PAB"
            )
        )

        if not ref or not fecha:
            continue

        dias = (
            fecha - hoy
        ).days

        if dias == 3:

            plantilla_id = "PAB003"

        elif dias == 0:

            plantilla_id = "PAB000"

        else:
            continue

        resumen[
            "candidatos"
        ] += 1

        # =============================
        # DATOS DEL CLIENTE
        # =============================

        nombre = valor_fila(
            fila,
            hp,
            "NOMBRE"
        )

        email = valor_fila(
            fila,
            hp,
            "EMAIL"
        )

        mora = valor_fila(
            fila,
            hp,
            "MORA"
        )

        encargado = valor_fila(
            fila,
            hp,
            "ENCARGADO"
        )

        cliente = clientes.get(
            ref,
            {}
        )

        respaldo = info_v2.get(
            ref,
            {}
        )

        if not nombre:

            nombre = (
                cliente.get(
                    "nombre",
                    ""
                )
                or
                respaldo.get(
                    "nombre",
                    ""
                )
            )

        if not email:

            email = (
                cliente.get(
                    "email",
                    ""
                )
                or
                respaldo.get(
                    "email",
                    ""
                )
            )

        if not mora:

            mora = cliente.get(
                "mora",
                ""
            )

        if not encargado:

            encargado = cliente.get(
                "encargado",
                ""
            )

        # =============================
        # VALIDACIONES
        # =============================

        if not email:

            resumen[
                "sin_email"
            ] += 1

            continue

        if es_mora_180(
            mora
        ):

            resumen[
                "mora_180"
            ] += 1

            continue

        if ref in exclusiones:

            resumen[
                "excluir_correo"
            ] += 1

            continue

        plantilla = plantillas.get(
            plantilla_id
        )

        if not plantilla:

            raise RuntimeError(
                f"No encontré {plantilla_id}."
            )

        if (
            plantilla[
                "estado"
            ]
            != "ACTIVA"
        ):

            raise RuntimeError(
                f"{plantilla_id} no está ACTIVA."
            )

        # =============================
        # VARIABLES
        # =============================

        valor_pab = valor_fila(
            fila,
            hp,
            "VALOR_PAB"
        )

        variables = {

            "{{NOMBRE}}":
                nombre,

            "{{REFERENCIA}}":
                ref,

            "{{FECHA_PAB}}":
                fecha.strftime(
                    "%d/%m/%Y"
                ),

            "{{VALOR_PAB}}":
                moneda(
                    valor_pab
                ),

            "{{TIPO_AVISO}}":
                (
                    "Pago programado para hoy"
                    if plantilla_id == "PAB000"
                    else
                    "Recordatorio 3 días antes"
                )
        }

        asunto = reemplazar_variables(
            plantilla[
                "asunto"
            ],
            variables
        )

        cuerpo = reemplazar_variables(
            plantilla[
                "cuerpo"
            ],
            variables
        )

        # =============================
        # ID ÚNICO
        # =============================

        id_envio = (

            f"ENV-PAB-"
            f"{fecha.strftime('%Y%m%d')}-"
            f"{ref}-"
            f"{plantilla_id}"

        )

        if id_envio in ids_cola:

            resumen[
                "duplicados"
            ] += 1

            continue

        # =============================
        # CONSTRUIR FILA COLA_ENVIO
        # =============================

        nueva = [
            ""
        ] * len(headers_cola)

        valores = {

            "ID_ENVIO":
                id_envio,

            "ID_CAMPAÑA":
                id_campana,

            "REFERENCIA":
                ref,

            "NOMBRE":
                nombre,

            "EMAIL":
                email,

            "PLANTILLA":
                plantilla_id,

            "ASUNTO":
                asunto,

            # MUY IMPORTANTE
            "ESTADO":
                ESTADO_NUEVO,

            "FECHA_PROG":
                datetime.now(
                    TZ
                ).strftime(
                    "%d/%m/%Y %H:%M"
                ),

            "FECHA_ENVIO":
                "",

            "INTENTOS":
                0,

            "ERROR":
                "",

            "ID_MENSAJE":
                "",

            "CUERPO":
                cuerpo,

            "ENCARGADO":
                encargado
        }

        for columna, valor in valores.items():

            if columna in hc:

                nueva[
                    hc[columna]
                ] = valor

        filas_nuevas.append(
            nueva
        )

        ids_cola.add(
            id_envio
        )

        resumen[
            "agregados_borrador"
        ] += 1

        if (
            len(
                resumen[
                    "ejemplos"
                ]
            )
            < 10
        ):

            resumen[
                "ejemplos"
            ].append({

                "referencia":
                    ref,

                "nombre":
                    nombre,

                "email":
                    email,

                "plantilla":
                    plantilla_id,

                "asunto":
                    asunto,

                "estado":
                    ESTADO_NUEVO,

                "id_envio":
                    id_envio

            })

    # ========================================================
    # ESCRIBIR
    # ========================================================

    if filas_nuevas:

        asegurar_campana(
            id_campana,
            len(filas_nuevas)
        )

        hoja_cola.append_rows(
            filas_nuevas,
            value_input_option="USER_ENTERED"
        )

    print(
        json.dumps(
            resumen,
            indent=2,
            ensure_ascii=False
        )
    )

    print(
        "\nIMPORTANTE:"
        "\nLos registros fueron creados como BORRADOR."
        "\nNo deben ser enviados por el motor automático."
    )


if __name__ == "__main__":

    generar_recordatorios()
