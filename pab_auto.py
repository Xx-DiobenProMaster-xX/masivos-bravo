import os
import json
import re
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials


TZ = ZoneInfo("America/Bogota")

MASIVOS_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
CARTERA_BEREX_ID = "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
EXCLUSIONES_ID = "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"

HOJA_PAB = "PAB_PROXIMOS"
HOJA_CLIENTES = "CLIENTES"
HOJA_PLANTILLAS = "PLANTILLAS"
HOJA_COLA = "COLA_ENVIO"
HOJA_EXCLUIR = "Excluir_correo"

INFO_CLIENTES_V2 = [
    "Info_Clientes_V2",
    "Hoja Info_Clientes_V2",
    ". Hoja Info_Clientes_V2",
]


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
    hoja
):

    gc = obtener_gc()

    return (
        gc
        .open_by_key(archivo_id)
        .worksheet(hoja)
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

    for formato in [
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
    ]:

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
            fila[h["Referencia"]]
        )

        if not ref:
            continue

        resultado[ref] = {

            "nombre": (
                texto(
                    fila[h["Nombre"]]
                )
                if "Nombre" in h
                else ""
            ),

            "email": (
                texto(
                    fila[h["Email"]]
                )
                if "Email" in h
                else ""
            ),

            "mora": (
                texto(
                    fila[h["Mora"]]
                )
                if "Mora" in h
                else ""
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

    return {
        referencia(fila[0])
        for fila in datos[1:]
        if fila
        and referencia(fila[0])
    }


# ============================================================
# PLANTILLAS
# ============================================================

def cargar_plantillas():

    hoja = obtener_hoja(
        MASIVOS_ID,
        HOJA_PLANTILLAS
    )

    datos = hoja.get_all_values()

    h = mapa_headers(
        datos[0]
    )

    resultado = {}

    for fila in datos[1:]:

        plantilla = texto(
            fila[
                h["ID_PLANTILLA"]
            ]
        ).upper()

        if not plantilla:
            continue

        resultado[plantilla] = texto(
            fila[h["ESTADO"]]
        ).upper()

    return resultado


# ============================================================
# IDs YA GENERADOS
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

    return {

        texto(
            fila[h["ID_ENVIO"]]
        )

        for fila in datos[1:]

        if (
            len(fila)
            > h["ID_ENVIO"]
        )
    }


# ============================================================
# SIMULACIÓN
# ============================================================

def simular():

    hoja_pab = obtener_hoja(
        MASIVOS_ID,
        HOJA_PAB
    )

    pab = hoja_pab.get_all_values()

    h = mapa_headers(
        pab[0]
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

    resultado = {

        "candidatos": 0,
        "se_agregarian": 0,
        "sin_email": 0,
        "mora_180": 0,
        "excluir_correo": 0,
        "duplicados": 0,
        "ejemplos": []
    }

    for fila in pab[1:]:

        ref = referencia(
            fila[h["REFERENCIA"]]
        )

        fecha = convertir_fecha(
            fila[h["FECHA_PAB"]]
        )

        if not ref or not fecha:
            continue

        dias = (
            fecha - hoy
        ).days

        if dias == 3:

            plantilla = "PAB003"

        elif dias == 0:

            plantilla = "PAB000"

        else:

            continue

        resultado[
            "candidatos"
        ] += 1

        email = (
            texto(
                fila[h["EMAIL"]]
            )
            if "EMAIL" in h
            else ""
        )

        mora = (
            texto(
                fila[h["MORA"]]
            )
            if "MORA" in h
            else ""
        )

        cliente = clientes.get(
            ref,
            {}
        )

        respaldo = info_v2.get(
            ref,
            {}
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

        if not email:

            resultado[
                "sin_email"
            ] += 1

            continue

        if es_mora_180(
            mora
        ):

            resultado[
                "mora_180"
            ] += 1

            continue

        if ref in exclusiones:

            resultado[
                "excluir_correo"
            ] += 1

            continue

        if (
            plantillas.get(
                plantilla
            )
            != "ACTIVA"
        ):

            raise RuntimeError(
                f"{plantilla} no está ACTIVA."
            )

        id_envio = (

            f"ENV-PAB-"
            f"{fecha.strftime('%Y%m%d')}-"
            f"{ref}-"
            f"{plantilla}"

        )

        if id_envio in ids_cola:

            resultado[
                "duplicados"
            ] += 1

            continue

        resultado[
            "se_agregarian"
        ] += 1

        if (
            len(
                resultado[
                    "ejemplos"
                ]
            )
            < 10
        ):

            resultado[
                "ejemplos"
            ].append({

                "referencia": ref,
                "email": email,
                "plantilla": plantilla,
                "id_envio": id_envio

            })

    print(
        json.dumps(
            resultado,
            indent=2,
            ensure_ascii=False
        )
    )


if __name__ == "__main__":

    simular()
