import os
import re
import json
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials as ServiceAccountCredentials


# ============================================================
# MODO DIAGNÓSTICO - NO ENVÍA NI MODIFICA NADA
# ============================================================

TZ = ZoneInfo("America/Bogota")

MASIVOS_SPREADSHEET_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
UNIDOS_EST_SPREADSHEET_ID = "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"
HOJA_UNIDOS_EST = "Unidos_Est"
HOJA_EXCLUIR = "Excluir_correo"

CARTERA_SPREADSHEET_ID = "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
HOJAS_INFO_CLIENTES_V2 = ["Info_Clientes_V2", "Hoja Info_Clientes_V2", ". Hoja Info_Clientes_V2"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


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


def cargar_maestro_clientes(gc):
    """
    Misma fuente de respaldo usada por la app:
    Info_Clientes_V2
    C = Referencia
    E = Nombre cliente
    F = Email
    """
    archivo = gc.open_by_key(CARTERA_SPREADSHEET_ID)

    hoja = None
    nombre_encontrado = None
    for nombre_hoja in HOJAS_INFO_CLIENTES_V2:
        try:
            hoja = archivo.worksheet(nombre_hoja)
            nombre_encontrado = nombre_hoja
            break
        except Exception:
            continue

    if hoja is None:
        disponibles = [ws.title for ws in archivo.worksheets()]
        raise RuntimeError(
            "No encontré Info_Clientes_V2. Probé: "
            + ", ".join(HOJAS_INFO_CLIENTES_V2)
            + ". Pestañas disponibles: "
            + ", ".join(disponibles)
        )

    valores = hoja.get("C:F")
    if len(valores) <= 1:
        print(f"Fuente clientes encontrada: {nombre_encontrado}, pero está vacía.")
        return {}

    maestro = {}
    for fila in valores[1:]:
        referencia = normalizar_referencia(
            fila[0] if len(fila) > 0 else ""
        )
        if not referencia:
            continue

        nombre = str(
            fila[2] if len(fila) > 2 else ""
        ).strip()
        email = str(
            fila[3] if len(fila) > 3 else ""
        ).strip().lower()

        actual = maestro.get(
            referencia,
            {"NOMBRE": "", "EMAIL": ""}
        )

        if nombre and not actual["NOMBRE"]:
            actual["NOMBRE"] = nombre
        if email and not actual["EMAIL"]:
            actual["EMAIL"] = email

        maestro[referencia] = actual

    print(f"Fuente clientes encontrada: {nombre_encontrado}")
    print(f"Referencias cargadas desde Info_Clientes_V2: {len(maestro)}")
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
    print("DIAGNÓSTICO - RECORDATORIOS DE APARTADO MENSUAL")
    print("=" * 72)
    print(f"Fecha/hora Colombia: {ahora.strftime('%d/%m/%Y %H:%M:%S')}")
    print("MODO DIAGNÓSTICO: NINGÚN CORREO SERÁ ENVIADO")
    print("MODO DIAGNÓSTICO: NO SE MODIFICARÁ NINGÚN GOOGLE SHEET")
    print()

    gc = cliente_sheets()

    fuente = gc.open_by_key(UNIDOS_EST_SPREADSHEET_ID)
    unidos = fuente.worksheet(HOJA_UNIDOS_EST).get("A:I")
    excluidos_raw = fuente.worksheet(HOJA_EXCLUIR).col_values(1)

    exclusiones = {
        normalizar_referencia(x)
        for x in excluidos_raw[1:]
        if normalizar_referencia(x)
    }

    maestro = cargar_maestro_clientes(gc)
    ids_existentes = cargar_ids_existentes(gc)

    total_filas = 0
    total_al_dia = 0
    fechas_invalidas = 0

    candidatos_3 = {}
    candidatos_0 = {}

    for fila in unidos[1:]:
        total_filas += 1

        referencia = normalizar_referencia(
            fila[0] if len(fila) > 0 else ""
        )
        status = str(
            fila[2] if len(fila) > 2 else ""
        ).strip()
        fecha = parsear_fecha(
            fila[7] if len(fila) > 7 else ""
        )

        if not referencia or normalizar(status) != "AL DIA":
            continue

        total_al_dia += 1

        if not fecha:
            fechas_invalidas += 1
            continue

        dias = (fecha - hoy).days
        registro = {
            "REFERENCIA": referencia,
            "FECHA": fecha,
            "DIAS": dias,
        }

        # Deduplicar por referencia + fecha.
        llave = (referencia, fecha.isoformat())
        if dias == 3:
            candidatos_3[llave] = registro
        elif dias == 0:
            candidatos_0[llave] = registro

    resumen = {
        "ALDIA003": {
            "brutos": len(candidatos_3),
            "excluidos": 0,
            "duplicados": 0,
            "sin_correo": 0,
            "listos": [],
        },
        "ALDIA000": {
            "brutos": len(candidatos_0),
            "excluidos": 0,
            "duplicados": 0,
            "sin_correo": 0,
            "listos": [],
        },
    }

    for plantilla, candidatos in (
        ("ALDIA003", candidatos_3),
        ("ALDIA000", candidatos_0),
    ):
        for registro in candidatos.values():
            referencia = registro["REFERENCIA"]
            fecha = registro["FECHA"]

            if referencia in exclusiones:
                resumen[plantilla]["excluidos"] += 1
                continue

            id_envio = (
                f"ENV-ALDIA-{fecha.strftime('%Y%m%d')}-"
                f"{referencia}-{plantilla}"
            )

            if id_envio in ids_existentes:
                resumen[plantilla]["duplicados"] += 1
                continue

            cliente = maestro.get(referencia, {})
            email = str(cliente.get("EMAIL", "")).strip().lower()
            nombre = str(cliente.get("NOMBRE", "")).strip()

            if not correo_valido(email):
                resumen[plantilla]["sin_correo"] += 1
                continue

            resumen[plantilla]["listos"].append({
                "REFERENCIA": referencia,
                "NOMBRE": nombre or "Cliente",
                "EMAIL": email,
                "FECHA": fecha.strftime("%d/%m/%Y"),
                "ID_ENVIO": id_envio,
            })

    print(f"Filas revisadas en Unidos_Est: {total_filas}")
    print(f"Filas con status actual 'Al día': {total_al_dia}")
    print(f"Filas Al día con fecha inválida/vacía en H: {fechas_invalidas}")
    print(f"Referencias en Excluir_correo: {len(exclusiones)}")
    print()

    for plantilla, titulo in (
        ("ALDIA003", "3 DÍAS ANTES"),
        ("ALDIA000", "MISMO DÍA"),
    ):
        r = resumen[plantilla]
        print("-" * 72)
        print(f"{plantilla} - {titulo}")
        print("-" * 72)
        print(f"Candidatos por fecha: {r['brutos']}")
        print(f"Excluidos por Excluir_correo: {r['excluidos']}")
        print(f"Ya existentes en COLA_ENVIO: {r['duplicados']}")
        print(f"Sin correo válido: {r['sin_correo']}")
        print(f"LISTOS PARA ENVIAR: {len(r['listos'])}")
        print()

        # Por seguridad no imprimimos emails completos en logs.
        for x in r["listos"][:20]:
            correo = x["EMAIL"]
            partes = correo.split("@", 1)
            correo_mask = (
                (partes[0][:2] + "***@" + partes[1])
                if len(partes) == 2 else "***"
            )
            print(
                f"  {x['REFERENCIA']} | {x['FECHA']} | "
                f"{correo_mask} | {x['NOMBRE']}"
            )

        if len(r["listos"]) > 20:
            print(
                f"  ... y {len(r['listos']) - 20} candidatos adicionales"
            )
        print()

    total_listos = (
        len(resumen["ALDIA003"]["listos"])
        + len(resumen["ALDIA000"]["listos"])
    )

    print("=" * 72)
    print(f"TOTAL DE CORREOS QUE SE ENVIARÍAN HOY: {total_listos}")
    print("=" * 72)
    print("DIAGNÓSTICO FINALIZADO.")
    print("NINGÚN CORREO FUE ENVIADO.")
    print("NINGUNA FILA FUE AGREGADA O MODIFICADA.")


if __name__ == "__main__":
    main()
