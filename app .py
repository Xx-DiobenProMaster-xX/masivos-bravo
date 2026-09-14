
import streamlit as st
import pandas as pd
import gspread
import google.auth
import json
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import unicodedata
import re

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Masivos Bravo",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

SPREADSHEET_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"

# Fuente donde vive Excluir_correo
EXCLUSIONES_SPREADSHEET_ID = "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"
HOJA_EXCLUIR_CORREO = "Excluir_correo"

# Base maestra para completar nombre y correo de clientes
CARTERA_BEREX_SPREADSHEET_ID = "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
HOJA_CARTERA_BEREX = "2. Cartera Berex"

# Segunda fuente de respaldo dentro del mismo archivo
# El código probará estos nombres de pestaña por si el nombre visible difiere.
HOJAS_INFO_CLIENTES_V2 = [
    "Info_Clientes_V2",
    "Hoja Info_Clientes_V2",
    ". Hoja Info_Clientes_V2"
]

TZ = ZoneInfo("America/Bogota")
AHORA = datetime.now(TZ)
HOY = AHORA.date()


# ============================================================
# ESTILOS
# ============================================================

st.markdown(
    """
    <style>

    .block-container {
        padding-top: 2.2rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(
            180deg,
            #37317e 0%,
            #292463 100%
        );
    }

    [data-testid="stSidebar"] * {
        color: white;
    }

    .titulo {
        font-size: 42px;
        line-height: 1.2;
        font-weight: 800;
        color: #37317e;
        margin-bottom: 5px;
    }

    .subtitulo {
        font-size: 18px;
        color: #737688;
        margin-bottom: 25px;
    }

    .pab-card {
        background: #ffffff;
        border: 1px solid #e7e7ef;
        border-radius: 16px;
        padding: 18px;
        margin-bottom: 12px;
    }

    .estado-ok {
        color: #159b69;
        font-weight: 700;
    }

    .estado-pendiente {
        color: #d69b00;
        font-weight: 700;
    }

    .estado-hoy {
        color: #d94a4a;
        font-weight: 700;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# GOOGLE SHEETS
# ============================================================

@st.cache_resource(show_spinner=False)
def obtener_gc():

    # --------------------------------------------------------
    # STREAMLIT CLOUD
    # Usa la cuenta de servicio guardada en Secrets
    # --------------------------------------------------------

    if "MI_JSON" in st.secrets:

        info = json.loads(
            st.secrets["MI_JSON"]
        )

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        credentials = (
            Credentials
            .from_service_account_info(
                info,
                scopes=scopes
            )
        )

    # --------------------------------------------------------
    # COLAB
    # Mantener compatibilidad para pruebas
    # --------------------------------------------------------

    else:

        credentials, _ = (
            google.auth.default()
        )

    return gspread.authorize(
        credentials
    )


@st.cache_resource(show_spinner=False)
def obtener_archivo():

    gc = obtener_gc()

    return gc.open_by_key(
        SPREADSHEET_ID
    )


@st.cache_data(ttl=300, show_spinner=False)
def cargar_hoja(nombre):

    archivo = obtener_archivo()

    hoja = archivo.worksheet(
        nombre
    )

    valores = hoja.get_all_values()

    if not valores:
        return pd.DataFrame()

    encabezados = valores[0]

    encabezados_finales = []
    usados = {}

    for i, encabezado in enumerate(encabezados):

        encabezado = str(
            encabezado
        ).strip()

        if not encabezado:
            encabezado = f"COLUMNA_{i+1}"

        if encabezado in usados:

            usados[encabezado] += 1

            encabezado = (
                f"{encabezado}_{usados[encabezado]}"
            )

        else:

            usados[encabezado] = 1

        encabezados_finales.append(
            encabezado
        )

    if len(valores) == 1:

        return pd.DataFrame(
            columns=encabezados_finales
        )

    df = pd.DataFrame(
        valores[1:],
        columns=encabezados_finales
    )

    df = df[
        ~df.apply(
            lambda fila:
            fila.astype(str)
            .str.strip()
            .eq("")
            .all(),
            axis=1
        )
    ].copy()

    return df


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def normalizar(valor):

    texto = str(
        valor or ""
    ).strip()

    texto = unicodedata.normalize(
        "NFD",
        texto
    )

    texto = "".join(
        x for x in texto
        if unicodedata.category(x) != "Mn"
    )

    return texto.upper()


def es_true(valor):

    return normalizar(valor) in {
        "TRUE",
        "VERDADERO",
        "SI",
        "1",
        "X"
    }


def convertir_fechas(serie):

    return pd.to_datetime(
        serie,
        errors="coerce",
        dayfirst=True
    )


def numero(valor):

    texto = str(
        valor or ""
    ).strip()

    if not texto:
        return 0.0

    texto = (
        texto
        .replace("$", "")
        .replace("COP", "")
        .replace(" ", "")
    )

    # 4,034,000
    if "," in texto and "." not in texto:
        texto = texto.replace(",", "")

    # 4.034.000
    elif "." in texto and "," not in texto:

        partes = texto.split(".")

        if len(partes) > 1 and all(
            len(p) == 3
            for p in partes[1:]
        ):
            texto = "".join(partes)

    # 4.034.000,50
    elif "." in texto and "," in texto:

        texto = (
            texto
            .replace(".", "")
            .replace(",", ".")
        )

    texto = re.sub(
        r"[^0-9.\-]",
        "",
        texto
    )

    try:
        return float(texto)

    except Exception:
        return 0.0


def moneda(valor):

    try:

        return (
            "$"
            + f"{float(valor):,.0f}"
            .replace(",", ".")
        )

    except Exception:

        return "$0"


def aviso_generado(valor):

    texto = normalizar(valor)

    return texto in {
        "GENERADO",
        "ENVIADO",
        "TRUE",
        "SI"
    }


# ============================================================
# PLANTILLAS PAB
# ============================================================

def obtener_plantilla_pab(id_plantilla):

    if plantillas.empty:
        return None

    if "ID_PLANTILLA" not in plantillas.columns:
        return None

    candidatos = plantillas[
        plantillas["ID_PLANTILLA"]
        .astype(str)
        .str.strip()
        .str.upper()
        == str(id_plantilla).strip().upper()
    ].copy()

    if candidatos.empty:
        return None

    if "ESTADO" in candidatos.columns:

        activas = candidatos[
            candidatos["ESTADO"]
            .astype(str)
            .str.strip()
            .str.upper()
            == "ACTIVA"
        ]

        if not activas.empty:
            candidatos = activas

    return candidatos.iloc[0]


def reemplazar_variables_pab(texto, fila, tipo_aviso):

    texto = str(texto or "")

    nombre = str(
        fila.get(
            "NOMBRE",
            ""
        )
    ).strip()

    referencia = str(
        fila.get(
            "REFERENCIA",
            ""
        )
    ).strip()

    fecha_original = str(
        fila.get(
            "FECHA_PAB",
            ""
        )
    ).strip()

    fecha_dt = fila.get(
        "_FECHA"
    )

    if pd.notna(fecha_dt):
        fecha_pab = fecha_dt.strftime("%d/%m/%Y")
    else:
        fecha_pab = fecha_original

    valor_original = fila.get(
        "VALOR_PAB",
        ""
    )

    valor_pab = moneda(
        numero(
            valor_original
        )
    )

    reemplazos = {
        "{{NOMBRE}}": nombre,
        "{{REFERENCIA}}": referencia,
        "{{FECHA_PAB}}": fecha_pab,
        "{{VALOR_PAB}}": valor_pab,
        "{{TIPO_AVISO}}": tipo_aviso
    }

    for variable, valor in reemplazos.items():
        texto = texto.replace(
            variable,
            str(valor)
        )

    return texto


def preparar_vista_previa_pab(fila):

    dias = fila.get(
        "_DIAS"
    )

    if dias == 0:
        id_plantilla = "PAB000"
        tipo_aviso = "Pago programado para hoy"

    elif dias == 3:
        id_plantilla = "PAB003"
        tipo_aviso = "Recordatorio 3 días antes"

    else:
        return None

    plantilla = obtener_plantilla_pab(
        id_plantilla
    )

    if plantilla is None:
        return {
            "id_plantilla": id_plantilla,
            "tipo_aviso": tipo_aviso,
            "error": (
                f"No encontré la plantilla {id_plantilla} "
                "en PLANTILLAS."
            )
        }

    asunto = reemplazar_variables_pab(
        plantilla.get(
            "ASUNTO",
            ""
        ),
        fila,
        tipo_aviso
    )

    cuerpo = reemplazar_variables_pab(
        plantilla.get(
            "CUERPO",
            ""
        ),
        fila,
        tipo_aviso
    )

    return {
        "id_plantilla": id_plantilla,
        "tipo_aviso": tipo_aviso,
        "asunto": asunto,
        "cuerpo": cuerpo,
        "error": None
    }



# ============================================================
# COLA PAB / SEGURIDAD
# ============================================================

def obtener_hoja_externa(spreadsheet_id, nombre_hoja):

    gc = obtener_gc()

    archivo = gc.open_by_key(
        spreadsheet_id
    )

    return archivo.worksheet(
        nombre_hoja
    )


def normalizar_referencia(valor):
    """
    Convierte referencias provenientes de distintas hojas a una llave común.

    Ejemplos que terminan como 3115580892:
    - 3115580892
    - "3115580892"
    - "3115580892.0"
    - "3.115.580.892"
    - "3,115,580,892"
    - " 3115580892 "
    """
    if valor is None:
        return ""

    texto = str(valor).strip()

    if not texto:
        return ""

    if texto.upper() in {"NAN", "NONE", "NULL"}:
        return ""

    # Quitar espacios normales y espacios no separables.
    texto = texto.replace("\xa0", "").replace(" ", "")

    # Caso típico de Sheets/Pandas: 3115580892.0
    if re.fullmatch(r"[+-]?\d+\.0+", texto):
        return texto.split(".")[0].lstrip("+")

    # Si es un entero con separadores de miles, quitarlos.
    if re.fullmatch(r"[+-]?\d{1,3}([.,]\d{3})+", texto):
        return re.sub(r"[.,]", "", texto).lstrip("+")

    # Si ya son solo dígitos, devolverlos.
    if re.fullmatch(r"[+-]?\d+", texto):
        return texto.lstrip("+")

    # Intentar notación científica o número decimal exacto.
    try:
        numero_ref = float(texto.replace(",", ""))
        if numero_ref.is_integer():
            return str(int(numero_ref))
    except Exception:
        pass

    # Último recurso: conservar solo dígitos.
    # Esto permite empatar referencias con caracteres invisibles o separadores.
    solo_digitos = re.sub(r"\D", "", texto)
    return solo_digitos


def valor_vacio(valor):

    texto = str(
        valor or ""
    ).strip()

    return (
        not texto
        or texto.upper() in {
            "NAN",
            "NONE",
            "NULL"
        }
    )


def entero_seguro(valor, default=0):
    """Convierte a entero sin romper la app por vacíos, NaN o texto no numérico."""
    try:
        if valor is None:
            return default
        texto = str(valor).strip()
        if not texto or texto.upper() in {"NAN", "NONE", "NULL"}:
            return default
        return int(float(texto))
    except Exception:
        try:
            return int(valor)
        except Exception:
            return default


@st.cache_data(ttl=600, show_spinner=False)
def cargar_maestro_cartera_berex():
    """
    Construye un diccionario de clientes desde 2. Cartera Berex.

    La búsqueda no depende de una sola columna: indexa cada cliente por
    Referencia, Referencia_Berex y Numero. Así cubrimos diferencias entre
    las referencias usadas por PAB_PROXIMOS y las guardadas en Cartera Berex.
    """

    hoja = obtener_hoja_externa(
        CARTERA_BEREX_SPREADSHEET_ID,
        HOJA_CARTERA_BEREX
    )

    # B:G:
    # B Referencia
    # C Referencia_Berex
    # D Cedula
    # E Nombre_Cliente
    # F Email
    # G Numero
    valores = hoja.get("B:G")

    if not valores:
        return {}

    encabezados = [
        str(x).strip()
        for x in valores[0]
    ]

    requeridos = {
        "Referencia",
        "Nombre_Cliente",
        "Email"
    }

    faltantes = requeridos - set(encabezados)

    if faltantes:
        raise ValueError(
            "Faltan columnas en 2. Cartera Berex: "
            + ", ".join(sorted(faltantes))
        )

    i_nombre = encabezados.index("Nombre_Cliente")
    i_email = encabezados.index("Email")

    columnas_llave = [
        c
        for c in [
            "Referencia",
            "Referencia_Berex",
            "Numero"
        ]
        if c in encabezados
    ]

    indices_llave = [
        encabezados.index(c)
        for c in columnas_llave
    ]

    maestro = {}

    def guardar_datos(llave, nombre, email):
        llave = normalizar_referencia(llave)

        if not llave:
            return

        if llave not in maestro:
            maestro[llave] = {
                "NOMBRE": "",
                "EMAIL": ""
            }

        # Solo completa; nunca reemplaza un dato válido por uno vacío.
        if (
            valor_vacio(maestro[llave].get("NOMBRE", ""))
            and not valor_vacio(nombre)
        ):
            maestro[llave]["NOMBRE"] = nombre

        if (
            valor_vacio(maestro[llave].get("EMAIL", ""))
            and not valor_vacio(email)
        ):
            maestro[llave]["EMAIL"] = email

    for fila in valores[1:]:
        nombre = str(
            fila[i_nombre]
            if len(fila) > i_nombre
            else ""
        ).strip()

        email = str(
            fila[i_email]
            if len(fila) > i_email
            else ""
        ).strip()

        for i_llave in indices_llave:
            llave = (
                fila[i_llave]
                if len(fila) > i_llave
                else ""
            )
            guardar_datos(
                llave,
                nombre,
                email
            )

    return maestro



@st.cache_data(ttl=600, show_spinner=False)
def cargar_maestro_info_clientes_v2():
    """
    Segunda fuente de respaldo.

    Columnas informadas por el usuario:
    C = Referencia
    E = Nombre cliente
    F = Email
    """

    gc = obtener_gc()
    archivo = gc.open_by_key(
        CARTERA_BEREX_SPREADSHEET_ID
    )

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
        raise ValueError(
            "No encontré la pestaña Info_Clientes_V2. "
            "Probé: " + ", ".join(HOJAS_INFO_CLIENTES_V2)
        )

    # C:F -> Referencia, (D), Nombre, Email
    valores = hoja.get("C:F")

    if not valores:
        return {}

    maestro = {}

    # No dependemos del texto exacto del encabezado;
    # usamos las posiciones indicadas por el usuario.
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
        ).strip()

        if referencia not in maestro:
            maestro[referencia] = {
                "NOMBRE": "",
                "EMAIL": "",
                "FUENTE": nombre_encontrado
            }

        if (
            valor_vacio(maestro[referencia].get("NOMBRE", ""))
            and not valor_vacio(nombre)
        ):
            maestro[referencia]["NOMBRE"] = nombre

        if (
            valor_vacio(maestro[referencia].get("EMAIL", ""))
            and not valor_vacio(email)
        ):
            maestro[referencia]["EMAIL"] = email

    return maestro


def enriquecer_pab_con_cartera_berex(df_pab):
    """
    Prioridad:
    1. Datos ya existentes en PAB_PROXIMOS
    2. Info_Clientes_V2

    Completa NOMBRE y EMAIL de forma independiente.
    """

    if df_pab.empty:
        return (
            df_pab.copy(),
            {
                "nombres_completados_info_v2": 0,
                "emails_completados_info_v2": 0,
                "sin_nombre": 0,
                "sin_email": 0,
                "error_info_v2": None
            }
        )

    df = df_pab.copy()

    if "NOMBRE" not in df.columns:
        df["NOMBRE"] = ""

    if "EMAIL" not in df.columns:
        df["EMAIL"] = ""

    if "REFERENCIA" not in df.columns:
        return (
            df,
            {
                "nombres_completados_info_v2": 0,
                "emails_completados_info_v2": 0,
                "sin_nombre": len(df),
                "sin_email": len(df),
                "error_info_v2": "PAB_PROXIMOS no tiene REFERENCIA."
            }
        )

    error_info_v2 = None

    try:
        maestro_info_v2 = cargar_maestro_info_clientes_v2()
    except Exception as e:
        maestro_info_v2 = {}
        error_info_v2 = str(e)

    nombres_completados = 0
    emails_completados = 0

    for idx in df.index:
        referencia = normalizar_referencia(
            df.at[idx, "REFERENCIA"]
        )

        if not referencia:
            continue

        datos = maestro_info_v2.get(referencia)

        if not datos:
            continue

        if valor_vacio(df.at[idx, "NOMBRE"]):
            nombre = datos.get("NOMBRE", "")

            if not valor_vacio(nombre):
                df.at[idx, "NOMBRE"] = nombre
                nombres_completados += 1

        if valor_vacio(df.at[idx, "EMAIL"]):
            email = datos.get("EMAIL", "")

            if not valor_vacio(email):
                df.at[idx, "EMAIL"] = email
                emails_completados += 1

    sin_nombre = int(
        df["NOMBRE"].apply(valor_vacio).sum()
    )

    sin_email = int(
        df["EMAIL"].apply(valor_vacio).sum()
    )

    return (
        df,
        {
            "nombres_completados_info_v2": nombres_completados,
            "emails_completados_info_v2": emails_completados,
            "sin_nombre": sin_nombre,
            "sin_email": sin_email,
            "error_info_v2": error_info_v2
        }
    )


@st.cache_data(ttl=600, show_spinner=False)
def obtener_referencias_excluidas():

    try:

        hoja = obtener_hoja_externa(
            EXCLUSIONES_SPREADSHEET_ID,
            HOJA_EXCLUIR_CORREO
        )

        valores = hoja.col_values(1)

        if not valores:
            return set()

        referencias = {
            str(v).strip()
            for v in valores[1:]
            if str(v).strip()
        }

        return referencias

    except Exception as e:

        raise PermissionError(
            "No pude consultar Excluir_correo. "
            "Por seguridad no se permitirá agregar el recordatorio "
            "a COLA_ENVIO hasta que la cuenta de servicio tenga acceso "
            f"al archivo de exclusiones. Detalle: {e}"
        )


def es_mora_180(valor):

    texto = normalizar(
        valor
    ).replace(
        "_",
        " "
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    ).strip()

    return texto in {
        "MORA 180",
        "180"
    }


def obtener_encabezados_hoja(hoja):

    encabezados = hoja.row_values(1)

    return [
        str(x).strip()
        for x in encabezados
    ]


def construir_fila_por_encabezados(
    encabezados,
    datos
):

    return [
        datos.get(
            encabezado,
            ""
        )
        for encabezado in encabezados
    ]


def existe_envio_pab_en_cola(
    referencia,
    id_plantilla,
    fecha_pab
):

    archivo = obtener_archivo()

    hoja = archivo.worksheet(
        "COLA_ENVIO"
    )

    valores = hoja.get_all_values()

    if len(valores) <= 1:
        return False

    encabezados = [
        str(x).strip()
        for x in valores[0]
    ]

    try:
        i_ref = encabezados.index(
            "REFERENCIA"
        )
        i_plantilla = encabezados.index(
            "PLANTILLA"
        )
        i_estado = encabezados.index(
            "ESTADO"
        )

    except ValueError:
        return False

    referencia = str(
        referencia
    ).strip()

    id_plantilla = str(
        id_plantilla
    ).strip().upper()

    for fila in valores[1:]:

        ref = (
            str(fila[i_ref]).strip()
            if len(fila) > i_ref
            else ""
        )

        plantilla = (
            str(fila[i_plantilla]).strip().upper()
            if len(fila) > i_plantilla
            else ""
        )

        estado = (
            str(fila[i_estado]).strip().upper()
            if len(fila) > i_estado
            else ""
        )

        if (
            ref == referencia
            and
            plantilla == id_plantilla
            and
            estado not in {
                "ERROR",
                "CANCELADO",
                "BLOQUEADO"
            }
        ):
            return True

    return False


def asegurar_campana_pab(
    id_campana
):

    archivo = obtener_archivo()

    hoja = archivo.worksheet(
        "CAMPAÑAS"
    )

    valores = hoja.get_all_values()

    if valores:

        encabezados = [
            str(x).strip()
            for x in valores[0]
        ]

    else:

        raise ValueError(
            "CAMPAÑAS no tiene encabezados."
        )

    if "ID_CAMPAÑA" not in encabezados:

        raise ValueError(
            "No encontré ID_CAMPAÑA en CAMPAÑAS."
        )

    i_id = encabezados.index(
        "ID_CAMPAÑA"
    )

    for fila in valores[1:]:

        valor = (
            str(fila[i_id]).strip()
            if len(fila) > i_id
            else ""
        )

        if valor == id_campana:
            return

    datos = {
        "ID_CAMPAÑA": id_campana,
        "NOMBRE_CAMPAÑA": (
            f"Recordatorios PaB "
            f"{HOY.strftime('%d/%m/%Y')}"
        ),
        "PLANTILLA": "PAB",
        "FILTRO": "PAB",
        "FECHA_ENVIO": HOY.strftime(
            "%d/%m/%Y"
        ),
        "HORA_ENVIO": AHORA.strftime(
            "%H:%M"
        ),
        "ESTADO": "PENDIENTE",
        "TOTAL_CLIENTES": 0,
        "ENVIADOS": 0,
        "PENDIENTES": 0,
        "ERRORES": 0,
        "FECHA_CREACIÓN": AHORA.strftime(
            "%d/%m/%Y %H:%M:%S"
        ),
        "COMENTARIOS": (
            "Campaña creada desde Masivos Bravo"
        )
    }

    fila_nueva = construir_fila_por_encabezados(
        encabezados,
        datos
    )

    hoja.append_row(
        fila_nueva,
        value_input_option="USER_ENTERED"
    )


def actualizar_contadores_campana_pab(
    id_campana
):

    archivo = obtener_archivo()

    hoja_cola = archivo.worksheet(
        "COLA_ENVIO"
    )

    cola_valores = hoja_cola.get_all_values()

    if len(cola_valores) <= 1:
        return

    encabezados_cola = [
        str(x).strip()
        for x in cola_valores[0]
    ]

    if (
        "ID_CAMPAÑA" not in encabezados_cola
        or
        "ESTADO" not in encabezados_cola
    ):
        return

    i_camp = encabezados_cola.index(
        "ID_CAMPAÑA"
    )
    i_estado = encabezados_cola.index(
        "ESTADO"
    )

    filas_campana = []

    for fila in cola_valores[1:]:

        camp = (
            str(fila[i_camp]).strip()
            if len(fila) > i_camp
            else ""
        )

        if camp == id_campana:
            filas_campana.append(
                fila
            )

    total = len(
        filas_campana
    )

    enviados = sum(
        1
        for fila in filas_campana
        if (
            str(
                fila[i_estado]
                if len(fila) > i_estado
                else ""
            ).strip().upper()
            == "ENVIADO"
        )
    )

    pendientes = sum(
        1
        for fila in filas_campana
        if (
            str(
                fila[i_estado]
                if len(fila) > i_estado
                else ""
            ).strip().upper()
            == "PENDIENTE"
        )
    )

    errores = sum(
        1
        for fila in filas_campana
        if (
            str(
                fila[i_estado]
                if len(fila) > i_estado
                else ""
            ).strip().upper()
            in {
                "ERROR",
                "BLOQUEADO"
            }
        )
    )

    hoja_camp = archivo.worksheet(
        "CAMPAÑAS"
    )

    camp_valores = hoja_camp.get_all_values()

    if len(camp_valores) <= 1:
        return

    encabezados_camp = [
        str(x).strip()
        for x in camp_valores[0]
    ]

    if "ID_CAMPAÑA" not in encabezados_camp:
        return

    i_id = encabezados_camp.index(
        "ID_CAMPAÑA"
    )

    fila_objetivo = None

    for numero_fila, fila in enumerate(
        camp_valores[1:],
        start=2
    ):

        camp = (
            str(fila[i_id]).strip()
            if len(fila) > i_id
            else ""
        )

        if camp == id_campana:
            fila_objetivo = numero_fila
            break

    if not fila_objetivo:
        return

    actualizaciones = {
        "TOTAL_CLIENTES": total,
        "ENVIADOS": enviados,
        "PENDIENTES": pendientes,
        "ERRORES": errores,
        "ESTADO": (
            "EN PROCESO"
            if pendientes > 0
            else "FINALIZADA"
        )
    }

    for encabezado, valor in actualizaciones.items():

        if encabezado in encabezados_camp:

            col = encabezados_camp.index(
                encabezado
            ) + 1

            hoja_camp.update_cell(
                fila_objetivo,
                col,
                valor
            )


def actualizar_estado_pab_proximos(
    id_evento,
    id_plantilla
):

    archivo = obtener_archivo()

    hoja = archivo.worksheet(
        "PAB_PROXIMOS"
    )

    valores = hoja.get_all_values()

    if len(valores) <= 1:
        raise ValueError(
            "PAB_PROXIMOS no contiene datos."
        )

    encabezados = [
        str(x).strip()
        for x in valores[0]
    ]

    if "ID_EVENTO" not in encabezados:

        raise ValueError(
            "No encontré ID_EVENTO en PAB_PROXIMOS."
        )

    i_evento = encabezados.index(
        "ID_EVENTO"
    )

    fila_objetivo = None

    for numero_fila, fila in enumerate(
        valores[1:],
        start=2
    ):

        evento = (
            str(fila[i_evento]).strip()
            if len(fila) > i_evento
            else ""
        )

        if evento == str(
            id_evento
        ).strip():

            fila_objetivo = numero_fila
            break

    if not fila_objetivo:

        raise ValueError(
            f"No encontré ID_EVENTO {id_evento}."
        )

    if id_plantilla == "PAB003":
        encabezado_estado = "AVISO_3_DIAS"

    else:
        encabezado_estado = "AVISO_HOY"

    if encabezado_estado not in encabezados:

        raise ValueError(
            f"No encontré {encabezado_estado} "
            "en PAB_PROXIMOS."
        )

    columna = encabezados.index(
        encabezado_estado
    ) + 1

    hoja.update_cell(
        fila_objetivo,
        columna,
        "GENERADO"
    )


def agregar_recordatorio_pab_a_cola(
    fila,
    vista_previa
):

    referencia = str(
        fila.get(
            "REFERENCIA",
            ""
        )
    ).strip()

    id_evento = str(
        fila.get(
            "ID_EVENTO",
            ""
        )
    ).strip()

    email = str(
        fila.get(
            "EMAIL",
            ""
        )
    ).strip()

    mora = fila.get(
        "MORA",
        ""
    )

    id_plantilla = str(
        vista_previa[
            "id_plantilla"
        ]
    ).strip().upper()

    if not referencia:
        raise ValueError(
            "El registro no tiene REFERENCIA."
        )

    if not id_evento:
        raise ValueError(
            "El registro no tiene ID_EVENTO."
        )

    if not email:
        raise ValueError(
            "El cliente no tiene correo."
        )

    if es_mora_180(
        mora
    ):
        raise ValueError(
            "Cliente excluido automáticamente por Mora 180."
        )

    referencias_excluidas = obtener_referencias_excluidas()

    if referencia in referencias_excluidas:
        raise ValueError(
            "Referencia bloqueada en Excluir_correo."
        )

    dias = fila.get(
        "_DIAS"
    )

    if (
        id_plantilla == "PAB000"
        and dias != 0
    ):
        raise ValueError(
            "PAB000 solo puede agregarse el día del pago."
        )

    if (
        id_plantilla == "PAB003"
        and dias != 3
    ):
        raise ValueError(
            "PAB003 solo puede agregarse exactamente 3 días antes."
        )

    fecha_pab = fila.get(
        "_FECHA"
    )

    if existe_envio_pab_en_cola(
        referencia,
        id_plantilla,
        fecha_pab
    ):
        raise ValueError(
            "Este recordatorio ya existe en COLA_ENVIO."
        )

    id_campana = (
        f"PAB-{HOY.strftime('%Y%m%d')}"
    )

    asegurar_campana_pab(
        id_campana
    )

    archivo = obtener_archivo()

    hoja_cola = archivo.worksheet(
        "COLA_ENVIO"
    )

    encabezados = obtener_encabezados_hoja(
        hoja_cola
    )

    id_envio = (
        f"ENV-PAB-"
        f"{HOY.strftime('%Y%m%d')}-"
        f"{referencia}-"
        f"{id_plantilla}"
    )

    datos = {
        "ID_ENVIO": id_envio,
        "ID_CAMPAÑA": id_campana,
        "REFERENCIA": referencia,
        "NOMBRE": str(
            fila.get(
                "NOMBRE",
                ""
            )
        ).strip(),
        "EMAIL": email,
        "PLANTILLA": id_plantilla,
        "ASUNTO": vista_previa.get(
            "asunto",
            ""
        ),
        "ESTADO": "PENDIENTE",
        "FECHA_PROG": AHORA.strftime(
            "%d/%m/%Y %H:%M"
        ),
        "FECHA_ENVIO": "",
        "INTENTOS": 0,
        "ERROR": "",
        "ID_MENSAJE": "",
        "CUERPO": vista_previa.get(
            "cuerpo",
            ""
        ),
        "ENCARGADO": str(
            fila.get(
                "ENCARGADO",
                ""
            )
        ).strip()
    }

    fila_nueva = construir_fila_por_encabezados(
        encabezados,
        datos
    )

    hoja_cola.append_row(
        fila_nueva,
        value_input_option="USER_ENTERED"
    )

    actualizar_estado_pab_proximos(
        id_evento,
        id_plantilla
    )

    actualizar_contadores_campana_pab(
        id_campana
    )

    st.cache_data.clear()

    return id_envio


# ============================================================
# CAMPAÑAS MANUALES / PREPARACIÓN SEGURA
# ============================================================

MAPA_PLANTILLAS_MORA = {
    "MORA_1": "T001",
    "MORA_30": "T030",
    "MORA_60": "T060",
    "MORA_90": "T090",
}

TIPOS_CAMPANA_MANUAL = [
    "MORA_1",
    "MORA_30",
    "MORA_60",
    "MORA_90",
    "PERSONALIZADA",
]


def obtener_ids_plantillas_activas():
    if plantillas.empty or "ID_PLANTILLA" not in plantillas.columns:
        return []

    vista = plantillas.copy()
    if "ESTADO" in vista.columns:
        activas = vista[
            vista["ESTADO"].astype(str).str.strip().str.upper() == "ACTIVA"
        ]
        if not activas.empty:
            vista = activas

    return sorted({
        str(x).strip()
        for x in vista["ID_PLANTILLA"].tolist()
        if str(x).strip()
    })


def obtener_plantilla_generica(id_plantilla):
    if plantillas.empty or "ID_PLANTILLA" not in plantillas.columns:
        return None

    candidatos = plantillas[
        plantillas["ID_PLANTILLA"].astype(str).str.strip().str.upper()
        == str(id_plantilla).strip().upper()
    ].copy()

    if candidatos.empty:
        return None

    if "ESTADO" in candidatos.columns:
        activas = candidatos[
            candidatos["ESTADO"].astype(str).str.strip().str.upper() == "ACTIVA"
        ]
        if not activas.empty:
            candidatos = activas

    return candidatos.iloc[0]


def reemplazar_variables_genericas(texto, fila):
    texto = str(texto or "")
    reemplazos = {
        "{{NOMBRE}}": str(fila.get("NOMBRE", "")).strip(),
        "{{REFERENCIA}}": str(fila.get("REFERENCIA", "")).strip(),
        "{{EMAIL}}": str(fila.get("EMAIL", "")).strip(),
        "{{MORA}}": str(fila.get("MORA", "")).strip(),
        "{{ENCARGADO}}": str(fila.get("ENCARGADO", "")).strip(),
        "{{SALDO}}": moneda(numero(fila.get("SALDO", 0))),
    }
    for variable, valor in reemplazos.items():
        texto = texto.replace(variable, str(valor))
    return texto


def crear_campana_manual(
    nombre,
    tipo,
    id_plantilla,
    fecha_envio,
    hora_envio,
    comentarios="",
    referencias_personalizadas=""
):
    nombre = str(nombre or "").strip()
    tipo = str(tipo or "").strip().upper()
    id_plantilla = str(id_plantilla or "").strip()
    referencias_personalizadas = str(referencias_personalizadas or "").strip()

    if not nombre:
        raise ValueError("Escribe un nombre para la campaña.")
    if tipo not in TIPOS_CAMPANA_MANUAL:
        raise ValueError("Tipo de campaña no válido.")
    if not id_plantilla:
        raise ValueError("La campaña debe tener una plantilla.")
    if obtener_plantilla_generica(id_plantilla) is None:
        raise ValueError(f"No encontré la plantilla {id_plantilla} en PLANTILLAS.")

    referencias_limpias = []
    if tipo == "PERSONALIZADA":
        partes = re.split(r"[\n,;\t ]+", referencias_personalizadas)
        vistas = set()
        for parte in partes:
            ref = normalizar_referencia(parte)
            if ref and ref not in vistas:
                vistas.add(ref)
                referencias_limpias.append(ref)
        if not referencias_limpias:
            raise ValueError(
                "En una campaña PERSONALIZADA debes pegar al menos una referencia."
            )

    archivo = obtener_archivo()
    hoja = archivo.worksheet("CAMPAÑAS")
    encabezados = obtener_encabezados_hoja(hoja)
    if "ID_CAMPAÑA" not in encabezados:
        raise ValueError("CAMPAÑAS no tiene la columna ID_CAMPAÑA.")

    # No modificamos encabezados ni estructura de CAMPAÑAS porque la hoja
    # puede tener rangos protegidos. Las referencias personalizadas se
    # conservan dentro de COMENTARIOS con una marca interna.

    ahora = datetime.now(TZ)
    base_id = f"MAN-{ahora.strftime('%Y%m%d-%H%M%S')}"
    id_campana = base_id

    existentes = set()
    valores = hoja.get_all_values()
    if len(valores) > 1:
        i_id = encabezados.index("ID_CAMPAÑA")
        existentes = {
            str(f[i_id]).strip()
            for f in valores[1:]
            if len(f) > i_id and str(f[i_id]).strip()
        }
    n = 2
    while id_campana in existentes:
        id_campana = f"{base_id}-{n}"
        n += 1

    datos = {
        "ID_CAMPAÑA": id_campana,
        "NOMBRE_CAMPAÑA": nombre,
        "PLANTILLA": id_plantilla,
        "FILTRO": tipo,
        "FECHA_ENVIO": fecha_envio.strftime("%d/%m/%Y"),
        "HORA_ENVIO": hora_envio.strftime("%H:%M"),
        "ESTADO": "BORRADOR",
        "TOTAL_CLIENTES": 0,
        "ENVIADOS": 0,
        "PENDIENTES": 0,
        "ERRORES": 0,
        "FECHA_CREACIÓN": ahora.strftime("%d/%m/%Y %H:%M:%S"),
        "COMENTARIOS": (
            (comentarios + "\n" if comentarios else "")
            + (f"[REFS_PERSONALIZADAS:{','.join(referencias_limpias)}]" if tipo == "PERSONALIZADA" else "")
        ).strip(),
    }
    hoja.append_row(
        construir_fila_por_encabezados(encabezados, datos),
        value_input_option="USER_ENTERED"
    )
    st.cache_data.clear()
    return id_campana


def fila_campana_por_id(id_campana):
    if campanas.empty or "ID_CAMPAÑA" not in campanas.columns:
        return None
    candidatos = campanas[
        campanas["ID_CAMPAÑA"].astype(str).str.strip() == str(id_campana).strip()
    ]
    if candidatos.empty:
        return None
    return candidatos.iloc[0]


def referencias_excluidas_normalizadas():
    return {
        normalizar_referencia(x)
        for x in obtener_referencias_excluidas()
        if normalizar_referencia(x)
    }


def preparar_clientes_campana(fila_campana):
    """Devuelve candidatos elegibles. No escribe ni envía nada."""
    resumen_vacio = {
        "total_base": 0,
        "elegibles": 0,
        "excluidos": 0,
        "mora_180": 0,
        "sin_email": 0,
        "duplicados": 0,
        "no_encontradas": 0,
    }
    if fila_campana is None:
        return pd.DataFrame(), resumen_vacio

    tipo = str(fila_campana.get("FILTRO", "")).strip().upper()
    id_plantilla = str(fila_campana.get("PLANTILLA", "")).strip()

    base = clientes.copy()

    # CLIENTES conserva los encabezados tal como están escritos en Sheets
    # (por ejemplo: Referencia, Nombre, Email, Mora). Pandas distingue
    # mayúsculas/minúsculas, por eso aquí los convertimos a nombres canónicos
    # sin exigir cambios en Google Sheets.
    def _clave_columna(nombre):
        texto = unicodedata.normalize("NFKD", str(nombre or ""))
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        return re.sub(r"[^A-Z0-9]+", "_", texto.upper()).strip("_")

    aliases = {
        "REFERENCIA": {"REFERENCIA", "REFERENCE", "REF"},
        "NOMBRE": {"NOMBRE", "NOMBRE_CLIENTE", "CLIENTE"},
        "EMAIL": {"EMAIL", "CORREO", "CORREO_ELECTRONICO", "E_MAIL"},
        "SALDO": {"SALDO", "SALDO_CLIENTE"},
        "MORA": {"MORA", "MORA_STATUS", "STATUS_MORA", "ESTADO_MORA"},
        "ENCARGADO": {"ENCARGADO", "PERSONA", "NEGOCIADOR", "RESPONSABLE"},
    }
    columnas_por_clave = {_clave_columna(c): c for c in base.columns}
    for canonica, posibles in aliases.items():
        if canonica in base.columns:
            continue
        origen = next((columnas_por_clave[p] for p in posibles if p in columnas_por_clave), None)
        base[canonica] = base[origen] if origen is not None else ""

    # Normalizamos primero para evitar que valores vacíos o formatos de Sheets
    # rompan la construcción de destinatarios.
    base["_REF"] = base["REFERENCIA"].apply(normalizar_referencia)
    base = base[base["_REF"] != ""].copy()

    no_encontradas = 0
    if tipo == "PERSONALIZADA":
        comentarios_guardados = str(fila_campana.get("COMENTARIOS", "") or "")
        match_refs = re.search(r"\[REFS_PERSONALIZADAS:([^\]]*)\]", comentarios_guardados)
        refs_txt = match_refs.group(1).strip() if match_refs else ""
        refs_solicitadas = []
        vistas = set()
        for parte in re.split(r"[\n,;\t ]+", refs_txt):
            ref = normalizar_referencia(parte)
            if ref and ref not in vistas:
                vistas.add(ref)
                refs_solicitadas.append(ref)

        if not refs_solicitadas:
            raise ValueError(
                "Esta campaña PERSONALIZADA no tiene referencias guardadas."
            )

        refs_en_base = set(base["_REF"].tolist())
        no_encontradas = len([r for r in refs_solicitadas if r not in refs_en_base])
        base = base[base["_REF"].isin(refs_solicitadas)].copy()

    elif tipo in MAPA_PLANTILLAS_MORA:
        objetivo = normalizar(tipo).replace("_", " ")
        base = base[
            base["MORA"].apply(
                lambda x: normalizar(x).replace("_", " ") == objetivo
            )
        ].copy()
    else:
        raise ValueError(f"Tipo de campaña no soportado: {tipo}")

    total_base = len(base)

    duplicados = entero_seguro(
        base.duplicated(subset=["_REF"], keep="first").sum()
    )
    base = base.drop_duplicates(subset=["_REF"], keep="first").copy()

    mask_180 = base["MORA"].apply(es_mora_180)
    n_mora_180 = entero_seguro(mask_180.sum())
    base = base[~mask_180].copy()

    excluidas = referencias_excluidas_normalizadas()
    mask_excl = base["_REF"].isin(excluidas)
    n_excl = entero_seguro(mask_excl.sum())
    base = base[~mask_excl].copy()

    mask_sin_email = base["EMAIL"].apply(valor_vacio)
    n_sin_email = entero_seguro(mask_sin_email.sum())
    base = base[~mask_sin_email].copy()

    plantilla = obtener_plantilla_generica(id_plantilla)
    if plantilla is None:
        raise ValueError(f"No encontré la plantilla {id_plantilla}.")

    base["ASUNTO_PREVIO"] = base.apply(
        lambda f: reemplazar_variables_genericas(plantilla.get("ASUNTO", ""), f),
        axis=1
    )
    base["CUERPO_PREVIO"] = base.apply(
        lambda f: reemplazar_variables_genericas(plantilla.get("CUERPO", ""), f),
        axis=1
    )

    return base, {
        "total_base": total_base,
        "elegibles": len(base),
        "excluidos": n_excl,
        "mora_180": n_mora_180,
        "sin_email": n_sin_email,
        "duplicados": duplicados,
        "no_encontradas": no_encontradas,
    }


def ids_envio_existentes():
    if cola.empty or "ID_ENVIO" not in cola.columns:
        return set()
    return {
        str(x).strip()
        for x in cola["ID_ENVIO"].tolist()
        if str(x).strip()
    }


def preparar_campana_en_cola(id_campana, df_destinatarios):
    """Escribe en COLA_ENVIO como BORRADOR. Nunca envía correos."""
    if df_destinatarios.empty:
        raise ValueError("No hay destinatarios elegibles para preparar.")

    fila_camp = fila_campana_por_id(id_campana)
    if fila_camp is None:
        raise ValueError("No encontré la campaña seleccionada.")

    estado = str(fila_camp.get("ESTADO", "")).strip().upper()
    if estado not in {"BORRADOR", "PREPARADA"}:
        raise ValueError(f"La campaña está en estado {estado} y no puede prepararse.")

    archivo = obtener_archivo()
    hoja_cola = archivo.worksheet("COLA_ENVIO")
    encabezados_cola = obtener_encabezados_hoja(hoja_cola)
    existentes = ids_envio_existentes()

    fecha_txt = str(fila_camp.get("FECHA_ENVIO", "")).strip()
    hora_txt = str(fila_camp.get("HORA_ENVIO", "")).strip()
    fecha_prog = f"{fecha_txt} {hora_txt}".strip()
    id_plantilla = str(fila_camp.get("PLANTILLA", "")).strip()

    filas = []
    agregados = 0
    omitidos = 0

    for _, f in df_destinatarios.iterrows():
        ref = normalizar_referencia(f.get("REFERENCIA", ""))
        if not ref:
            omitidos += 1
            continue

        id_envio = f"ENV-{id_campana}-{ref}"
        if id_envio in existentes:
            omitidos += 1
            continue

        datos = {
            "ID_ENVIO": id_envio,
            "ID_CAMPAÑA": id_campana,
            "REFERENCIA": ref,
            "NOMBRE": str(f.get("NOMBRE", "")).strip(),
            "EMAIL": str(f.get("EMAIL", "")).strip(),
            "PLANTILLA": id_plantilla,
            "ASUNTO": str(f.get("ASUNTO_PREVIO", "")).strip(),
            "ESTADO": "BORRADOR",
            "FECHA_PROG": fecha_prog,
            "FECHA_ENVIO": "",
            "INTENTOS": 0,
            "ERROR": "",
            "ID_MENSAJE": "",
            "CUERPO": str(f.get("CUERPO_PREVIO", "")).strip(),
            "ENCARGADO": str(f.get("ENCARGADO", "")).strip(),
        }
        filas.append(construir_fila_por_encabezados(encabezados_cola, datos))
        existentes.add(id_envio)
        agregados += 1

    if filas:
        hoja_cola.append_rows(filas, value_input_option="USER_ENTERED")

    # Actualizar campaña: sigue sin ser enviable.
    hoja_camp = archivo.worksheet("CAMPAÑAS")
    valores_camp = hoja_camp.get_all_values()
    enc_camp = [str(x).strip() for x in valores_camp[0]]
    i_id = enc_camp.index("ID_CAMPAÑA")
    fila_obj = None
    for n, f in enumerate(valores_camp[1:], start=2):
        if len(f) > i_id and str(f[i_id]).strip() == id_campana:
            fila_obj = n
            break

    if fila_obj:
        cambios = {
            "ESTADO": "PREPARADA",
            "TOTAL_CLIENTES": agregados,
            "PENDIENTES": agregados,
            "ENVIADOS": 0,
            "ERRORES": 0,
        }
        for enc, val in cambios.items():
            if enc in enc_camp:
                hoja_camp.update_cell(fila_obj, enc_camp.index(enc) + 1, val)

    st.cache_data.clear()
    return agregados, omitidos


def _buscar_fila_campana_en_sheet(hoja_camp, id_campana):
    """Devuelve (numero_fila, encabezados) para una campaña."""
    valores = hoja_camp.get_all_values()
    if not valores:
        raise ValueError("La hoja CAMPAÑAS está vacía.")
    encabezados = [str(x).strip() for x in valores[0]]
    if "ID_CAMPAÑA" not in encabezados:
        raise ValueError("CAMPAÑAS no tiene la columna ID_CAMPAÑA.")
    i_id = encabezados.index("ID_CAMPAÑA")
    for numero_fila, fila in enumerate(valores[1:], start=2):
        if len(fila) > i_id and str(fila[i_id]).strip() == str(id_campana).strip():
            return numero_fila, encabezados
    raise ValueError("No encontré la campaña seleccionada en CAMPAÑAS.")


def _actualizar_campos_campana(id_campana, cambios):
    archivo = obtener_archivo()
    hoja_camp = archivo.worksheet("CAMPAÑAS")
    fila_obj, encabezados = _buscar_fila_campana_en_sheet(hoja_camp, id_campana)
    for encabezado, valor in cambios.items():
        if encabezado in encabezados:
            hoja_camp.update_cell(fila_obj, encabezados.index(encabezado) + 1, valor)


def programar_campana_segura(id_campana):
    """
    Marca la CAMPAÑA como PROGRAMADA, pero mantiene TODOS sus correos en
    COLA_ENVIO como BORRADOR. Por diseño, esta función NO habilita envíos.
    """
    fila = fila_campana_por_id(id_campana)
    if fila is None:
        raise ValueError("No encontré la campaña seleccionada.")
    estado = str(fila.get("ESTADO", "")).strip().upper()
    if estado != "PREPARADA":
        raise ValueError("Solo una campaña PREPARADA puede programarse.")

    # Verificar que existan correos preparados y que ninguno deje BORRADOR.
    archivo = obtener_archivo()
    hoja_cola = archivo.worksheet("COLA_ENVIO")
    valores = hoja_cola.get_all_values()
    if not valores:
        raise ValueError("COLA_ENVIO está vacía.")
    encabezados = [str(x).strip() for x in valores[0]]
    if "ID_CAMPAÑA" not in encabezados or "ESTADO" not in encabezados:
        raise ValueError("COLA_ENVIO no tiene ID_CAMPAÑA o ESTADO.")
    i_camp = encabezados.index("ID_CAMPAÑA")
    i_estado = encabezados.index("ESTADO")
    filas_camp = [
        f for f in valores[1:]
        if len(f) > i_camp and str(f[i_camp]).strip() == str(id_campana).strip()
    ]
    if not filas_camp:
        raise ValueError("La campaña no tiene correos preparados en COLA_ENVIO.")
    estados_no_seguros = {
        str(f[i_estado]).strip().upper()
        for f in filas_camp if len(f) > i_estado
    } - {"BORRADOR"}
    if estados_no_seguros:
        raise ValueError(
            "Hay correos de esta campaña fuera de BORRADOR: "
            + ", ".join(sorted(estados_no_seguros))
        )

    _actualizar_campos_campana(id_campana, {"ESTADO": "PROGRAMADA"})
    st.cache_data.clear()
    return len(filas_camp)


def cancelar_preparacion_campana(id_campana):
    """
    Elimina de COLA_ENVIO únicamente los BORRADORES de la campaña y devuelve
    CAMPAÑAS a BORRADOR. Nunca toca filas enviadas o habilitadas.
    """
    fila = fila_campana_por_id(id_campana)
    if fila is None:
        raise ValueError("No encontré la campaña seleccionada.")
    estado = str(fila.get("ESTADO", "")).strip().upper()
    if estado not in {"PREPARADA", "PROGRAMADA"}:
        raise ValueError("Solo se puede cancelar una campaña PREPARADA o PROGRAMADA.")

    archivo = obtener_archivo()
    hoja_cola = archivo.worksheet("COLA_ENVIO")
    valores = hoja_cola.get_all_values()
    eliminadas = 0

    if valores:
        encabezados = [str(x).strip() for x in valores[0]]
        if "ID_CAMPAÑA" not in encabezados or "ESTADO" not in encabezados:
            raise ValueError("COLA_ENVIO no tiene ID_CAMPAÑA o ESTADO.")
        i_camp = encabezados.index("ID_CAMPAÑA")
        i_estado = encabezados.index("ESTADO")
        filas_borrar = []
        estados_bloqueantes = set()
        for numero_fila, f in enumerate(valores[1:], start=2):
            if len(f) <= i_camp or str(f[i_camp]).strip() != str(id_campana).strip():
                continue
            est = str(f[i_estado]).strip().upper() if len(f) > i_estado else ""
            if est == "BORRADOR":
                filas_borrar.append(numero_fila)
            else:
                estados_bloqueantes.add(est or "VACÍO")

        if estados_bloqueantes:
            raise ValueError(
                "No puedo cancelar porque existen correos fuera de BORRADOR: "
                + ", ".join(sorted(estados_bloqueantes))
            )

        # Borrar de abajo hacia arriba para no desplazar números de fila.
        for numero_fila in reversed(filas_borrar):
            hoja_cola.delete_rows(numero_fila)
            eliminadas += 1

    _actualizar_campos_campana(
        id_campana,
        {
            "ESTADO": "BORRADOR",
            "TOTAL_CLIENTES": 0,
            "PENDIENTES": 0,
            "ENVIADOS": 0,
            "ERRORES": 0,
        }
    )
    st.cache_data.clear()
    return eliminadas


# ============================================================
# ESCRITURA RESPUESTAS
# ============================================================

def buscar_fila_respuesta_por_id(
    hoja,
    id_respuesta
):

    valores = hoja.col_values(1)

    id_buscado = str(
        id_respuesta
    ).strip()

    for numero_fila, valor in enumerate(
        valores,
        start=1
    ):

        if str(valor).strip() == id_buscado:

            return numero_fila

    return None


def guardar_comentario(
    id_respuesta,
    comentario
):

    archivo = obtener_archivo()

    hoja = archivo.worksheet(
        "RESPUESTAS"
    )

    fila = buscar_fila_respuesta_por_id(
        hoja,
        id_respuesta
    )

    if not fila:

        raise ValueError(
            f"No encontré ID_RESPUESTA: {id_respuesta}"
        )

    # P = COMENTARIOS
    hoja.update_cell(
        fila,
        16,
        comentario
    )

    st.cache_data.clear()


def marcar_respuesta_gestionada(
    id_respuesta,
    comentario
):

    archivo = obtener_archivo()

    hoja = archivo.worksheet(
        "RESPUESTAS"
    )

    fila = buscar_fila_respuesta_por_id(
        hoja,
        id_respuesta
    )

    if not fila:

        raise ValueError(
            f"No encontré ID_RESPUESTA: {id_respuesta}"
        )

    # O = CHEK
    hoja.update_cell(
        fila,
        15,
        "TRUE"
    )

    # P = COMENTARIOS
    hoja.update_cell(
        fila,
        16,
        comentario
    )

    st.cache_data.clear()


# ============================================================
# CARGAR DATOS
# ============================================================

try:

    clientes = cargar_hoja(
        "CLIENTES"
    )

    campanas = cargar_hoja(
        "CAMPAÑAS"
    )

    cola = cargar_hoja(
        "COLA_ENVIO"
    )

    respuestas = cargar_hoja(
        "RESPUESTAS"
    )

    pab = cargar_hoja(
        "PAB_PROXIMOS"
    )

    plantillas = cargar_hoja(
        "PLANTILLAS"
    )

except Exception as e:

    st.error(
        "No pude conectar con Google Sheets."
    )

    st.exception(e)

    st.stop()


# ============================================================
# COMPLETAR NOMBRE / EMAIL DE PAB DESDE CARTERA BEREX
# ============================================================

pab, info_enriquecimiento_pab = (
    enriquecer_pab_con_cartera_berex(
        pab
    )
)


# ============================================================
# PREPARAR RESPUESTAS
# ============================================================

if "CHEK" in respuestas.columns:

    respuestas["_ATENDIDA"] = (
        respuestas["CHEK"]
        .map(es_true)
    )

else:

    respuestas["_ATENDIDA"] = False


if "FECHA_RESPUESTA" in respuestas.columns:

    respuestas["_FECHA"] = convertir_fechas(
        respuestas["FECHA_RESPUESTA"]
    )

else:

    respuestas["_FECHA"] = pd.NaT


limite_24h = (
    AHORA.replace(tzinfo=None)
    - timedelta(hours=24)
)


respuestas["_MAS_24H"] = (
    respuestas["_FECHA"].notna()
    &
    (
        respuestas["_FECHA"]
        <= limite_24h
    )
    &
    (
        ~respuestas["_ATENDIDA"]
    )
)


def texto_tiempo_pendiente(fecha):
    if pd.isna(fecha):
        return "Sin fecha"
    try:
        delta = AHORA.replace(tzinfo=None) - fecha.to_pydatetime()
        horas = max(0, int(delta.total_seconds() // 3600))
        if horas < 24:
            return f"{horas} h"
        dias = horas // 24
        resto = horas % 24
        return f"{dias} d {resto} h"
    except Exception:
        return "Sin fecha"

respuestas["_TIEMPO_PENDIENTE"] = respuestas["_FECHA"].apply(texto_tiempo_pendiente)


total_respuestas = len(
    respuestas
)

total_atendidas = int(
    respuestas["_ATENDIDA"].sum()
)

total_pendientes = int(
    (~respuestas["_ATENDIDA"]).sum()
)

pendientes_24 = int(
    respuestas["_MAS_24H"].sum()
)


# ============================================================
# PREPARAR PAB
# ============================================================

if "FECHA_PAB" in pab.columns:

    pab["_FECHA"] = convertir_fechas(
        pab["FECHA_PAB"]
    )

else:

    pab["_FECHA"] = pd.NaT


if "VALOR_PAB" in pab.columns:

    pab["_VALOR"] = (
        pab["VALOR_PAB"]
        .apply(numero)
    )

else:

    pab["_VALOR"] = 0.0


if "AVISO_3_DIAS" in pab.columns:

    pab["_AVISO_3"] = (
        pab["AVISO_3_DIAS"]
        .apply(aviso_generado)
    )

else:

    pab["_AVISO_3"] = False


if "AVISO_HOY" in pab.columns:

    pab["_AVISO_HOY"] = (
        pab["AVISO_HOY"]
        .apply(aviso_generado)
    )

else:

    pab["_AVISO_HOY"] = False


pab["_DIAS"] = (
    pab["_FECHA"].dt.date
    .apply(
        lambda x:
        (x - HOY).days
        if pd.notna(x)
        else None
    )
)


pab_hoy_mask = (
    pab["_DIAS"] == 0
)

pab_3_dias_mask = (
    pab["_DIAS"].notna()
    &
    (
        pab["_DIAS"] >= 0
    )
    &
    (
        pab["_DIAS"] <= 3
    )
)


total_pab_hoy = int(
    pab_hoy_mask.sum()
)


total_pab_3_dias = int(
    pab_3_dias_mask.sum()
)


valor_pab_hoy = float(
    pab.loc[
        pab_hoy_mask,
        "_VALOR"
    ].sum()
)


pendientes_aviso_hoy = int(
    (
        pab_hoy_mask
        &
        (~pab["_AVISO_HOY"])
    ).sum()
)


# Para recordatorio 3 días:
# únicamente pagos que son exactamente dentro de 3 días

pab_exactamente_3_mask = (
    pab["_DIAS"] == 3
)

pendientes_aviso_3 = int(
    (
        pab_exactamente_3_mask
        &
        (~pab["_AVISO_3"])
    ).sum()
)


total_recordatorios_pendientes = (
    pendientes_aviso_hoy
    +
    pendientes_aviso_3
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        "## Masivos Correos"
    )

    st.caption(
        "Crecemos juntos"
    )

    st.markdown("---")

    menu = st.radio(
        "Menú",
        [
            "🏠 Inicio",
            "📧 Campañas",
            "🏦 Pagos a Banco",
            "💬 Respuestas",
            "⚠️ Pendientes",
            "📝 Plantillas",
            "🕘 Historial",
            "⚙️ Configuración"
        ],
        label_visibility="collapsed"
    )

    st.markdown("---")

    if st.button(
        "🔄 Actualizar datos",
        use_container_width=True
    ):

        st.cache_data.clear()

        st.rerun()

    st.caption(
        AHORA.strftime(
            "Actualizado: %d/%m/%Y %I:%M %p"
        )
    )


# ============================================================
# INICIO
# ============================================================

if menu == "🏠 Inicio":

    st.markdown(
        '<div class="titulo">'
        'Masivos Bravo'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitulo">'
        'Gestión de campañas, respuestas y comunicaciones con clientes'
        '</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # KPIs OPERATIVOS DE HOY
    # --------------------------------------------------------

    enviados_hoy = 0
    errores_cola = 0

    if not cola.empty:
        if "ESTADO" in cola.columns:
            estados_cola = cola["ESTADO"].astype(str).str.strip().str.upper()
            errores_cola = int(estados_cola.isin(["ERROR", "BLOQUEADO"]).sum())
        else:
            estados_cola = pd.Series("", index=cola.index)

        if "FECHA_ENVIO" in cola.columns:
            fechas_envio = convertir_fechas(cola["FECHA_ENVIO"])
            enviados_hoy = int(
                (
                    (estados_cola == "ENVIADO")
                    &
                    (fechas_envio.dt.date == HOY)
                ).sum()
            )

    respuestas_hoy = 0
    if not respuestas.empty and "_FECHA" in respuestas.columns:
        respuestas_hoy = int(
            (respuestas["_FECHA"].dt.date == HOY).sum()
        )

    campanas_activas = 0
    if not campanas.empty and "ESTADO" in campanas.columns:
        campanas_activas = int(
            campanas["ESTADO"]
            .astype(str)
            .str.strip()
            .str.upper()
            .isin(["BORRADOR", "PROGRAMADA", "PENDIENTE", "EN PROCESO"])
            .sum()
        )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "📨 Correos enviados hoy",
        enviados_hoy
    )

    c2.metric(
        "💬 Respuestas nuevas hoy",
        respuestas_hoy
    )

    c3.metric(
        "🔴 Pendientes +24h",
        pendientes_24
    )

    c4.metric(
        "🏦 PaB hoy",
        total_pab_hoy
    )

    st.markdown("---")

    p1, p2, p3, p4 = st.columns(4)

    p1.metric(
        "📅 PaB próximos 3 días",
        total_pab_3_dias
    )

    p2.metric(
        "💰 Valor PaB hoy",
        moneda(valor_pab_hoy)
    )

    p3.metric(
        "⏳ Recordatorios PaB pendientes",
        total_recordatorios_pendientes
    )

    p4.metric(
        "📧 Campañas abiertas",
        campanas_activas
    )

    # --------------------------------------------------------
    # ACCIONES RÁPIDAS
    # --------------------------------------------------------

    st.markdown("---")
    st.subheader("⚡ Acciones rápidas")
    st.caption("Accede a las funciones que más usa el equipo.")

    a1, a2, a3, a4 = st.columns(4)

    with a1:
        if st.button(
            "➕ Nueva campaña",
            use_container_width=True,
            type="primary"
        ):
            st.session_state["navegar_a"] = "📧 Campañas"
            st.session_state["abrir_nueva_campana"] = True
            st.rerun()

    with a2:
        if st.button(
            "💬 Ver respuestas",
            use_container_width=True
        ):
            st.session_state["navegar_a"] = "💬 Respuestas"
            st.rerun()

    with a3:
        if st.button(
            "🏦 Ver pagos a banco",
            use_container_width=True
        ):
            st.session_state["navegar_a"] = "🏦 Pagos a Banco"
            st.rerun()

    with a4:
        if st.button(
            "⚠️ Ver pendientes",
            use_container_width=True
        ):
            st.session_state["navegar_a"] = "⚠️ Pendientes"
            st.rerun()

    st.markdown("---")

    izquierda, derecha = st.columns([1.65, 1])

    with izquierda:
        st.subheader("🕘 Actividad reciente")

        if cola.empty:
            st.info("Todavía no hay actividad registrada en COLA_ENVIO.")
        else:
            actividad = cola.copy().tail(10).iloc[::-1]

            columnas_actividad = [
                c
                for c in [
                    "FECHA_ENVIO",
                    "FECHA_PROG",
                    "REFERENCIA",
                    "PLANTILLA",
                    "ESTADO",
                    "ENCARGADO"
                ]
                if c in actividad.columns
            ]

            st.dataframe(
                actividad[columnas_actividad],
                use_container_width=True,
                hide_index=True,
                height=330
            )

    with derecha:
        st.subheader("⚙️ Estado operativo")

        st.success(
            "🏦 **Pagos a Banco**\n\n"
            "Automatización diaria en preparación para las 8:00 a. m."
        )

        st.info(
            "📧 **Campañas de mora**\n\n"
            "Mora 1, 30, 60 y 90 se crean y programan manualmente por el equipo."
        )

        if errores_cola:
            st.warning(
                f"⚠️ Hay {errores_cola} registros con ERROR o BLOQUEADO en COLA_ENVIO."
            )
        else:
            st.success("✅ Sin errores activos detectados en la cola.")

# ============================================================
# PAGOS A BANCO
# ============================================================

elif menu == "🏦 Pagos a Banco":

    st.markdown(
        '<div class="titulo">'
        '🏦 Pagos a Banco'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitulo">'
        'Seguimiento de próximos pagos y recordatorios a clientes'
        '</div>',
        unsafe_allow_html=True
    )

    error_info_v2 = info_enriquecimiento_pab.get(
        "error_info_v2"
    )

    if error_info_v2:
        st.warning(
            "⚠️ No pude consultar Info_Clientes_V2: "
            + error_info_v2
        )

    nv2 = info_enriquecimiento_pab.get(
        "nombres_completados_info_v2",
        0
    )
    ev2 = info_enriquecimiento_pab.get(
        "emails_completados_info_v2",
        0
    )

    if nv2 or ev2:
        st.success(
            "✅ Datos de clientes completados: "
            f"{nv2} nombres / {ev2} correos"
        )

    faltan_nombre = info_enriquecimiento_pab.get(
        "sin_nombre",
        0
    )
    faltan_email = info_enriquecimiento_pab.get(
        "sin_email",
        0
    )

    if faltan_nombre or faltan_email:
        st.caption(
            f"ℹ️ Aún faltan {faltan_nombre} nombres y "
            f"{faltan_email} correos después de consultar Info_Clientes_V2."
        )

    # --------------------------------------------------------
    # MÉTRICAS
    # --------------------------------------------------------

    m1, m2, m3, m4 = st.columns(
        4
    )

    m1.metric(
        "🏦 Pagos hoy",
        total_pab_hoy
    )

    m2.metric(
        "📅 Hoy + próximos 3 días",
        total_pab_3_dias
    )

    m3.metric(
        "💰 Valor pagos hoy",
        moneda(
            valor_pab_hoy
        )
    )

    m4.metric(
        "⚠️ Avisos pendientes",
        total_recordatorios_pendientes
    )

    st.markdown("---")

    # --------------------------------------------------------
    # FILTROS
    # --------------------------------------------------------

    st.subheader(
        "🔎 Buscar y filtrar"
    )

    f1, f2, f3 = st.columns(
        [1.5, 1, 1]
    )

    with f1:

        buscar_pab = st.text_input(
            "Buscar",
            placeholder=(
                "Referencia, deuda, nombre, correo..."
            ),
            key="buscar_pab"
        )

    with f2:

        encargados = []

        if "ENCARGADO" in pab.columns:

            encargados = sorted(
                [
                    str(x).strip()
                    for x in pab[
                        "ENCARGADO"
                    ].unique()
                    if str(x).strip()
                ]
            )

        encargado_pab = st.selectbox(
            "Encargado",
            ["Todos"] + encargados,
            key="encargado_pab"
        )

    with f3:

        negociadores = []

        if "NEGOCIADOR" in pab.columns:

            negociadores = sorted(
                [
                    str(x).strip()
                    for x in pab[
                        "NEGOCIADOR"
                    ].unique()
                    if str(x).strip()
                ]
            )

        negociador_pab = st.selectbox(
            "Negociador",
            ["Todos"] + negociadores,
            key="negociador_pab"
        )

    f4, f5, f6 = st.columns(
        3
    )

    with f4:

        semaforos = []

        if "SEMAFORO" in pab.columns:

            semaforos = sorted(
                [
                    str(x).strip()
                    for x in pab[
                        "SEMAFORO"
                    ].unique()
                    if str(x).strip()
                ]
            )

        semaforo_pab = st.selectbox(
            "Semáforo",
            ["Todos"] + semaforos,
            key="semaforo_pab"
        )

    with f5:

        ventana_pab = st.selectbox(
            "Fecha",
            [
                "Próximos pagos",
                "Hoy",
                "Próximos 3 días",
                "Este mes",
                "Todos"
            ],
            key="ventana_pab"
        )

    with f6:

        aviso_pab = st.selectbox(
            "Estado recordatorio",
            [
                "Todos",
                "Aviso pendiente",
                "Aviso generado"
            ],
            key="aviso_pab"
        )

    # --------------------------------------------------------
    # FILTRAR
    # --------------------------------------------------------

    vista_pab = pab.copy()


    if buscar_pab.strip():

        texto = (
            buscar_pab
            .strip()
            .lower()
        )

        columnas_busqueda = [
            c
            for c in [
                "REFERENCIA",
                "ID_DEUDA",
                "NOMBRE",
                "EMAIL",
                "NEGOCIADOR",
                "ENCARGADO"
            ]
            if c in vista_pab.columns
        ]

        mascara = pd.Series(
            False,
            index=vista_pab.index
        )

        for col in columnas_busqueda:

            mascara |= (
                vista_pab[col]
                .astype(str)
                .str.lower()
                .str.contains(
                    texto,
                    regex=False,
                    na=False
                )
            )

        vista_pab = vista_pab[
            mascara
        ]


    if (
        encargado_pab != "Todos"
        and
        "ENCARGADO" in vista_pab.columns
    ):

        vista_pab = vista_pab[
            vista_pab["ENCARGADO"]
            == encargado_pab
        ]


    if (
        negociador_pab != "Todos"
        and
        "NEGOCIADOR" in vista_pab.columns
    ):

        vista_pab = vista_pab[
            vista_pab["NEGOCIADOR"]
            == negociador_pab
        ]


    if (
        semaforo_pab != "Todos"
        and
        "SEMAFORO" in vista_pab.columns
    ):

        vista_pab = vista_pab[
            vista_pab["SEMAFORO"]
            == semaforo_pab
        ]


    if ventana_pab == "Próximos pagos":

        vista_pab = vista_pab[
            vista_pab["_DIAS"].notna()
            &
            (
                vista_pab["_DIAS"]
                >= 0
            )
        ]


    elif ventana_pab == "Hoy":

        vista_pab = vista_pab[
            vista_pab["_DIAS"]
            == 0
        ]


    elif ventana_pab == "Próximos 3 días":

        vista_pab = vista_pab[
            vista_pab["_DIAS"].notna()
            &
            (
                vista_pab["_DIAS"]
                >= 0
            )
            &
            (
                vista_pab["_DIAS"]
                <= 3
            )
        ]


    elif ventana_pab == "Este mes":

        vista_pab = vista_pab[
            vista_pab["_FECHA"].notna()
            &
            (
                vista_pab["_FECHA"].dt.month
                == HOY.month
            )
            &
            (
                vista_pab["_FECHA"].dt.year
                == HOY.year
            )
        ]


    if aviso_pab != "Todos":

        def tiene_aviso_correcto(fila):

            dias = fila["_DIAS"]

            if dias == 0:
                return bool(
                    fila["_AVISO_HOY"]
                )

            if dias == 3:
                return bool(
                    fila["_AVISO_3"]
                )

            return False


        estado_aviso_fila = (
            vista_pab.apply(
                tiene_aviso_correcto,
                axis=1
            )
        )


        if aviso_pab == "Aviso generado":

            vista_pab = vista_pab[
                estado_aviso_fila
            ]


        elif aviso_pab == "Aviso pendiente":

            vista_pab = vista_pab[
                ~estado_aviso_fila
            ]


    vista_pab = vista_pab.sort_values(
        "_FECHA",
        ascending=True,
        na_position="last"
    )

    st.markdown("---")

    # --------------------------------------------------------
    # RESULTADOS
    # --------------------------------------------------------

    st.write(
        f"### 📋 {len(vista_pab)} pagos encontrados"
    )

    if vista_pab.empty:

        st.info(
            "No hay pagos que cumplan "
            "los filtros seleccionados."
        )

    else:

        # Tabla resumen
        columnas_tabla = [
            c
            for c in [
                "REFERENCIA",
                "ID_DEUDA",
                "NOMBRE",
                "FECHA_PAB",
                "VALOR_PAB",
                "NEGOCIADOR",
                "ENCARGADO",
                "SEMAFORO",
                "N_PAGO_ACTUAL",
                "N_PAGOS_RESTANTES",
                "AVISO_3_DIAS",
                "AVISO_HOY"
            ]
            if c in vista_pab.columns
        ]

        st.dataframe(
            vista_pab[
                columnas_tabla
            ],
            use_container_width=True,
            hide_index=True,
            height=360
        )

        st.markdown("---")

        st.subheader(
            "🔍 Detalle de pagos"
        )

        cantidad_pab = st.selectbox(
            "Mostrar detalle de",
            [
                10,
                20,
                50
            ],
            index=0
        )


        for idx, fila in (
            vista_pab
            .head(cantidad_pab)
            .iterrows()
        ):

            nombre = str(
                fila.get(
                    "NOMBRE",
                    "Cliente"
                )
            ).strip()

            referencia = str(
                fila.get(
                    "REFERENCIA",
                    ""
                )
            ).strip()

            fecha_pab = str(
                fila.get(
                    "FECHA_PAB",
                    ""
                )
            ).strip()

            valor_pab = str(
                fila.get(
                    "VALOR_PAB",
                    ""
                )
            ).strip()

            encargado = str(
                fila.get(
                    "ENCARGADO",
                    ""
                )
            ).strip()

            negociador = str(
                fila.get(
                    "NEGOCIADOR",
                    ""
                )
            ).strip()

            semaforo = str(
                fila.get(
                    "SEMAFORO",
                    ""
                )
            ).strip()

            dias = fila.get(
                "_DIAS"
            )

            aviso_3 = bool(
                fila.get(
                    "_AVISO_3",
                    False
                )
            )

            aviso_hoy_fila = bool(
                fila.get(
                    "_AVISO_HOY",
                    False
                )
            )


            if dias == 0:

                estado_fecha = (
                    "🔴 HOY"
                )

            elif dias == 1:

                estado_fecha = (
                    "🟠 Mañana"
                )

            elif dias == 3:

                estado_fecha = (
                    "🟡 En 3 días"
                )

            elif (
                dias is not None
                and
                dias > 0
            ):

                estado_fecha = (
                    f"🟢 En {int(dias)} días"
                )

            elif (
                dias is not None
                and
                dias < 0
            ):

                estado_fecha = (
                    f"⚪ Vencido hace "
                    f"{abs(int(dias))} días"
                )

            else:

                estado_fecha = (
                    "Fecha no disponible"
                )


            titulo = (
                f"{estado_fecha} · "
                f"{nombre} · "
                f"{referencia}"
            )


            with st.expander(
                titulo,
                expanded=False
            ):

                d1, d2, d3, d4 = st.columns(
                    4
                )

                d1.write(
                    "**Fecha PaB**"
                )

                d1.write(
                    fecha_pab
                )

                d2.write(
                    "**Valor PaB**"
                )

                d2.write(
                    valor_pab
                )

                d3.write(
                    "**Encargado**"
                )

                d3.write(
                    encargado
                )

                d4.write(
                    "**Semáforo**"
                )

                d4.write(
                    semaforo
                )

                st.markdown("---")

                e1, e2, e3 = st.columns(
                    3
                )

                e1.write(
                    "**Negociador**"
                )

                e1.write(
                    negociador
                )

                e2.write(
                    "**Pago actual**"
                )

                e2.write(
                    fila.get(
                        "N_PAGO_ACTUAL",
                        ""
                    )
                )

                e3.write(
                    "**Pagos restantes**"
                )

                e3.write(
                    fila.get(
                        "N_PAGOS_RESTANTES",
                        ""
                    )
                )

                st.markdown("---")

                a1, a2 = st.columns(
                    2
                )

                if aviso_3:

                    a1.success(
                        "✅ Recordatorio 3 días generado"
                    )

                else:

                    a1.warning(
                        "⏳ Recordatorio 3 días pendiente"
                    )


                if aviso_hoy_fila:

                    a2.success(
                        "✅ Recordatorio del día generado"
                    )

                else:

                    a2.warning(
                        "⏳ Recordatorio del día pendiente"
                    )

                st.markdown("---")

                vista_previa = preparar_vista_previa_pab(
                    fila
                )

                if vista_previa is None:

                    st.info(
                        "ℹ️ Este pago no requiere un recordatorio "
                        "hoy. La vista previa aparecerá cuando falten "
                        "exactamente 3 días o cuando sea la fecha del pago."
                    )

                elif vista_previa.get("error"):

                    st.error(
                        vista_previa["error"]
                    )

                else:

                    st.markdown(
                        "#### ✉️ Vista previa del correo"
                    )

                    st.caption(
                        f"Plantilla: {vista_previa['id_plantilla']} · "
                        f"{vista_previa['tipo_aviso']}"
                    )

                    email_cliente = str(
                        fila.get(
                            "EMAIL",
                            ""
                        )
                    ).strip()

                    if email_cliente:

                        st.write(
                            f"**Para:** {email_cliente}"
                        )

                    else:

                        st.warning(
                            "⚠️ Este registro no tiene correo del cliente."
                        )

                    st.write(
                        "**Asunto**"
                    )

                    st.code(
                        vista_previa["asunto"],
                        language=None
                    )

                    st.write(
                        "**Cuerpo**"
                    )

                    st.markdown(
                        vista_previa["cuerpo"]
                    )

                    st.info(
                        "🔒 Este botón solo agrega el recordatorio a "
                        "COLA_ENVIO. No envía el correo directamente."
                    )

                    puede_agregar = True
                    motivo_bloqueo = ""

                    if not email_cliente:

                        puede_agregar = False
                        motivo_bloqueo = (
                            "El cliente no tiene correo."
                        )

                    elif es_mora_180(
                        fila.get(
                            "MORA",
                            ""
                        )
                    ):

                        puede_agregar = False
                        motivo_bloqueo = (
                            "Cliente excluido por Mora 180."
                        )

                    elif (
                        vista_previa["id_plantilla"] == "PAB000"
                        and fila.get("_AVISO_HOY", False)
                    ):

                        puede_agregar = False
                        motivo_bloqueo = (
                            "El recordatorio de hoy ya figura como generado."
                        )

                    elif (
                        vista_previa["id_plantilla"] == "PAB003"
                        and fila.get("_AVISO_3", False)
                    ):

                        puede_agregar = False
                        motivo_bloqueo = (
                            "El recordatorio de 3 días ya figura como generado."
                        )

                    if not puede_agregar:

                        st.warning(
                            f"⚠️ {motivo_bloqueo}"
                        )

                    confirmar = st.checkbox(
                        "Confirmo que revisé la vista previa y quiero "
                        "agregar este recordatorio a COLA_ENVIO.",
                        key=f"confirmar_pab_{idx}"
                    )

                    if st.button(
                        "📤 Agregar recordatorio a COLA_ENVIO",
                        key=f"agregar_pab_{idx}",
                        use_container_width=True,
                        type="primary",
                        disabled=(
                            (not confirmar)
                            or
                            (not puede_agregar)
                        )
                    ):

                        try:

                            id_envio = agregar_recordatorio_pab_a_cola(
                                fila,
                                vista_previa
                            )

                            st.success(
                                "✅ Recordatorio agregado correctamente "
                                f"a COLA_ENVIO. ID: {id_envio}"
                            )

                            st.rerun()

                        except Exception as e:

                            st.error(
                                f"❌ No se pudo agregar: {e}"
                            )

                st.caption(
                    "Masivos Bravo valida duplicados, Mora 180 y "
                    "Excluir_correo antes de escribir en COLA_ENVIO."
                )


# ============================================================
# RESPUESTAS
# ============================================================

elif menu == "💬 Respuestas":

    st.markdown(
        '<div class="titulo">'
        '💬 Respuestas'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitulo">'
        'Bandeja de respuestas recibidas de clientes'
        '</div>',
        unsafe_allow_html=True
    )

    m1, m2, m3, m4 = st.columns(
        4
    )

    m1.metric(
        "💬 Total",
        total_respuestas
    )

    m2.metric(
        "🟡 Pendientes",
        total_pendientes
    )

    m3.metric(
        "🔴 Pendientes +24h",
        pendientes_24
    )

    m4.metric(
        "✅ Atendidas",
        total_atendidas
    )

    st.markdown("---")

    st.subheader(
        "🔎 Buscar y filtrar"
    )

    f1, f2, f3 = st.columns(
        [1.5, 1, 1]
    )

    with f1:

        buscar = st.text_input(
            "Buscar",
            placeholder=(
                "Nombre, referencia, correo "
                "o respuesta..."
            )
        )

    with f2:

        encargados_resp = []

        if "ENCARGADO" in respuestas.columns:

            encargados_resp = sorted(
                [
                    str(x).strip()
                    for x in respuestas[
                        "ENCARGADO"
                    ].unique()
                    if str(x).strip()
                ]
            )

        encargado_filtro = st.selectbox(
            "Encargado",
            ["Todos"] + encargados_resp,
            key="encargado_resp"
        )

    with f3:

        estado_filtro = st.selectbox(
            "Estado",
            [
                "Pendientes",
                "Pendientes +24h",
                "Atendidas",
                "Todas"
            ]
        )

    vista = respuestas.copy()

    if estado_filtro == "Pendientes":

        vista = vista[
            ~vista["_ATENDIDA"]
        ]

    elif estado_filtro == "Pendientes +24h":

        vista = vista[
            vista["_MAS_24H"]
        ]

    elif estado_filtro == "Atendidas":

        vista = vista[
            vista["_ATENDIDA"]
        ]

    if (
        encargado_filtro != "Todos"
        and
        "ENCARGADO" in vista.columns
    ):

        vista = vista[
            vista["ENCARGADO"]
            == encargado_filtro
        ]

    if buscar.strip():

        texto = (
            buscar
            .lower()
            .strip()
        )

        columnas_busqueda = [
            c
            for c in [
                "NOMBRE_CLIENTE",
                "REFERENCIA",
                "EMAIL_CLIENTE",
                "RESPUESTA",
                "ASUNTO"
            ]
            if c in vista.columns
        ]

        mascara = pd.Series(
            False,
            index=vista.index
        )

        for col in columnas_busqueda:

            mascara |= (
                vista[col]
                .astype(str)
                .str.lower()
                .str.contains(
                    texto,
                    regex=False,
                    na=False
                )
            )

        vista = vista[
            mascara
        ]

    vista = vista.sort_values(
        "_FECHA",
        ascending=False,
        na_position="last"
    )

    st.markdown("---")

    st.write(
        f"### 📥 {len(vista)} resultados"
    )

    for idx, fila in (
        vista
        .head(50)
        .iterrows()
    ):

        id_respuesta = str(
            fila.get(
                "ID_RESPUESTA",
                ""
            )
        ).strip()

        nombre = str(
            fila.get(
                "NOMBRE_CLIENTE",
                "Cliente"
            )
        ).strip()

        referencia = str(
            fila.get(
                "REFERENCIA",
                ""
            )
        ).strip()

        encargado = str(
            fila.get(
                "ENCARGADO",
                ""
            )
        ).strip()

        asunto = str(
            fila.get(
                "ASUNTO",
                ""
            )
        ).strip()

        respuesta_cliente = str(
            fila.get(
                "RESPUESTA",
                ""
            )
        ).strip()

        comentario_actual = str(
            fila.get(
                "COMENTARIOS",
                ""
            )
        ).strip()

        atendida = bool(
            fila["_ATENDIDA"]
        )

        mas_24h = bool(
            fila["_MAS_24H"]
        )

        if atendida:

            estado_txt = (
                "✅ Atendida"
            )

        elif mas_24h:

            estado_txt = (
                "🔴 Pendiente +24h"
            )

        else:

            estado_txt = (
                "🟡 Pendiente"
            )

        titulo = (
            f"{estado_txt} · "
            f"{nombre}"
        )

        if referencia:

            titulo += (
                f" · {referencia}"
            )

        with st.expander(
            titulo,
            expanded=False
        ):

            a, b, c = st.columns(
                3
            )

            a.write(
                "**Cliente**"
            )

            a.write(
                nombre
            )

            b.write(
                "**Referencia**"
            )

            b.write(
                referencia
            )

            c.write(
                "**Encargado**"
            )

            c.write(
                encargado
            )

            d, e, f = st.columns(3)
            d.write("**Fecha respuesta**")
            fecha_txt = fila.get("FECHA_RESPUESTA", "")
            d.write(str(fecha_txt) if str(fecha_txt).strip() else "—")
            e.write("**Campaña**")
            camp_txt = str(fila.get("ID_CAMPAÑA", "")).strip()
            e.write(camp_txt if camp_txt else "—")
            f.write("**Tiempo pendiente**")
            f.write("Gestionada" if atendida else str(fila.get("_TIEMPO_PENDIENTE", "Sin fecha")))

            email_txt = str(fila.get("EMAIL_CLIENTE", "")).strip()
            if email_txt:
                st.caption(f"📧 {email_txt}")

            st.markdown("---")

            st.write(
                "**Asunto**"
            )

            st.write(
                asunto
            )

            st.write(
                "**Respuesta del cliente**"
            )

            st.info(
                respuesta_cliente
                if respuesta_cliente
                else "Sin texto"
            )

            st.markdown("---")

            comentario_nuevo = st.text_area(
                "🗒️ Comentario interno",
                value=comentario_actual,
                placeholder=(
                    "Escribe aquí el comentario "
                    "de gestión..."
                ),
                height=100,
                key=f"comentario_{id_respuesta}_{idx}"
            )

            if atendida:

                st.success(
                    "✅ Esta respuesta ya está gestionada."
                )

                if st.button(
                    "💾 Guardar comentario",
                    key=f"guardar_{id_respuesta}_{idx}",
                    use_container_width=True
                ):

                    try:

                        guardar_comentario(
                            id_respuesta,
                            comentario_nuevo
                        )

                        st.rerun()

                    except Exception as e:

                        st.error(
                            f"Error: {e}"
                        )

            else:

                col1, col2 = st.columns(
                    2
                )

                with col1:

                    if st.button(
                        "💾 Guardar comentario",
                        key=f"guardar_{id_respuesta}_{idx}",
                        use_container_width=True
                    ):

                        try:

                            guardar_comentario(
                                id_respuesta,
                                comentario_nuevo
                            )

                            st.rerun()

                        except Exception as e:

                            st.error(
                                f"Error: {e}"
                            )

                with col2:

                    if st.button(
                        "✅ Marcar gestionada",
                        key=f"gestionar_{id_respuesta}_{idx}",
                        use_container_width=True,
                        type="primary"
                    ):

                        try:

                            marcar_respuesta_gestionada(
                                id_respuesta,
                                comentario_nuevo
                            )

                            st.rerun()

                        except Exception as e:

                            st.error(
                                f"Error: {e}"
                            )


# ============================================================
# CAMPAÑAS
# ============================================================

elif menu == "📧 Campañas":

    st.markdown(
        '<div class="titulo">📧 Campañas</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="subtitulo">Crea, revisa y prepara campañas manuales sin enviar correos automáticamente</div>',
        unsafe_allow_html=True
    )

    st.info(
        "🔒 Modo seguro: crear o preparar una campaña NO envía correos. "
        "Los registros preparados quedan en estado BORRADOR."
    )

    tab_nueva, tab_preparar, tab_historial = st.tabs([
        "➕ Nueva campaña",
        "👥 Revisar y preparar",
        "📋 Campañas creadas",
    ])

    with tab_nueva:
        st.subheader("Crear campaña manual")

        tipo_campana = st.selectbox(
            "Tipo / segmento",
            TIPOS_CAMPANA_MANUAL,
            key="tipo_nueva_campana"
        )

        es_personalizada = tipo_campana == "PERSONALIZADA"
        ids_plantillas = obtener_ids_plantillas_activas()

        with st.form("form_nueva_campana", clear_on_submit=False):
            c1, c2 = st.columns(2)

            with c1:
                nombre_campana = st.text_input(
                    "Nombre de la campaña",
                    placeholder=(
                        "Ej. Seguimiento personalizado - septiembre"
                        if es_personalizada
                        else "Ej. Seguimiento Mora 30 - septiembre"
                    )
                )

            with c2:
                fecha_campana = st.date_input(
                    "Fecha programada",
                    value=HOY
                )
                hora_campana = st.time_input(
                    "Hora programada",
                    value=AHORA.replace(second=0, microsecond=0).time()
                )

            if es_personalizada:
                if ids_plantillas:
                    plantilla_campana = st.selectbox(
                        "Plantilla",
                        ids_plantillas,
                        help="Elige la plantilla que se usará para esta campaña personalizada."
                    )
                else:
                    plantilla_campana = ""
                    st.error("No encontré plantillas activas en PLANTILLAS.")

                referencias_personalizadas = st.text_area(
                    "Referencias de clientes",
                    placeholder=(
                        "Pega una referencia por línea. También puedes separarlas por coma.\n"
                        "Ejemplo:\n3227405997\n3145427821\n3175113385"
                    ),
                    height=180,
                    help="El sistema buscará nombre y correo en CLIENTES y aplicará las exclusiones antes de preparar."
                )
            else:
                plantilla_campana = MAPA_PLANTILLAS_MORA.get(tipo_campana, "")
                st.text_input(
                    "Plantilla",
                    value=plantilla_campana,
                    disabled=True,
                    help="La plantilla se asigna automáticamente según la mora."
                )
                referencias_personalizadas = ""

            comentarios_campana = st.text_area(
                "Comentarios",
                placeholder="Opcional",
                height=90
            )

            crear = st.form_submit_button(
                "💾 Crear como BORRADOR",
                type="primary",
                use_container_width=True
            )

        if crear:
            try:
                id_nueva = crear_campana_manual(
                    nombre_campana,
                    tipo_campana,
                    plantilla_campana,
                    fecha_campana,
                    hora_campana,
                    comentarios_campana,
                    referencias_personalizadas
                )
                st.success(
                    f"✅ Campaña {id_nueva} creada como BORRADOR. No se envió ningún correo."
                )
                st.rerun()
            except Exception as e:
                st.error(f"❌ No se pudo crear la campaña: {e}")

    with tab_preparar:
        st.subheader("Revisar y preparar campaña")
        st.caption(
            "Selecciona un borrador para ver exactamente qué clientes entrarían. "
            "Nada se envía en esta etapa."
        )

        if campanas.empty or "ID_CAMPAÑA" not in campanas.columns:
            st.info("Todavía no hay campañas disponibles.")
        else:
            estados_permitidos = {"BORRADOR", "PREPARADA", "PROGRAMADA"}
            estados_serie = (
                campanas["ESTADO"]
                if "ESTADO" in campanas.columns
                else pd.Series("", index=campanas.index)
            )
            disponibles = campanas[
                estados_serie.astype(str).str.strip().str.upper().isin(estados_permitidos)
            ].copy()

            if disponibles.empty:
                st.info("No hay campañas en BORRADOR, PREPARADA o PROGRAMADA para revisar.")
            else:
                opciones = disponibles["ID_CAMPAÑA"].astype(str).tolist()
                id_sel = st.selectbox(
                    "Campaña",
                    opciones,
                    format_func=lambda x: (
                        f"{x} · "
                        f"{str(disponibles.loc[disponibles['ID_CAMPAÑA'].astype(str) == x, 'NOMBRE_CAMPAÑA'].iloc[0])}"
                        if "NOMBRE_CAMPAÑA" in disponibles.columns else x
                    )
                )

                fila_sel = fila_campana_por_id(id_sel)

                if fila_sel is not None:
                    d1, d2, d3, d4 = st.columns(4)
                    d1.metric("Tipo", str(fila_sel.get("FILTRO", "")))
                    d2.metric("Plantilla", str(fila_sel.get("PLANTILLA", "")))
                    d3.metric("Fecha", str(fila_sel.get("FECHA_ENVIO", "")))
                    d4.metric("Estado", str(fila_sel.get("ESTADO", "")))

                    tipo_sel = str(fila_sel.get("FILTRO", "")).strip().upper()

                    try:
                        candidatos, resumen = preparar_clientes_campana(fila_sel)

                        if tipo_sel == "PERSONALIZADA":
                            r1, r2, r3, r4, r5, r6 = st.columns(6)
                            r1.metric("Elegibles", len(candidatos))
                            r2.metric("No encontradas", resumen.get("no_encontradas", 0))
                            r3.metric("Excluir_correo", resumen.get("excluidos", 0))
                            r4.metric("Mora 180", resumen.get("mora_180", 0))
                            r5.metric("Sin correo", resumen.get("sin_email", 0))
                            r6.metric("Duplicados", resumen.get("duplicados", 0))
                        else:
                            r1, r2, r3, r4, r5 = st.columns(5)
                            r1.metric("Elegibles", len(candidatos))
                            r2.metric("Excluir_correo", resumen.get("excluidos", 0))
                            r3.metric("Mora 180", resumen.get("mora_180", 0))
                            r4.metric("Sin correo", resumen.get("sin_email", 0))
                            r5.metric("Duplicados", resumen.get("duplicados", 0))

                        if candidatos.empty:
                            st.info("No hay destinatarios elegibles con estas reglas.")
                        else:
                            mostrar = [
                                c for c in [
                                    "REFERENCIA", "NOMBRE", "EMAIL", "MORA",
                                    "ENCARGADO", "SALDO", "ASUNTO_PREVIO"
                                ] if c in candidatos.columns
                            ]
                            st.dataframe(
                                candidatos[mostrar].head(500),
                                use_container_width=True,
                                hide_index=True
                            )

                            st.caption(
                                f"Vista previa: {min(len(candidatos), 500)} de {len(candidatos)} destinatarios."
                            )

                            estado_actual = str(fila_sel.get("ESTADO", "")).strip().upper()

                            if estado_actual in {"PREPARADA", "PROGRAMADA"}:
                                if estado_actual == "PREPARADA":
                                    st.success(
                                        "✅ Esta campaña ya fue preparada. Los correos siguen en BORRADOR."
                                    )
                                    st.markdown("### 🗓️ Programar envío")
                                    fecha_programada = str(fila_sel.get("FECHA_ENVIO", "")).strip() or "Sin fecha"
                                    hora_programada = str(fila_sel.get("HORA_ENVIO", "")).strip() or "Sin hora"
                                    p1, p2, p3 = st.columns(3)
                                    p1.metric("Destinatarios", len(candidatos))
                                    p2.metric("Fecha programada", fecha_programada)
                                    p3.metric("Hora programada", hora_programada)
                                    st.warning(
                                        "🔒 Programar ahora SOLO cambia la campaña a PROGRAMADA. "
                                        "Los correos permanecen en BORRADOR y NO se enviarán."
                                    )
                                    confirmar_programacion = st.checkbox(
                                        "Confirmo fecha, hora, plantilla y destinatarios.",
                                        key=f"confirmar_programacion_{id_sel}"
                                    )
                                    col_prog, col_cancel = st.columns(2)
                                    with col_prog:
                                        if st.button(
                                            "🗓️ Programar campaña (sin enviar)",
                                            type="primary",
                                            use_container_width=True,
                                            disabled=not confirmar_programacion,
                                            key=f"programar_{id_sel}"
                                        ):
                                            try:
                                                cantidad = programar_campana_segura(id_sel)
                                                st.success(
                                                    f"✅ Campaña PROGRAMADA con {cantidad} correos. "
                                                    "Todos continúan en BORRADOR; no se envió nada."
                                                )
                                                st.rerun()
                                            except Exception as e:
                                                st.error(f"❌ No se pudo programar: {e}")
                                    with col_cancel:
                                        if st.button(
                                            "↩️ Cancelar preparación",
                                            use_container_width=True,
                                            key=f"cancelar_preparacion_{id_sel}"
                                        ):
                                            try:
                                                eliminadas = cancelar_preparacion_campana(id_sel)
                                                st.success(
                                                    f"↩️ Campaña devuelta a BORRADOR. "
                                                    f"Se retiraron {eliminadas} borradores de COLA_ENVIO."
                                                )
                                                st.rerun()
                                            except Exception as e:
                                                st.error(f"❌ No se pudo cancelar: {e}")
                                else:
                                    st.success(
                                        "🗓️ Campaña PROGRAMADA. Por seguridad, sus correos continúan en BORRADOR y NO se enviarán."
                                    )
                                    fecha_programada = str(fila_sel.get("FECHA_ENVIO", "")).strip() or "Sin fecha"
                                    hora_programada = str(fila_sel.get("HORA_ENVIO", "")).strip() or "Sin hora"
                                    p1, p2, p3 = st.columns(3)
                                    p1.metric("Destinatarios", len(candidatos))
                                    p2.metric("Fecha", fecha_programada)
                                    p3.metric("Hora", hora_programada)
                                    confirmar_cancelacion = st.checkbox(
                                        "Confirmo que quiero cancelar esta programación y volver a BORRADOR.",
                                        key=f"confirmar_cancelacion_{id_sel}"
                                    )
                                    if st.button(
                                        "↩️ Cancelar programación y volver a BORRADOR",
                                        use_container_width=True,
                                        disabled=not confirmar_cancelacion,
                                        key=f"cancelar_programacion_{id_sel}"
                                    ):
                                        try:
                                            eliminadas = cancelar_preparacion_campana(id_sel)
                                            st.success(
                                                f"↩️ Programación cancelada. Se retiraron {eliminadas} "
                                                "borradores de COLA_ENVIO."
                                            )
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ No se pudo cancelar: {e}")
                            else:
                                confirmar = st.checkbox(
                                    "Confirmo que revisé los destinatarios. Preparar en COLA_ENVIO como BORRADOR.",
                                    key=f"confirmar_{id_sel}"
                                )

                                if st.button(
                                    "📥 Preparar campaña",
                                    type="primary",
                                    use_container_width=True,
                                    disabled=not confirmar,
                                    key=f"preparar_{id_sel}"
                                ):
                                    try:
                                        agregados, omitidos = preparar_campana_en_cola(
                                            id_sel,
                                            candidatos
                                        )
                                        st.success(
                                            f"✅ Preparada: {agregados} destinatarios en BORRADOR. "
                                            f"Omitidos por duplicado: {omitidos}. No se envió ningún correo."
                                        )
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"❌ No se pudo preparar: {e}")

                    except Exception as e:
                        st.error(f"❌ No pude construir los destinatarios: {e}")

    with tab_historial:
        st.subheader("Panel de campañas")
        st.caption(
            "Consulta el estado operativo de cada campaña y sus destinatarios. "
            "Este panel es solo de control: no activa envíos."
        )

        if campanas.empty:
            st.info("No hay campañas registradas.")
        else:
            vista_camp = campanas.copy()

            # Normalizar columnas mínimas sin modificar Google Sheets.
            for col in [
                "ID_CAMPAÑA", "NOMBRE_CAMPAÑA", "PLANTILLA", "FILTRO",
                "FECHA_ENVIO", "HORA_ENVIO", "ESTADO", "TOTAL_CLIENTES",
                "ENVIADOS", "PENDIENTES", "ERRORES", "FECHA_CREACIÓN"
            ]:
                if col not in vista_camp.columns:
                    vista_camp[col] = ""

            # Conteos reales desde COLA_ENVIO. Así el panel no depende de que
            # los contadores históricos de CAMPAÑAS estén actualizados.
            conteos_cola = {}
            if not cola.empty and "ID_CAMPAÑA" in cola.columns:
                cola_panel = cola.copy()
                if "ESTADO" not in cola_panel.columns:
                    cola_panel["ESTADO"] = ""
                cola_panel["_ID_CAMP"] = cola_panel["ID_CAMPAÑA"].astype(str).str.strip()
                cola_panel["_ESTADO"] = cola_panel["ESTADO"].astype(str).str.strip().str.upper()

                for camp_id, grupo in cola_panel.groupby("_ID_CAMP"):
                    estados = grupo["_ESTADO"]
                    enviados = int(estados.eq("ENVIADO").sum())
                    errores = int(estados.eq("ERROR").sum())
                    borradores = int(estados.eq("BORRADOR").sum())
                    pendientes = int(estados.isin(["PENDIENTE", "PROGRAMADO", "PROGRAMADA"]).sum())
                    conteos_cola[camp_id] = {
                        "DESTINATARIOS": len(grupo),
                        "ENVIADOS_REAL": enviados,
                        "PENDIENTES_REAL": pendientes,
                        "BORRADORES_REAL": borradores,
                        "ERRORES_REAL": errores,
                    }

            def _conteo_camp(row, campo, fallback=0):
                camp_id = str(row.get("ID_CAMPAÑA", "")).strip()
                if camp_id in conteos_cola:
                    return conteos_cola[camp_id].get(campo, fallback)
                try:
                    valor = row.get("TOTAL_CLIENTES", fallback)
                    if valor in (None, ""):
                        return fallback
                    return int(float(str(valor).replace(",", "")))
                except Exception:
                    return fallback

            vista_camp["DESTINATARIOS"] = vista_camp.apply(
                lambda r: _conteo_camp(r, "DESTINATARIOS", 0), axis=1
            )
            vista_camp["ENVIADOS_PANEL"] = vista_camp.apply(
                lambda r: conteos_cola.get(str(r.get("ID_CAMPAÑA", "")).strip(), {}).get(
                    "ENVIADOS_REAL", 0
                ), axis=1
            )
            vista_camp["PENDIENTES_PANEL"] = vista_camp.apply(
                lambda r: conteos_cola.get(str(r.get("ID_CAMPAÑA", "")).strip(), {}).get(
                    "PENDIENTES_REAL", 0
                ), axis=1
            )
            vista_camp["BORRADORES_PANEL"] = vista_camp.apply(
                lambda r: conteos_cola.get(str(r.get("ID_CAMPAÑA", "")).strip(), {}).get(
                    "BORRADORES_REAL", 0
                ), axis=1
            )
            vista_camp["ERRORES_PANEL"] = vista_camp.apply(
                lambda r: conteos_cola.get(str(r.get("ID_CAMPAÑA", "")).strip(), {}).get(
                    "ERRORES_REAL", 0
                ), axis=1
            )

            if "FECHA_CREACIÓN" in vista_camp.columns:
                vista_camp["_ORDEN"] = pd.to_datetime(
                    vista_camp["FECHA_CREACIÓN"], errors="coerce", dayfirst=True
                )
                vista_camp = vista_camp.sort_values(
                    "_ORDEN", ascending=False, na_position="last"
                )

            estados_panel = vista_camp["ESTADO"].astype(str).str.strip().str.upper()
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Campañas", len(vista_camp))
            m2.metric("Borrador", int(estados_panel.eq("BORRADOR").sum()))
            m3.metric(
                "Preparadas / programadas",
                int(estados_panel.isin(["PREPARADA", "PROGRAMADA"]).sum())
            )
            m4.metric(
                "Finalizadas",
                int(estados_panel.isin(["FINALIZADA", "FINALIZADA CON ERRORES"]).sum())
            )

            filtro_estado = st.multiselect(
                "Filtrar por estado",
                sorted([x for x in estados_panel.unique().tolist() if x]),
                default=[],
                key="filtro_estado_panel_campanas"
            )
            vista_filtrada = vista_camp.copy()
            if filtro_estado:
                vista_filtrada = vista_filtrada[
                    vista_filtrada["ESTADO"].astype(str).str.strip().str.upper().isin(filtro_estado)
                ].copy()

            tabla_panel = pd.DataFrame({
                "CAMPAÑA": vista_filtrada["NOMBRE_CAMPAÑA"].astype(str),
                "TIPO": vista_filtrada["FILTRO"].astype(str),
                "PLANTILLA": vista_filtrada["PLANTILLA"].astype(str),
                "ESTADO": vista_filtrada["ESTADO"].astype(str),
                "DESTINATARIOS": vista_filtrada["DESTINATARIOS"],
                "BORRADORES": vista_filtrada["BORRADORES_PANEL"],
                "ENVIADOS": vista_filtrada["ENVIADOS_PANEL"],
                "PENDIENTES": vista_filtrada["PENDIENTES_PANEL"],
                "ERRORES": vista_filtrada["ERRORES_PANEL"],
                "FECHA": vista_filtrada["FECHA_ENVIO"].astype(str),
                "HORA": vista_filtrada["HORA_ENVIO"].astype(str),
            })

            st.dataframe(
                tabla_panel,
                use_container_width=True,
                hide_index=True
            )

            st.markdown("### 🔎 Detalle de campaña")
            ids_detalle = vista_camp["ID_CAMPAÑA"].astype(str).tolist()
            if ids_detalle:
                id_detalle = st.selectbox(
                    "Selecciona una campaña",
                    ids_detalle,
                    format_func=lambda x: (
                        f"{x} · "
                        f"{str(vista_camp.loc[vista_camp['ID_CAMPAÑA'].astype(str) == x, 'NOMBRE_CAMPAÑA'].iloc[0])}"
                    ),
                    key="detalle_panel_campanas"
                )

                detalle = vista_camp[
                    vista_camp["ID_CAMPAÑA"].astype(str) == str(id_detalle)
                ].iloc[0]
                estado_detalle = str(detalle.get("ESTADO", "")).strip().upper()

                d1, d2, d3, d4 = st.columns(4)
                d1.metric("Estado", estado_detalle or "—")
                d2.metric("Destinatarios", int(detalle.get("DESTINATARIOS", 0) or 0))
                d3.metric("Enviados", int(detalle.get("ENVIADOS_PANEL", 0) or 0))
                d4.metric("Errores", int(detalle.get("ERRORES_PANEL", 0) or 0))

                st.write(
                    f"**Tipo:** {detalle.get('FILTRO', '')}  |  "
                    f"**Plantilla:** {detalle.get('PLANTILLA', '')}  |  "
                    f"**Programación:** {detalle.get('FECHA_ENVIO', '')} {detalle.get('HORA_ENVIO', '')}"
                )

                if not cola.empty and "ID_CAMPAÑA" in cola.columns:
                    detalle_cola = cola[
                        cola["ID_CAMPAÑA"].astype(str).str.strip() == str(id_detalle).strip()
                    ].copy()
                else:
                    detalle_cola = pd.DataFrame()

                if detalle_cola.empty:
                    st.info(
                        "Esta campaña todavía no tiene registros en COLA_ENVIO. "
                        "Si está en BORRADOR, puedes prepararla desde la pestaña Revisar y preparar."
                    )
                else:
                    columnas_detalle = [
                        c for c in [
                            "REFERENCIA", "NOMBRE", "EMAIL", "PLANTILLA",
                            "ESTADO", "FECHA_PROG", "FECHA_ENVIO", "INTENTOS", "ERROR"
                        ] if c in detalle_cola.columns
                    ]
                    st.dataframe(
                        detalle_cola[columnas_detalle],
                        use_container_width=True,
                        hide_index=True
                    )

                if estado_detalle == "BORRADOR":
                    st.info("➡️ Acción disponible: revisar destinatarios y preparar la campaña.")
                elif estado_detalle == "PREPARADA":
                    st.info("➡️ Acción disponible: programar o cancelar la preparación.")
                elif estado_detalle == "PROGRAMADA":
                    st.success(
                        "🗓️ Campaña programada. Por seguridad, los correos siguen en BORRADOR y no se enviarán."
                    )
                elif estado_detalle in {"FINALIZADA", "FINALIZADA CON ERRORES"}:
                    st.success("✅ Campaña cerrada.")


# ============================================================
# PENDIENTES
# ============================================================

elif menu == "⚠️ Pendientes":

    st.markdown('<div class="titulo">⚠️ Centro de pendientes</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitulo">Prioriza respuestas de clientes que todavía no han sido gestionadas</div>',
        unsafe_allow_html=True
    )

    pendientes_resp = respuestas[~respuestas["_ATENDIDA"]].copy()
    pendientes_resp = pendientes_resp.sort_values("_FECHA", ascending=True, na_position="last")

    p1, p2, p3, p4 = st.columns(4)
    p1.metric("🟡 Sin gestionar", len(pendientes_resp))
    p2.metric("🔴 +24 horas", pendientes_24)
    p3.metric("🟢 Menos de 24h", max(0, len(pendientes_resp) - pendientes_24))
    p4.metric("🏦 PaB pendientes", total_recordatorios_pendientes)

    st.markdown("---")

    if pendientes_resp.empty:
        st.success("✅ No hay respuestas pendientes de gestión.")
    else:
        f1, f2 = st.columns([1, 1])
        with f1:
            prioridad = st.selectbox(
                "Prioridad",
                ["Todas", "Solo +24h", "Menos de 24h"],
                key="prioridad_pendientes"
            )
        with f2:
            encargados_p = []
            if "ENCARGADO" in pendientes_resp.columns:
                encargados_p = sorted([
                    str(x).strip() for x in pendientes_resp["ENCARGADO"].unique()
                    if str(x).strip()
                ])
            encargado_p = st.selectbox(
                "Encargado", ["Todos"] + encargados_p, key="encargado_pendientes"
            )

        vista_p = pendientes_resp.copy()
        if prioridad == "Solo +24h":
            vista_p = vista_p[vista_p["_MAS_24H"]]
        elif prioridad == "Menos de 24h":
            vista_p = vista_p[~vista_p["_MAS_24H"]]
        if encargado_p != "Todos" and "ENCARGADO" in vista_p.columns:
            vista_p = vista_p[vista_p["ENCARGADO"].astype(str).str.strip() == encargado_p]

        columnas_p = [c for c in [
            "FECHA_RESPUESTA", "REFERENCIA", "NOMBRE_CLIENTE", "EMAIL_CLIENTE",
            "ID_CAMPAÑA", "ASUNTO", "ENCARGADO", "_TIEMPO_PENDIENTE"
        ] if c in vista_p.columns]
        tabla_p = vista_p[columnas_p].copy()
        tabla_p = tabla_p.rename(columns={"_TIEMPO_PENDIENTE": "TIEMPO PENDIENTE"})
        st.dataframe(tabla_p.head(500), use_container_width=True, hide_index=True)
        st.caption(f"Mostrando {min(len(vista_p), 500)} de {len(vista_p)} respuestas pendientes.")

        st.markdown("---")
        st.subheader("🛠️ Gestionar pendiente")

        opciones_p = vista_p.index.tolist()
        if opciones_p:
            idx_p = st.selectbox(
                "Selecciona una respuesta",
                opciones_p,
                format_func=lambda i: (
                    f"{'🔴' if bool(vista_p.loc[i, '_MAS_24H']) else '🟡'} "
                    f"{str(vista_p.loc[i].get('NOMBRE_CLIENTE', 'Cliente'))} · "
                    f"{str(vista_p.loc[i].get('REFERENCIA', ''))} · "
                    f"{str(vista_p.loc[i].get('_TIEMPO_PENDIENTE', ''))}"
                ),
                key="detalle_pendiente"
            )
            fp = vista_p.loc[idx_p]
            idp = str(fp.get("ID_RESPUESTA", "")).strip()

            q1, q2, q3, q4 = st.columns(4)
            q1.metric("Cliente", str(fp.get("NOMBRE_CLIENTE", "")) or "—")
            q2.metric("Referencia", str(fp.get("REFERENCIA", "")) or "—")
            q3.metric("Encargado", str(fp.get("ENCARGADO", "")) or "—")
            q4.metric("Pendiente", str(fp.get("_TIEMPO_PENDIENTE", "")) or "—")

            st.write("**Asunto**")
            st.write(str(fp.get("ASUNTO", "")) or "—")
            st.write("**Respuesta del cliente**")
            st.info(str(fp.get("RESPUESTA", "")) or "Sin texto")

            comentario_p = st.text_area(
                "🗒️ Comentario interno",
                value=str(fp.get("COMENTARIOS", "")).strip(),
                height=100,
                key=f"comentario_pendiente_{idp}_{idx_p}"
            )
            c1, c2 = st.columns(2)
            with c1:
                if st.button("💾 Guardar comentario", key=f"guardar_p_{idp}_{idx_p}", use_container_width=True):
                    try:
                        guardar_comentario(idp, comentario_p)
                        st.success("Comentario guardado.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
            with c2:
                if st.button("✅ Marcar gestionada", key=f"gestionar_p_{idp}_{idx_p}", use_container_width=True, type="primary"):
                    try:
                        marcar_respuesta_gestionada(idp, comentario_p)
                        st.success("Respuesta marcada como gestionada.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")


# ============================================================
# PLANTILLAS
# ============================================================

elif menu == "📝 Plantillas":

    st.title(
        "📝 Plantillas"
    )

    st.dataframe(
        plantillas,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# HISTORIAL
# ============================================================

elif menu == "🕘 Historial":

    st.title(
        "🕘 Historial"
    )

    st.dataframe(
        cola,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# CONFIGURACIÓN
# ============================================================

elif menu == "⚙️ Configuración":

    st.title(
        "⚙️ Configuración"
    )

    st.success(
        "✅ Google Sheets conectado"
    )

    st.write(
        "**Respuestas:** escritura habilitada"
    )

    st.write(
        "**Pagos a Banco:** vista previa + agregar a COLA_ENVIO"
    )

    st.write(
        "**Datos PaB:** completa NOMBRE y EMAIL desde "
        "Info_Clientes_V2 cuando estén vacíos"
    )

    st.write(
        "**Zona horaria:** America/Bogota"
    )
