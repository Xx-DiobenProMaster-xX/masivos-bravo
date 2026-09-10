
import streamlit as st
import pandas as pd
import gspread
import google.auth
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

def obtener_gc():

    credentials, _ = google.auth.default()

    return gspread.authorize(
        credentials
    )


def obtener_archivo():

    gc = obtener_gc()

    return gc.open_by_key(
        SPREADSHEET_ID
    )


@st.cache_data(ttl=60)
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
        "## BRAVO S.A.S."
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
        'Gestión de campañas y comunicaciones con clientes'
        '</div>',
        unsafe_allow_html=True
    )

    c1, c2, c3, c4 = st.columns(
        4
    )

    c1.metric(
        "💬 Respuestas",
        total_respuestas
    )

    c2.metric(
        "🟡 Pendientes",
        total_pendientes
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

    p1, p2, p3 = st.columns(
        3
    )

    p1.metric(
        "PaB próximos 3 días",
        total_pab_3_dias
    )

    p2.metric(
        "Valor PaB hoy",
        moneda(
            valor_pab_hoy
        )
    )

    p3.metric(
        "Recordatorios pendientes",
        total_recordatorios_pendientes
    )


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

                st.caption(
                    "🔒 PaB continúa en modo lectura. "
                    "No se generan correos desde esta pantalla todavía."
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

    st.title(
        "📧 Campañas"
    )

    columnas = [
        c
        for c in campanas.columns
        if not c.startswith(
            "COLUMNA_"
        )
    ]

    st.dataframe(
        campanas[
            columnas
        ],
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# PENDIENTES
# ============================================================

elif menu == "⚠️ Pendientes":

    st.title(
        "⚠️ Centro de pendientes"
    )

    p1, p2, p3 = st.columns(
        3
    )

    p1.metric(
        "Respuestas sin gestionar",
        total_pendientes
    )

    p2.metric(
        "Respuestas +24h",
        pendientes_24
    )

    p3.metric(
        "Recordatorios PaB pendientes",
        total_recordatorios_pendientes
    )


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
        "**Pagos a Banco:** solo lectura"
    )

    st.write(
        "**Zona horaria:** America/Bogota"
    )
