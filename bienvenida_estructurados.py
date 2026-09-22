import os, re, json, unicodedata
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials

TZ = ZoneInfo("America/Bogota")

PIPELINE_SPREADSHEET_ID = "1H3sYEtkeu47POnu8xZMaMtID1Vj53YIcWblWeZ8d0rc"
PIPELINE_SHEET = "BD 2026"

MASIVOS_SPREADSHEET_ID = "1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
CARTERA_SPREADSHEET_ID = "13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
ESTRUCTURADOS_SPREADSHEET_ID = "15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"

HOJAS_INFO_CLIENTES = ["Info_Clientes_V2", "Hoja Info_Clientes_V2", ". Hoja Info_Clientes_V2"]
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

def main():
    ahora = datetime.now(TZ)
    hoy = ahora.date()
    desde = hoy - timedelta(days=6)

    print("=" * 76)
    print("DIAGNÓSTICO 7 DÍAS - BIENVENIDA A ESTRUCTURADOS")
    print(f"Fecha/hora Colombia: {ahora.strftime('%d/%m/%Y %H:%M:%S')}")
    print(
        f"Ventana revisada: {desde.strftime('%d/%m/%Y')} "
        f"al {hoy.strftime('%d/%m/%Y')}"
    )
    print("NINGÚN CORREO SERÁ ENVIADO / NINGÚN SHEET SERÁ MODIFICADO")
    print("=" * 76)

    gc = sheets()
    filas = (
        gc.open_by_key(PIPELINE_SPREADSHEET_ID)
        .worksheet(PIPELINE_SHEET)
        .get("A:P")
    )
    maestro = maestro_clientes(gc)
    excluir = exclusiones(gc)
    existentes = ids_existentes(gc)

    # Una referencia puede aparecer varias veces.
    # Consolidamos por referencia + fecha de liquidación.
    # Si al menos una fila de ese evento tiene P=TRUE,
    # consideramos la referencia estructurada para esa fecha.
    eventos = {}
    invalidas = 0
    filas_ventana = 0
    filas_true = 0

    for f in filas[1:]:
        ref = norm_ref(f[7] if len(f) > 7 else "")        # H
        fecha = parse_fecha(f[2] if len(f) > 2 else "")  # C
        est = es_true(f[15] if len(f) > 15 else "")      # P

        if not fecha:
            invalidas += 1
            continue

        if not (desde <= fecha <= hoy):
            continue

        filas_ventana += 1

        if not ref:
            continue

        clave = (ref, fecha.isoformat())
        evento = eventos.setdefault(
            clave,
            {
                "REFERENCIA": ref,
                "FECHA": fecha,
                "ESTRUCTURADO": False,
                "FILAS": 0,
            },
        )
        evento["FILAS"] += 1

        if est:
            evento["ESTRUCTURADO"] = True
            filas_true += 1

    estructurados = [
        e for e in eventos.values()
        if e["ESTRUCTURADO"]
    ]

    print(f"Filas revisadas en BD 2026: {max(len(filas) - 1, 0)}")
    print(f"Filas dentro de la ventana: {filas_ventana}")
    print(f"Filas dentro de la ventana con P=TRUE: {filas_true}")
    print(f"Eventos únicos estructurados: {len(estructurados)}")
    print(f"Filas con fecha inválida/vacía: {invalidas}")
    print()

    print("=" * 76)
    print("RESUMEN POR FECHA")
    print("=" * 76)

    for n in range(7):
        fecha = desde + timedelta(days=n)
        eventos_fecha = [
            e for e in estructurados
            if e["FECHA"] == fecha
        ]
        refs_fecha = {
            e["REFERENCIA"]
            for e in eventos_fecha
        }
        print(
            f"{fecha.strftime('%d/%m/%Y')} | "
            f"{len(refs_fecha)} referencias estructuradas"
        )

    print()
    print("=" * 76)
    print("DETALLE DE ESTRUCTURADOS EN LOS ÚLTIMOS 7 DÍAS")
    print("=" * 76)

    total_con_email = 0
    total_sin_email = 0
    total_excluidos = 0

    for fecha in sorted({e["FECHA"] for e in estructurados}):
        eventos_fecha = sorted(
            [e for e in estructurados if e["FECHA"] == fecha],
            key=lambda x: x["REFERENCIA"],
        )

        print()
        print(
            f"--- {fecha.strftime('%d/%m/%Y')} "
            f"({len(eventos_fecha)} referencias) ---"
        )

        for e in eventos_fecha:
            ref = e["REFERENCIA"]
            cli = maestro.get(ref, {})
            nombre = str(cli.get("NOMBRE", "")).strip() or "SIN NOMBRE"
            email = str(cli.get("EMAIL", "")).strip().lower()

            estado = "OK"

            if ref in excluir:
                estado = "EXCLUIR_CORREO"
                total_excluidos += 1
            elif not correo_valido(email):
                estado = "SIN_CORREO"
                total_sin_email += 1
            else:
                total_con_email += 1

            # Solo como referencia diagnóstica:
            # muestra si el eventual ID de bienvenida ya existe.
            id_envio = (
                "ENV-ESTBIENV-"
                f"{fecha.strftime('%Y%m%d')}-"
                f"{ref}-"
                f"{PLANTILLA_ID}"
            )
            if id_envio in existentes:
                estado += " | YA_EN_COLA"

            email_mostrar = ocultar_email(email) if email else "SIN EMAIL"

            print(
                f"{ref} | "
                f"{email_mostrar} | "
                f"{nombre} | "
                f"{estado}"
            )

    print()
    print("=" * 76)
    print("RESUMEN DE CALIDAD")
    print("=" * 76)
    print(f"Estructurados únicos/eventos encontrados: {len(estructurados)}")
    print(f"Con correo válido: {total_con_email}")
    print(f"Sin correo válido: {total_sin_email}")
    print(f"En Excluir_correo: {total_excluidos}")
    print()
    print("DIAGNÓSTICO FINALIZADO.")
    print("NINGÚN CORREO FUE ENVIADO.")
    print("NINGUNA FILA FUE AGREGADA O MODIFICADA.")
    print("=" * 76)


if __name__ == "__main__":
    main()
