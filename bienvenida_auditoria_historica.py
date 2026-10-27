"""Auditoría histórica de bienvenidas. SOLO LECTURA; no utiliza Gmail."""
import os
import json
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials
from bienvenida_estructurados import (
    PIPELINE_SPREADSHEET_ID, PIPELINE_SHEET, PIPELINE_SHEET_MES,
    MASIVOS_SPREADSHEET_ID, ESTRUCTURADOS_SPREADSHEET_ID, HOJA_EXCLUIR,
    PLANTILLA_ID, SCOPES, parse_fecha, norm_ref, es_true,
)

TZ = ZoneInfo("America/Bogota")
START = date(2026, 9, 24)

def main():
    hoy = datetime.now(TZ).date()
    hasta = hoy - timedelta(days=1)
    raw = os.environ.get("MI_JSON", "").strip()
    if not raw:
        raise RuntimeError("Falta MI_JSON")
    gc = gspread.authorize(Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES))
    pipeline = gc.open_by_key(PIPELINE_SPREADSHEET_ID)
    fuentes = [
        (PIPELINE_SHEET_MES, pipeline.worksheet(PIPELINE_SHEET_MES).get("A:P")),
        (PIPELINE_SHEET, pipeline.worksheet(PIPELINE_SHEET).get("A:P")),
    ]
    eventos = {}
    for origen, datos in fuentes:
        for fila in datos[1:]:
            fecha = parse_fecha(fila[2] if len(fila)>2 else "")
            referencia = norm_ref(fila[7] if len(fila)>7 else "")
            estructurado = es_true(fila[15] if len(fila)>15 else "")
            if fecha and referencia and estructurado and START <= fecha <= hasta:
                eventos[(referencia, fecha)] = origen

    excluir = {
        norm_ref(x)
        for x in gc.open_by_key(ESTRUCTURADOS_SPREADSHEET_ID).worksheet(HOJA_EXCLUIR).col_values(1)[1:]
        if norm_ref(x)
    }
    cola = gc.open_by_key(MASIVOS_SPREADSHEET_ID).worksheet("COLA_ENVIO").get_all_values()
    if not cola: raise RuntimeError("COLA_ENVIO vacía")
    cab = {str(v).strip(): i for i, v in enumerate(cola[0])}
    for nombre in ("ID_ENVIO", "REFERENCIA", "PLANTILLA", "ESTADO", "ID_MENSAJE"):
        if nombre not in cab: raise RuntimeError("Columna ausente " + nombre)
    def valor(fila, nombre):
        i = cab[nombre]
        return str(fila[i]).strip() if i<len(fila) else ""
    por_id = {valor(f, "ID_ENVIO"):f for f in cola[1:] if valor(f,"ID_ENVIO")}
    # Otros registros de bienvenida se revisan por referencia, para evitar
    # concluir erróneamente que una bienvenida histórica con otro ID no existió.
    otras_bienvenidas = {}
    for f in cola[1:]:
        if valor(f,"PLANTILLA") == PLANTILLA_ID or valor(f,"ID_ENVIO").startswith("ENV-ESTBIENV-"):
            otras_bienvenidas.setdefault(norm_ref(valor(f,"REFERENCIA")), []).append(f)
    totales = {"ENVIADO":0,"AMBIGUO":0,"EXCLUIDO":0,"PENDIENTE":0,"REVISAR_OTRO_REGISTRO":0}
    for referencia,fecha in sorted(eventos,key=lambda x:(x[1],x[0])):
        ident = f"ENV-ESTBIENV-{fecha:%Y%m%d}-{referencia}-{PLANTILLA_ID}"
        actual = por_id.get(ident)
        registros = ([actual] if actual is not None else []) + [f for f in otras_bienvenidas.get(referencia,[]) if f is not actual]
        if any(valor(f,"ESTADO").upper()=="ENVIADO" or valor(f,"ID_MENSAJE") for f in registros):
            estado="ENVIADO"
        elif actual is not None:
            estado="AMBIGUO"
        elif registros:
            estado="REVISAR_OTRO_REGISTRO"
        elif referencia in excluir:
            estado="EXCLUIDO"
        else:
            estado="PENDIENTE"
        totales[estado]+=1
        # No mostrar correos, nombres ni datos personales en logs.
        if estado!="ENVIADO":
            print(f"{estado} | REF {referencia} | FECHA {fecha:%d/%m/%Y} | ORIGEN {eventos[(referencia,fecha)]}")
    print("VENTANA:",START.isoformat(),"a",hasta.isoformat())
    print("EVENTOS ESTRUCTURADOS:",len(eventos))
    print("RESULTADO AUDITORÍA:",json.dumps(totales,ensure_ascii=False))
    print("SOLO LECTURA: sin Gmail, sin escrituras a Google Sheets.")

if __name__=="__main__":
    main()
