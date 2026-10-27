"""Recuperación HISTÓRICA controlada, independiente de la bienvenida diaria."""
import os,json
from datetime import datetime,date,timedelta
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials
from bienvenida_estructurados import (
  PIPELINE_SPREADSHEET_ID,PIPELINE_SHEET,PIPELINE_SHEET_MES,
  MASIVOS_SPREADSHEET_ID,ESTRUCTURADOS_SPREADSHEET_ID,HOJA_EXCLUIR,
  PLANTILLA_ID,SCOPES,parse_fecha,norm_ref,es_true,
  maestro_clientes,maestro_asignaciones_vigentes,correo_valido,
  gmail_service,html_bienvenida_estructurados,enviar_gmail,ASUNTO_BIENVENIDA,
  agregar_cola,ocultar_email,
)
TZ=ZoneInfo("America/Bogota")
INICIO=date(2026,9,24)
def main():
  hoy=datetime.now(TZ).date()
  fin=hoy-timedelta(days=1)
  real=os.getenv("BIENVENIDA_HISTORICA_ENVIOS_REALES","NO").strip().upper()=="SI"
  limite=max(0,min(int(os.getenv("BIENVENIDA_HISTORICA_LIMITE","1")),200))
  if not os.getenv("MI_JSON"):raise RuntimeError("Falta MI_JSON")
  gc=gspread.authorize(Credentials.from_service_account_info(json.loads(os.environ["MI_JSON"]),scopes=SCOPES))
  p=gc.open_by_key(PIPELINE_SPREADSHEET_ID)
  eventos={}
  for nombre in (PIPELINE_SHEET,PIPELINE_SHEET_MES):
    for fila in p.worksheet(nombre).get("A:P")[1:]:
      fecha=parse_fecha(fila[2] if len(fila)>2 else "")
      ref=norm_ref(fila[7] if len(fila)>7 else "")
      if ref and fecha and INICIO<=fecha<=fin and es_true(fila[15] if len(fila)>15 else ""):
        eventos[(ref,fecha)]=nombre
  excluidos={norm_ref(v) for v in gc.open_by_key(ESTRUCTURADOS_SPREADSHEET_ID).worksheet(HOJA_EXCLUIR).col_values(1)[1:] if norm_ref(v)}
  ws=gc.open_by_key(MASIVOS_SPREADSHEET_ID).worksheet("COLA_ENVIO")
  datos=ws.get_all_values()
  columnas={str(x).strip():i for i,x in enumerate(datos[0])}
  req={"ID_ENVIO","ESTADO","REFERENCIA","PLANTILLA","ID_MENSAJE"}
  if not req.issubset(columnas):raise RuntimeError("Faltan columnas de cola: "+str(req-set(columnas)))
  def campo(f,n):
    k=columnas[n]
    return str(f[k]).strip() if k<len(f) else ""
  prev={}
  for fila in datos[1:]:
    if campo(fila,"PLANTILLA")==PLANTILLA_ID or campo(fila,"ID_ENVIO").startswith("ENV-ESTBIENV-"):
      prev.setdefault(norm_ref(campo(fila,"REFERENCIA")),[]).append(fila)
  maestro=maestro_clientes(gc)
  respaldo=maestro_asignaciones_vigentes(gc,hoy)
  enviados=0
  stats={"YA_REGISTRADO":0,"EXCLUIDO":0,"SIN_CORREO":0,"LISTO":0,"ENVIADO":0}
  gmail=None
  print("MODO:", "ENVIO REAL" if real else "SIMULACION SIN ENVIOS","LIMITE:",limite)
  print("VENTANA:",INICIO,"a",fin)
  for (ref,fecha),origen in sorted(eventos.items(),key=lambda e:(e[0][1],e[0][0])):
    # Seguro: si hay CUALQUIER bienvenida histórica en cola, no volver a enviar.
    if prev.get(ref):
      stats["YA_REGISTRADO"]+=1
      continue
    if ref in excluidos:
      stats["EXCLUIDO"]+=1
      continue
    a,b=maestro.get(ref,{}),respaldo.get(ref,{})
    ea=str(a.get("EMAIL","")).strip().lower()
    eb=str(b.get("EMAIL","")).strip().lower()
    email=ea if correo_valido(ea) else eb if correo_valido(eb) else ""
    if not email:
      stats["SIN_CORREO"]+=1
      continue
    stats["LISTO"]+=1
    if not real or enviados>=limite:continue
    nombre=str(a.get("NOMBRE","") or b.get("NOMBRE","") or "Cliente").strip()
    envio_id=f"ENV-ESTBIENV-{fecha:%Y%m%d}-{ref}-{PLANTILLA_ID}"
    # Conciliación final justo antes de reservar, incluso si otras ejecuciones escribieron.
    actuales=ws.get_all_values()
    encabezado={str(c).strip():i for i,c in enumerate(actuales[0])}
    for need in ("ID_ENVIO","REFERENCIA","PLANTILLA"):
      if need not in encabezado:raise RuntimeError("Falta "+need)
    if any(
       (str(f[encabezado["ID_ENVIO"]]).startswith("ENV-ESTBIENV-") or
        str(f[encabezado["PLANTILLA"]])==PLANTILLA_ID)
       and norm_ref(f[encabezado["REFERENCIA"]])==ref
       for f in actuales[1:]
       if len(f)>max(encabezado.values())
    ):
      print("OMITIDO POR REGISTRO NUEVO",ref)
      continue
    # Reservar antes de enviar, y verificar la persistencia.
    agregar_cola(ws,envio_id,ref,nombre,email,fecha,"ENVIANDO")
    posiciones=[i+1 for i,v in enumerate(ws.col_values(1)) if v==envio_id]
    if len(posiciones)!=1:raise RuntimeError("Reserva no confirmada; NO enviar "+ref)
    fila_num=posiciones[0]
    prev.setdefault(ref,[]).append([])
    if gmail is None:gmail=gmail_service()
    try:
      resp=enviar_gmail(gmail,email,ASUNTO_BIENVENIDA,html_bienvenida_estructurados(nombre))
      mensaje=str(resp.get("id","")).strip()
      if not mensaje:raise RuntimeError("Gmail sin message id")
      timestamp=datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
      ws.batch_update([
        {"range":f"M{fila_num}","values":[[mensaje]]},
        {"range":f"J{fila_num}","values":[[timestamp]]},
        {"range":f"H{fila_num}","values":[["ENVIADO"]]}
      ],value_input_option="RAW")
      enviados+=1
      stats["ENVIADO"]+=1
      print("ENVIADO",ref,ocultar_email(email))
    except Exception as exc:
      print("ENVIO AMBIGUO:",ref,type(exc).__name__,"; queda ENVIANDO. No reintentar.")
      raise
  print("RESUMEN:",json.dumps(stats,ensure_ascii=False))
  print("NO cambia workflow ni script de bienvenida diaria.")
if __name__=="__main__":main()
