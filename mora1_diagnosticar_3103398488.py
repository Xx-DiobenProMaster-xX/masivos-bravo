"""Diagnóstico de una referencia de Mora 1 sin escribir ni enviar."""
import json,os,collections
from datetime import datetime
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials
from mora1_dioben import ID_DI0BEN,ID_MASIVOS,rows,value,ref,norm,parse_date,DIAS_ESPECIALES
k="3103398488"
today=datetime.now(ZoneInfo("America/Bogota")).date()
cr=Credentials.from_service_account_info(json.loads(os.environ["MI_JSON"]),scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
gc=gspread.authorize(cr);source=gc.open_by_key(ID_DI0BEN);target=gc.open_by_key(ID_MASIVOS)
hs,seguimiento=rows(source.worksheet("Seguimiento"))
hm,seleccion=rows(source.worksheet("Masivos_Correos"))
hc,cola=rows(target.worksheet("COLA_ENVIO"))
he,clientes=rows(target.worksheet("CLIENTES"))
ss=[r for r in seguimiento if ref(value(r,hs,"REFERENCIA"))==k]
ms=[r for r in seleccion if ref(value(r,hm,"REFERENCIA"))==k]
cs=[r for r in cola if ref(value(r,hc,"REFERENCIA"))==k]
cl=[r for r in clientes if ref(value(r,he,"REFERENCIA"))==k]
print("REFERENCIA:",k,"| FECHA BOGOTÁ:",today.isoformat())
print("FILAS: Seguimiento",len(ss),"Masivos_Correos",len(ms),"COLA_ENVIO",len(cs),"CLIENTES",len(cl))
if not ss:print("CAUSA: NO APARECE EN Seguimiento")
if len(ss)>1:print("CAUSA: REFERENCIA DUPLICADA EN Seguimiento")
if not ms:print("CAUSA: NO APARECE EN Masivos_Correos")
if len(ms)>1:print("AVISO: REFERENCIA DUPLICADA EN Masivos_Correos")
for i,r in enumerate(ss,1):
    mora=norm(value(r,hs,"MORA ACTUAL"))
    exc=norm(value(r,hs,"EXCLUIR CORREO"))
    fecha=value(r,hs,"FECHA_INICIO_MORA_1")
    parsed=parse_date(fecha)
    print("SEGUIMIENTO",i,"MORA_ACTUAL=",mora,"EXCLUIR_CORREO=",exc or "(vacío)","FECHA_INICIO_MORA_1=",fecha or "(vacío)")
    if mora!="MORA 1":print("CAUSA: NO ESTÁ EN MORA 1")
    elif exc in {"TRUE","VERDADERO","SI","1","X"}:print("CAUSA EXACTA: EXCLUIDO POR EXCLUIR CORREO")
    elif not parsed or parsed>today:print("CAUSA: FECHA DE INICIO DE MORA 1 INVÁLIDA, VACÍA O FUTURA")
    else:
        day=(today-parsed).days+1
        plantilla=DIAS_ESPECIALES.get(day) or ("T001" if day%2==0 else None)
        print("DÍA DEL CICLO:",day,"PLANTILLA_CORRESPONDIENTE:",plantilla or "NINGUNA HOY")
for i,r in enumerate(ms,1):
    habilitado=norm(value(r,hm,"HABILITADO"))
    tipo=norm(value(r,hm,"TIPO CORREO"))
    print("MASIVOS_CORREOS",i,"HABILITADO=",habilitado or "(vacío)","TIPO_CORREO=",tipo or "(vacío)")
    if habilitado!="SI":print("CAUSA: NO HABILITADO PARA CORREO")
    elif tipo not in {"MORA_1","MORA 1"}:print("CAUSA: TIPO DE CORREO DIFERENTE DE MORA 1")
for r in cs[-12:]:
    print("HISTORIAL","PLANTILLA=",value(r,hc,"PLANTILLA"),"ESTADO=",value(r,hc,"ESTADO"),"FECHA_PROG=",value(r,hc,"FECHA_PROG"),"FECHA_ENVIO=",value(r,hc,"FECHA_ENVIO"),"INTENTOS=",value(r,hc,"INTENTOS"),"ERROR=",value(r,hc,"ERROR")[:100])
print("SOLO LECTURA; NO SE ENVIÓ CORREO; NO SE MODIFICÓ NINGUNA HOJA.")
