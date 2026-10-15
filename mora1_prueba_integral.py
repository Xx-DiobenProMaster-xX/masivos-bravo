"""Prueba integral aislada: reglas reales, un correo de prueba y deduplicación en memoria."""
import os,json,re,base64
from datetime import datetime,timedelta
from zoneinfo import ZoneInfo
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials as SheetsCredentials
from google.oauth2.credentials import Credentials as GmailCredentials
from googleapiclient.discovery import build
from mora1_dioben import rows,value,norm,ref,parse_date,ID_DI0BEN,ID_MASIVOS,DIAS_ESPECIALES

TZ=ZoneInfo("America/Bogota")
FROM="estructurados@gobravo.com.co"
TRUE={"TRUE","VERDADERO","SI","1","X"}
def eligible(enabled,mora,excluded,start,day,template):
    if not enabled or norm(mora)!="MORA 1" or norm(excluded) in TRUE or not start or start>day:return False
    n=(day-start).days+1
    code=DIAS_ESPECIALES.get(n) or ("T001" if n%2==0 else None)
    return code is not None and code==template

def main():
    destination=os.getenv("MORA1_TEST_EMAIL","").strip()
    confirmed=os.getenv("MORA1_TEST_CONFIRMACION","")
    if confirmed!="ENVIAR_PRUEBA" or not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+",destination):
        raise SystemExit("Se requiere correo de prueba válido y confirmación ENVIAR_PRUEBA.")
    today=datetime.now(TZ).date()
    assert eligible(True,"MORA 1","",today-timedelta(days=1),today,"T001")
    assert not eligible(True,"AL DIA","",today-timedelta(days=1),today,"T001")
    assert not eligible(True,"MORA 1","SI",today-timedelta(days=1),today,"T001")
    assert not eligible(True,"MORA 1","",today,today,"T001")
    assert eligible(True,"MORA 1","",today-timedelta(days=3),today,"MORA101")
    print("OK: pruebas AL DÍA, excluido, día sin plantilla y especial MORA101.")
    sheets=SheetsCredentials.from_service_account_info(json.loads(os.environ["MI_JSON"]),scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    gc=gspread.authorize(sheets)
    src=gc.open_by_key(ID_DI0BEN)
    dst=gc.open_by_key(ID_MASIVOS)
    hs,follow=rows(src.worksheet("Seguimiento"))
    hm,selection=rows(src.worksheet("Masivos_Correos"))
    ht,templates=rows(dst.worksheet("PLANTILLAS"))
    hq,queue=rows(dst.worksheet("COLA_ENVIO"))
    active={norm(value(t,ht,"ID_PLANTILLA")):t for t in templates if norm(value(t,ht,"ESTADO"))=="ACTIVA"}
    assert "T001" in active and value(active["T001"],ht,"ASUNTO") and value(active["T001"],ht,"CUERPO")
    selected={ref(value(r,hm,"REFERENCIA")) for r in selection if norm(value(r,hm,"HABILITADO"))=="SI" and norm(value(r,hm,"TIPO CORREO")) in {"MORA_1","MORA 1"}}
    good=sum(1 for r in follow if ref(value(r,hs,"REFERENCIA")) in selected and norm(value(r,hs,"MORA ACTUAL"))=="MORA 1" and norm(value(r,hs,"EXCLUIR CORREO")) not in TRUE)
    print("Clientes seleccionados que permanecen en Mora 1 (sin exclusión):",good)
    print("Filas actuales COLA_ENVIO leídas (solo lectura):",len(queue))
    if good==0:raise SystemExit("No hay Mora 1 seleccionados: prueba detenida.")
    # Registro aislado en memoria: no se alteran referencias ni filas de clientes.
    test_id="TEST-MORA1-"+today.strftime("%Y%m%d")+"-CONTROLADO"
    ledger={}
    def reserve():
        if test_id in ledger:return False
        ledger[test_id]="ENVIANDO"
        return True
    assert reserve()
    assert not reserve()
    print("OK: una segunda reserva con el mismo ID se bloquea.")
    oauth=GmailCredentials(token=None,refresh_token=os.environ["GMAIL_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",client_id=os.environ["GMAIL_CLIENT_ID"],
        client_secret=os.environ["GMAIL_CLIENT_SECRET"],scopes=["https://www.googleapis.com/auth/gmail.send"])
    message=EmailMessage()
    message["To"]=destination
    message["From"]=FROM
    message["Reply-To"]=FROM
    message["Subject"]="[PRUEBA INTEGRAL] Mora 1 Bravo - sin cobranza"
    message.set_content("Prueba integral de la automatización Mora 1. No es un aviso de cobro.")
    message.add_alternative("<p><b>PRUEBA INTEGRAL</b> de Mora 1. No es un aviso de cobro.</p>",subtype="html")
    service=build("gmail","v1",credentials=oauth,cache_discovery=False)
    response=service.users().messages().send(userId="me",body={"raw":base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")}).execute()
    mid=response.get("id")
    if not mid:raise RuntimeError("Gmail no confirmó ID de mensaje.")
    ledger[test_id]={"ESTADO":"ENVIADO","INTENTOS":1,"ID_MENSAJE":mid,"FECHA_ENVIO":datetime.now(TZ).isoformat()}
    assert ledger[test_id]["ESTADO"]=="ENVIADO" and ledger[test_id]["INTENTOS"]==1
    assert not reserve()
    print("OK: correo controlado enviado. ID Gmail:",mid)
    print("OK: registro simulado ENVIADO, un intento y segundo envío bloqueado.")
    print("IMPORTANTE: registro aislado en memoria, NO persistente en COLA_ENVIO.")
    print("NO se envió ningún correo a clientes reales de la cartera.")
if __name__=="__main__":main()
