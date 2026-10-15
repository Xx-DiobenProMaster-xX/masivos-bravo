"""Prueba controlada de persistencia: hoja separada, un correo, idempotencia."""
import os,json,re,base64
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials as SheetsCredentials
from google.oauth2.credentials import Credentials as GmailCredentials
from googleapiclient.discovery import build
from mora1_dioben import ID_MASIVOS

TZ=ZoneInfo("America/Bogota")
SHEET="MORA1_PRUEBA_PERSISTENCIA"
HEADERS=["CLAVE_PRUEBA","DESTINATARIO","ESTADO","INTENTOS","FECHA_ENVIO","ID_MENSAJE","ERROR"]
FROM="estructurados@gobravo.com.co"

def main():
    email=os.environ.get("MORA1_TEST_EMAIL","").strip().lower()
    confirm=os.environ.get("MORA1_TEST_CONFIRMACION","")
    if confirm!="ENVIAR_PRUEBA" or not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+",email):
        raise SystemExit("Se requiere destinatario válido y confirmación ENVIAR_PRUEBA.")
    key=os.environ.get("MORA1_TEST_CLAVE","").strip()
    if not re.fullmatch(r"[A-Za-z0-9_-]{5,60}",key):raise SystemExit("CLAVE de prueba inválida; 5 a 60 caracteres alfanuméricos o guion.")
    creds=SheetsCredentials.from_service_account_info(json.loads(os.environ["MI_JSON"]),scopes=["https://www.googleapis.com/auth/spreadsheets"])
    book=gspread.authorize(creds).open_by_key(ID_MASIVOS)
    try:ws=book.worksheet(SHEET)
    except gspread.WorksheetNotFound:
        # Aislado: no crea ni cambia COLA_ENVIO, CAMPAÑAS, CLIENTES.
        ws=book.add_worksheet(title=SHEET,rows=200,cols=len(HEADERS))
        ws.update("A1:G1",[HEADERS])
    data=ws.get_all_values()
    if not data or data[0][:len(HEADERS)]!=HEADERS:
        raise RuntimeError("Encabezados de prueba inesperados; se detuvo sin modificar datos.")
    found=[(n,r) for n,r in enumerate(data[1:],start=2) if r and r[0]==key]
    if len(found)>1:raise RuntimeError("La clave de prueba está duplicada; se detuvo.")
    if found:
        n,row=found[0]
        if len(row)>1 and row[1].strip().lower()!=email:
            raise RuntimeError("La clave ya pertenece a otro destinatario.")
        print("Clave ya registrada. Estado:",row[2] if len(row)>2 else "(vacío)")
        print("NO SE ENVIÓ OTRO CORREO. Segunda ejecución bloqueada de forma persistente.")
        return
    # Registrar reserva ANTES de enviar; reintentos con la misma clave no reenvían.
    ws.append_row([key,email,"ENVIANDO",1,"","",""],value_input_option="RAW")
    current=ws.get_all_values()
    matches=[(i,r) for i,r in enumerate(current[1:],start=2) if r and r[0]==key]
    if len(matches)!=1:raise RuntimeError("No se confirmó reserva única; no se envió correo.")
    row_number=matches[0][0]
    try:
        oauth=GmailCredentials(token=None,refresh_token=os.environ["GMAIL_REFRESH_TOKEN"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=os.environ["GMAIL_CLIENT_ID"],client_secret=os.environ["GMAIL_CLIENT_SECRET"],
            scopes=["https://www.googleapis.com/auth/gmail.send"])
        mail=EmailMessage()
        mail["To"]=email
        mail["From"]=FROM
        mail["Reply-To"]=FROM
        mail["Subject"]="[PRUEBA PERSISTENCIA] Mora 1 Bravo - sin cobranza"
        mail.set_content("Prueba técnica de persistencia. No es aviso de cobro.")
        mail.add_alternative("<p><strong>PRUEBA TÉCNICA DE PERSISTENCIA</strong>. No es aviso de cobro.</p>",subtype="html")
        service=build("gmail","v1",credentials=oauth,cache_discovery=False)
        result=service.users().messages().send(userId="me",body={"raw":base64.urlsafe_b64encode(mail.as_bytes()).decode("ascii")}).execute()
        message_id=str(result.get("id",""))
        if not message_id:raise RuntimeError("No se recibió un ID de Gmail.")
        ws.batch_update([
            {"range":f"C{row_number}","values":[["ENVIADO"]]},
            {"range":f"E{row_number}","values":[[datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S")]]},
            {"range":f"F{row_number}","values":[[message_id]]},
        ],value_input_option="RAW")
        print("OK: envío de prueba registrado en",SHEET)
        print("ESTADO: ENVIADO | INTENTOS: 1 | ID Gmail:",message_id)
        print("Ejecuta nuevamente con la MISMA clave y destinatario: debe bloquearse.")
    except Exception as exc:
        ws.batch_update([{"range":f"C{row_number}","values":[["ERROR"]]},
                         {"range":f"G{row_number}","values":[[str(type(exc).__name__)]]}],value_input_option="RAW")
        print("El envío tuvo resultado fallido o ambiguo; NO repetir con la misma clave.")
        raise

if __name__=="__main__":main()
