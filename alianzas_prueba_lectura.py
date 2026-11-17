"""Piloto aislado Alianzas - una sola dirección controlada. Sin acceso a Sheets."""
import os
import base64
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

REMITENTE = "alianzasco@gobravo.com.co"
DESTINO = "dioben2523@gmail.com"
GMAIL_SEND = "https://www.googleapis.com/auth/gmail.send"
GMAIL_SETTINGS = "https://www.googleapis.com/auth/gmail.settings.basic"

def main():
    habilitado = os.environ.get("ALIANZAS_PRUEBA_CONFIRMADA", "").strip()
    if habilitado != "ENVIAR_UNA_PRUEBA":
        print("SIMULACION: no se envio Gmail.")
        print("Remitente esperado:", REMITENTE)
        print("Destinatario unico:", DESTINO)
        print("Solicitud de lectura: cabecera Disposition-Notification-To (no garantizada).")
        return

    variables = ("GMAIL_CLIENT_ID", "GMAIL_CLIENT_SECRET", "ALIANZAS_GMAIL_REFRESH_TOKEN")
    if any(not os.environ.get(k) for k in variables):
        raise RuntimeError("Faltan credenciales OAuth de Gmail configuradas en GitHub.")

    credenciales = Credentials(
        token=None,
        refresh_token=os.environ["ALIANZAS_GMAIL_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["GMAIL_CLIENT_ID"],
        client_secret=os.environ["GMAIL_CLIENT_SECRET"],
        scopes=[GMAIL_SEND, GMAIL_SETTINGS],
    )
    svc = build("gmail", "v1", credentials=credenciales, cache_discovery=False)
    aliases = svc.users().settings().sendAs().list(userId="me").execute().get("sendAs", [])
    autorizado = any(
        str(alias.get("sendAsEmail", "")).strip().lower() == REMITENTE
        and alias.get("verificationStatus") == "accepted"
        for alias in aliases
    )
    if not autorizado:
        raise RuntimeError("ENVIO BLOQUEADO: Alianzas no esta verificado como Enviar correo como en la cuenta OAuth. No se envio mensaje.")
    print("OK: remitente Alianzas verificado en Gmail.")
    msg = EmailMessage()
    msg["From"] = f"Bravo Alianzas <{REMITENTE}>"
    msg["To"] = DESTINO
    msg["Reply-To"] = REMITENTE
    msg["Subject"] = "[PRUEBA CONTROLADA] Alianzas Bravo - confirmacion de lectura"
    msg["Disposition-Notification-To"] = REMITENTE
    msg.set_content(
        "PRUEBA CONTROLADA de Alianzas Bravo.\n"
        "Este correo valida el remitente y una solicitud experimental de confirmacion de lectura.\n"
        "No es un mensaje de cobro y no requiere ninguna accion."
    )
    msg.add_alternative("""<html><body style="font-family:Arial,sans-serif;background:#f8f7fc;padding:28px">
      <table role="presentation" style="max-width:540px;margin:auto;background:white;border-top:5px solid #3a2980;border-collapse:collapse" width="100%">
        <tr><td style="padding:28px"><h2 style="color:#3a2980">BRAVO | Alianzas</h2>
        <p>Esta es una <strong>prueba controlada</strong> del remitente de Alianzas.</p>
        <p>Hemos solicitado una confirmacion de lectura experimental.
        Su recepcion depende de la compatibilidad y las preferencias del destinatario.</p>
        <p style="color:#666;font-size:13px">No es una comunicacion de cobro.</p>
        </td></tr></table></body></html>""", subtype="html")
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
    resultado = svc.users().messages().send(userId="me", body={"raw":raw}).execute()
    if not resultado.get("id"):
        raise RuntimeError("Gmail no devolvio identificador de mensaje; verificar antes de repetir.")
    print("GMAIL ACEPTO EL ENVIO DE PRUEBA.")
    print("ID_MENSAJE:", resultado["id"])
    print("Remitente solicitado:",REMITENTE)
    print("Destino:",DESTINO)
    print("Fecha Bogotá:",datetime.now(ZoneInfo("America/Bogota")).strftime("%Y-%m-%d %H:%M"))
    print("NOTA: cuenta personal gmail.com no soporta confirmaciones de lectura nativas.")
    print("Verificar manualmente el remitente real y el mensaje recibido antes de usar masivos.")

if __name__=="__main__":
    main()
