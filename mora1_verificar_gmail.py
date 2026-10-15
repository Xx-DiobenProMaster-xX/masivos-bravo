"""Valida OAuth de Gmail usando solo gmail.send; nunca envía correos."""
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError

names=("GMAIL_CLIENT_ID","GMAIL_CLIENT_SECRET","GMAIL_REFRESH_TOKEN")
missing=[n for n in names if not os.environ.get(n,"").strip()]
if missing:
    raise SystemExit("ERROR: faltan secretos en GitHub: "+", ".join(missing))

creds=Credentials(
    token=None,
    refresh_token=os.environ["GMAIL_REFRESH_TOKEN"],
    token_uri="https://oauth2.googleapis.com/token",
    client_id=os.environ["GMAIL_CLIENT_ID"],
    client_secret=os.environ["GMAIL_CLIENT_SECRET"],
    scopes=["https://www.googleapis.com/auth/gmail.send"],
)
try:
    creds.refresh(Request())
except RefreshError:
    print("ERROR: Google rechazó la renovación OAuth. Consultar con el administrador.")
    raise SystemExit(1)
except Exception as exc:
    print("ERROR: problema de conexión al renovar OAuth:",type(exc).__name__)
    raise SystemExit(1)
if not creds.valid:
    raise SystemExit("ERROR: la credencial renovada no es válida.")

print("OK: refresh token vigente; Google emitió un access token.")
if creds.granted_scopes:
    if "https://www.googleapis.com/auth/gmail.send" in creds.granted_scopes:
        print("OK: gmail.send aparece entre los permisos otorgados.")
    else:
        raise SystemExit("ERROR: gmail.send no aparece entre los permisos otorgados.")
else:
    print("AVISO: Google no informó los permisos otorgados; confirmar con una prueba controlada posterior.")
print("NO SE ENVIÓ NINGÚN CORREO.")
print("Nota: users.getProfile requiere permisos distintos de gmail.send; por eso se eliminó esa consulta de esta prueba.")
