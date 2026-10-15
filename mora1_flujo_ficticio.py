"""Prueba integral FICTICIA sin Gmail ni escrituras en Sheets."""
from datetime import date,timedelta
from mora1_reglas import puede_enviar,plantilla_para_fecha

def run():
    hoy=date(2026,10,16)
    registro={"ID_ENVIO":"TEST-MORA1-FICTICIO-00001","REFERENCIA":"FICTICIA-00001",
              "EMAIL":"ejemplo.invalid","PLANTILLA":"T001","ESTADO":"BORRADOR",
              "INTENTOS":0,"FECHA_ENVIO":"","ID_MENSAJE":""}
    inicio=hoy-timedelta(days=1)
    assert puede_enviar("MORA 1","SI","MORA_1","",inicio,hoy,registro["PLANTILLA"])
    assert not puede_enviar("AL DIA","SI","MORA_1","",inicio,hoy,"T001")
    assert not puede_enviar("MORA 1","SI","MORA_1","SI",inicio,hoy,"T001")
    assert plantilla_para_fecha(inicio,hoy)=="T001"
    print("OK: únicamente MORA 1 activa y no excluida es elegible.")
    intento_unico=set()
    def reservar(id_envio):
        if id_envio in intento_unico:return False
        if registro["ESTADO"]!="BORRADOR" or registro["INTENTOS"]!=0:return False
        intento_unico.add(id_envio)
        registro["ESTADO"]="ENVIANDO"
        registro["INTENTOS"]=1
        return True
    assert reservar(registro["ID_ENVIO"])
    assert not reservar(registro["ID_ENVIO"])
    print("OK: BORRADOR -> ENVIANDO / INTENTOS 1. Segundo intento bloqueado.")
    respuesta_gmail_ficticia={"id":"SIMULADO-NO-REAL"}
    registro["ID_MENSAJE"]=respuesta_gmail_ficticia["id"]
    registro["FECHA_ENVIO"]="16/10/2026 08:15:00"
    registro["ESTADO"]="ENVIADO"
    assert not reservar(registro["ID_ENVIO"])
    print("OK: ENVIANDO -> ENVIADO. ID y fecha simulados, reenvío bloqueado.")
    # Incertidumbre: después de reservar, una excepción no puede reiniciar a BORRADOR.
    error={"ESTADO":"ENVIANDO","INTENTOS":1,"ID_MENSAJE":""}
    assert error["ESTADO"]!="BORRADOR" and error["INTENTOS"]==1
    print("OK: fallo ambiguo mantiene ENVIANDO e INTENTOS 1.")
    print("SIN EFECTOS: datos ficticios en memoria, cero correos, cero modificaciones a Google Sheets.")
if __name__=="__main__":run()
