"""Reglas comunes Mora 1 sin dependencias externas."""
from datetime import date
DIAS_ESPECIALES={4:"MORA101",10:"MORA102",15:"MORA103",19:"MORA104",24:"MORA105"}
def plantilla_para_fecha(inicio,hoy):
    if not isinstance(inicio,date) or not isinstance(hoy,date) or inicio>hoy:
        return None
    dia=(hoy-inicio).days+1
    return DIAS_ESPECIALES.get(dia) or ("T001" if dia%2==0 else None)
def puede_enviar(mora_actual,habilitado,tipo_correo,excluir,inicio,hoy,plantilla):
    if str(mora_actual or "").strip().upper()!="MORA 1":return False
    if str(habilitado or "").strip().upper()!="SI":return False
    if str(tipo_correo or "").strip().upper() not in ("MORA_1","MORA 1"):return False
    if str(excluir or "").strip().upper() in ("SI","TRUE","VERDADERO","1","X"):return False
    esperado=plantilla_para_fecha(inicio,hoy)
    return bool(esperado and plantilla==esperado)
