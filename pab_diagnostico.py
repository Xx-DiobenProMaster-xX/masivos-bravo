"""Diagnóstico PaB solo lectura. Reproduce decisiones de pab_send_auto sin Gmail."""
from collections import Counter
from datetime import datetime
from zoneinfo import ZoneInfo
from pab_send_auto import (
    gc, MASIVOS_ID, txt, ref, fecha, headers, getv,
    load_client_map,load_exclusions,load_templates,es_mora_180
)
def main():
    today=datetime.now(ZoneInfo("America/Bogota")).date()
    g=gc()
    book=g.open_by_key(MASIVOS_ID)
    ws=book.worksheet("PAB_PROXIMOS")
    data=ws.get_all_values()
    if not data:
        print("PAB_PROXIMOS vacía.")
        return
    h=headers(data[0])
    cola=book.worksheet("COLA_ENVIO").get_all_values()
    qh=headers(cola[0])
    existentes={getv(r,qh,"ID_ENVIO"):getv(r,qh,"ESTADO").upper() for r in cola[1:] if getv(r,qh,"ID_ENVIO")}
    clientes=load_client_map(g)
    exclusiones=load_exclusions(g)
    plantillas=load_templates(g)
    causas=Counter()
    candidatos=Counter()
    enviados_potenciales=0
    print("FECHA:",today,"| FILAS PAB_PROXIMOS:",len(data)-1)
    for n,r in enumerate(data[1:],start=2):
        fp=fecha(getv(r,h,"FECHA_PAB"))
        if fp is None:
            causas["FECHA_PAB_INVALIDA"]+=1
            continue
        dias=(fp-today).days
        if dias not in (0,3):
            causas["FUERA_DE_VENTANA_HOY_O_3_DIAS"]+=1
            continue
        pid="PAB000" if dias==0 else "PAB003"
        flag="AVISO_HOY" if dias==0 else "AVISO_3_DIAS"
        candidatos[pid]+=1
        rr=ref(getv(r,h,"REFERENCIA"))
        motivo=None
        if getv(r,h,flag):motivo="AVISO_PREVIO_MARCADO"
        elif not rr:motivo="REFERENCIA_INVALIDA"
        elif rr in exclusiones:motivo="EXCLUIDO"
        else:
            cli=clientes.get(rr,{})
            email=getv(r,h,"EMAIL") or cli.get("email","")
            mora=getv(r,h,"MORA") or cli.get("mora","")
            if not email:motivo="SIN_CORREO"
            elif es_mora_180(mora):motivo="MORA_180"
            elif pid not in plantillas:motivo="PLANTILLA_FALTANTE"
            else:
                eid=f"ENV-PAB-{today:%Y%m%d}-{rr}-{pid}"
                estado=existentes.get(eid)
                if estado=="ENVIADO":motivo="COLA_YA_ENVIADO"
                elif estado and estado not in ("BORRADOR","PENDIENTE"):motivo="COLA_NO_REINTENTAR_"+estado
        if motivo:
            causas[motivo]+=1
            print(f"OMITIDO | fila {n} | ref {rr or '(vacía)'} | {pid} | {motivo}")
        else:
            enviados_potenciales+=1
            print(f"APTO | fila {n} | ref {rr} | {pid} | fecha {fp:%d/%m/%Y}")
    print("CANDIDATOS EN VENTANA:",dict(candidatos))
    print("APTOS PARA PROCESAR:",enviados_potenciales)
    print("CAUSAS:",dict(causas))
    print("SOLO LECTURA: no Gmail, no actualizaciones de Sheets.")
if __name__=="__main__":main()
