"""Exportación Python a Excel con fórmulas y gráficos XY nativos.

No contiene macros ni vínculos a otros libros. Todas las fórmulas se escriben
en inglés, como exige el formato XLSX, incluso si Excel se usa en español.
Se guardan valores calculados en Python como caché; Excel recalcula al editar.
El número de corrientes queda fijado al exportar. Para añadir/quitar corrientes,
usar Python o la interfaz y generar de nuevo el libro.
"""
from __future__ import annotations
import io
import math
import numpy as np
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col, xl_rowcol_to_cell as cell


def export_excel(r, destination=None):
    """Devuelve bytes si destination=None; si hay ruta, escribe allí."""
    s=r["streams"]; n=len(s); m=2*n; z=2*m-1
    if n>20:
        raise ValueError("La exportación detallada admite 20 corrientes. Divide el caso o amplía este límite en el código.")
    sink=io.BytesIO() if destination is None else str(destination)
    wb=xlsxwriter.Workbook(sink, {"strings_to_formulas":False,"strings_to_urls":False})
    wb.set_calc_mode("auto")
    wb.set_properties({"title":"Pinch Lab · Integración energética", "author":"Carlos Galán",
                       "comments":"Modelo reproducible en Python. Datos iniciales: diapositivas UPC 27 y 30."})
    fmt={
        "title":wb.add_format({"font_name":"Aptos Display","font_size":22,"bold":True,"font_color":"#142D4E"}),
        "sub":wb.add_format({"font_name":"Aptos","font_size":10,"font_color":"#52657D","text_wrap":True,"valign":"top"}),
        "head":wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#142D4E","text_wrap":True,"valign":"vcenter"}),
        "num":wb.add_format({"font_name":"Aptos","font_size":11,"num_format":"#,##0.00;[Red](#,##0.00);–"}),
        "input":wb.add_format({"font_name":"Aptos","font_color":"#175CA4","bg_color":"#EAF4FE","num_format":"#,##0.00"}),
        "text":wb.add_format({"font_name":"Aptos","font_size":11,"font_color":"#243C58"}),
        "metric":wb.add_format({"font_name":"Aptos","font_size":17,"bold":True,"font_color":"#148A83","num_format":"#,##0.00"}),
        "pct":wb.add_format({"num_format":"0.0%","font_color":"#243C58"}),
        "bad":wb.add_format({"bg_color":"#FBE5E3","font_color":"#AD3028"}),
        "hot":wb.add_format({"font_color":"#C64238","num_format":"#,##0.00"}),
        "cold":wb.add_format({"font_color":"#2275C9","num_format":"#,##0.00"}),
    }
    names=["Resumen","Entradas","Matriz_Q","Equipos","Cascada","Compuestas","Tramos_red","Sensibilidad","Guia"]
    sh={name:wb.add_worksheet(name) for name in names}
    for name,ws in sh.items():
        ws.hide_gridlines(2); ws.set_zoom(90); ws.set_default_row(21)
        ws.set_column(0,30,15,fmt["num"])
        ws.set_tab_color("#148A83" if name=="Resumen" else "#2275C9" if name=="Entradas" else "#A8B8CA")
        ws.freeze_panes(3,1); ws.set_landscape(); ws.fit_to_pages(1,0)
        ws.set_paper(9); ws.set_margins(.3,.3,.4,.4)
        ws.set_footer("&LPinch Lab&RPágina &P")
    def title(name,text,note):
        ws=sh[name]; ws.merge_range("A1:J1",text,fmt["title"]); ws.set_row(0,35)
        ws.merge_range("A2:J2",note,fmt["sub"]); ws.set_row(1,33)
    def head(ws,row,values):
        ws.write_row(row,0,values,fmt["head"]); ws.set_row(row,36)
    def f(ws,row,c,formula,value=0,style="num"):
        ws.write_formula(row,c,formula,fmt[style],value)
    def ir(c):
        return f"'Entradas'!${c}$11:${c}${10+n}"
    def tr(c):
        return f"'Tramos_red'!${col(c)}$4:${col(c)}${3+z}"

    # 1. Entradas: única fuente de datos editable.
    ws=sh["Entradas"]
    title("Entradas","Datos del proceso","Edita las celdas azules. CP = caudal × calor específico. Unidades originales del PDF.")
    ws.write("A4","ΔT mínimo (K)",fmt["text"]); ws.write("B4",r["dtmin"],fmt["input"])
    ws.write("A5","U (kJ/h·m²·K)",fmt["text"]); ws.write("B5",r["u"],fmt["input"])
    ws.merge_range("C5:H5","U=0: área sin calcular. U no viene en el PDF; introduce tu hipótesis si la conoces.",fmt["sub"])
    ws.merge_range("A7:K7","",fmt["text"])
    valid=f'AND(ISNUMBER(B4),B4>=0,ISNUMBER(B5),B5>=0,COUNT(B11:D{10+n})={3*n},MIN(B11:B{10+n})>0,MIN(C11:D{10+n})>=-273.15,SUMPRODUCT(--(C11:C{10+n}=D11:D{10+n}))=0,COUNTA(A11:A{10+n})={n},SUMPRODUCT(1/COUNTIF(A11:A{10+n},A11:A{10+n}))={n})'
    f(ws,6,0,f'=IF({valid},"Datos válidos","REVISAR ENTRADAS: números, CP, Tin ≠ Tout e identificadores únicos")',"Datos válidos","text")
    ws.merge_range("A8:K8","Fuente: PDF aportado · Synthesis of Heat Exchanger Networks · diapositivas 27 y 30. H1/H2/C3/C4 corresponden a 1/2/3/4.",fmt["sub"])
    head(ws,9,["ID","CP (kJ/h·K)","Tin (°C)","Tout (°C)","Tipo calculado","|Q| (kJ/h)","Q con signo","T baja","T alta","T* baja","T* alta"])
    for i,a in enumerate(s):
        rr=10+i; e=rr+1
        ws.write_row(rr,0,[a.id,a.cp,a.tin,a.tout],fmt["input"])
        f(ws,rr,4,f'=IF(C{e}>D{e},"HOT",IF(C{e}<D{e},"COLD","ERROR"))',"HOT" if a.hot else "COLD","text")
        f(ws,rr,5,f'=B{e}*ABS(C{e}-D{e})',a.duty)
        f(ws,rr,6,f'=B{e}*(C{e}-D{e})',a.cp*(a.tin-a.tout))
        f(ws,rr,7,f'=MIN(C{e},D{e})',a.low)
        f(ws,rr,8,f'=MAX(C{e},D{e})',a.high)
        f(ws,rr,9,f'=H{e}+IF(E{e}="HOT",-$B$4/2,$B$4/2)',a.low+(-r["dtmin"]/2 if a.hot else r["dtmin"]/2))
        f(ws,rr,10,f'=I{e}+IF(E{e}="HOT",-$B$4/2,$B$4/2)',a.high+(-r["dtmin"]/2 if a.hot else r["dtmin"]/2))
    ws.set_column("A:A",21); ws.freeze_panes(10,1); ws.autofilter(9,0,9+n,10)
    ws.data_validation("B4:B5",{"validate":"decimal","criteria":">=","value":0,"ignore_blank":False,"error_type":"stop","error_message":"Introduce un número no negativo."})
    ws.data_validation(f"B11:B{10+n}",{"validate":"decimal","criteria":">","value":0,"ignore_blank":False})
    ws.data_validation(f"C11:D{10+n}",{"validate":"decimal","criteria":">=","value":-273.15,"ignore_blank":False})
    ws.conditional_format(f"E11:E{10+n}",{"type":"text","criteria":"containing","value":"HOT","format":fmt["hot"]})
    ws.conditional_format(f"E11:E{10+n}",{"type":"text","criteria":"containing","value":"COLD","format":fmt["cold"]})
    ws.print_area(0,0,10+n,10)

    # 2. Cascada: el residual en cada fila corresponde a su temperatura B.
    title("Cascada","Tabla de intervalos y cascada","T*hot=T−ΔTmin/2; T*cold=T+ΔTmin/2. Los intervalos repetidos tienen anchura cero.")
    ws=sh["Cascada"]
    head(ws,2,["T* sin ordenar","T* descendente","ΔT intervalo","ΣCP hot","ΣCP cold","CP neto","ΔH intervalo","Residual bruto","Residual ajustado","T* pinch","Th pinch","Tc pinch","Posición"])
    rawts=[v for a in s for v in (a.low+(-r["dtmin"]/2 if a.hot else r["dtmin"]/2),a.high+(-r["dtmin"]/2 if a.hot else r["dtmin"]/2))]
    for k in range(m):
        rr=3+k; e=rr+1; src=11+k//2
        f(ws,rr,0,f"='Entradas'!{('J' if k%2==0 else 'K')}{src}",rawts[k])
        f(ws,rr,1,f'=LARGE($A$4:$A${m+3},{k+1})',float(r["levels"][k]))
        if k<m-1:
            f(ws,rr,2,f'=B{e}-B{e+1}',float(r["levels"][k]-r["levels"][k+1]))
            for c,t,cache in [(3,"HOT",r["hot_cp"][k]),(4,"COLD",r["cold_cp"][k])]:
                f(ws,rr,c,f'=SUMPRODUCT(--({ir("E")}="{t}"),--({ir("J")}<=(B{e}+B{e+1})/2),--({ir("K")} >=(B{e}+B{e+1})/2),{ir("B")})',float(cache))
            f(ws,rr,5,f'=D{e}-E{e}',float(r["hot_cp"][k]-r["cold_cp"][k]))
            f(ws,rr,6,f'=C{e}*F{e}',float(r["dh"][k]))
        f(ws,rr,7,'=0' if k==0 else f'=H{e-1}+G{e-1}',float(r["raw"][k]))
        f(ws,rr,8,f"=H{e}+'Resumen'!$B$7",float(r["residual"][k]))
        isp=abs(r["residual"][k])<=r["tolerance"]
        f(ws,rr,9,f'=IF(ABS(I{e})<=\'Resumen\'!$B$14,B{e},"")',float(r["levels"][k]) if isp else "")
        f(ws,rr,10,f'=IF(ISNUMBER(J{e}),J{e}+\'Entradas\'!$B$4/2,"")',float(r["levels"][k]+r["dtmin"]/2) if isp else "")
        f(ws,rr,11,f'=IF(ISNUMBER(J{e}),J{e}-\'Entradas\'!$B$4/2,"")',float(r["levels"][k]-r["dtmin"]/2) if isp else "")
        f(ws,rr,12,f'=IF(ISNUMBER(J{e}),IF(OR(B{e}=$B$4,B{e}=$B${m+3}),"umbral","interior"),"")',("umbral" if k in [0,m-1] else "interior") if isp else "","text")
    ws.print_area(0,0,m+2,12)

    # 3. Compuestas con m nodos de temperatura común y sus integrales.
    title("Compuestas","Integración de las curvas compuestas","Hhot inicia en 0. Hcold se desplaza QCmin. Segmentos verticales interiores indican huecos sin carga.")
    ws=sh["Compuestas"]
    head(ws,2,["T sin ordenar","T ascendente","T media","ΣCP hot","ΣCP cold","ΔH hot","ΔH cold","H hot","H cold origen","H cold ajustada","T hot gráfico","T cold gráfico","T* hot","T* cold","Índice tramo"])
    realraw=[v for a in s for v in (a.low,a.high)]
    for k in range(m):
        rr=3+k; e=rr+1; src=11+k//2
        f(ws,rr,0,f"='Entradas'!{('H' if k%2==0 else 'I')}{src}",realraw[k])
        f(ws,rr,1,f'=SMALL($A$4:$A${m+3},{k+1})',float(r["hot"]["t"][k]))
        if k<m-1:
            f(ws,rr,2,f'=(B{e}+B{e+1})/2',float((r["hot"]["t"][k]+r["hot"]["t"][k+1])/2))
            for c,t,cache in [(3,"HOT",r["hot"]["cp"][k]),(4,"COLD",r["cold"]["cp"][k])]:
                f(ws,rr,c,f'=SUMPRODUCT(--({ir("E")}="{t}"),--({ir("H")}<=C{e}),--({ir("I")}>=C{e}),{ir("B")})',float(cache))
            f(ws,rr,5,f'=D{e}*(B{e+1}-B{e})',float(r["hot"]["h"][k+1]-r["hot"]["h"][k]))
            f(ws,rr,6,f'=E{e}*(B{e+1}-B{e})',float(r["cold"]["h"][k+1]-r["cold"]["h"][k]))
            ws.write(rr,14,k+1)
        f(ws,rr,7,'=0' if k==0 else f'=H{e-1}+F{e-1}',float(r["hot"]["h"][k]))
        f(ws,rr,8,'=0' if k==0 else f'=I{e-1}+G{e-1}',float(r["cold"]["h"][k]-r["qc"]))
        f(ws,rr,9,f"=I{e}+'Resumen'!$B$8",float(r["cold"]["h"][k]))
        for c,energy,tshift,curve in [(10,"H",-1,r["hot"]),(11,"I",1,r["cold"])]:
            conditions=[]
            if k>0: conditions.append(f'{energy}{e}>{energy}{e-1}')
            if k<m-1: conditions.append(f'{energy}{e+1}>{energy}{e}')
            cond='OR('+','.join(conditions)+')'
            active=(k>0 and curve["h"][k]>curve["h"][k-1]) or (k<m-1 and curve["h"][k+1]>curve["h"][k])
            f(ws,rr,c,f'=IF({cond},B{e},"")',float(curve["t"][k]) if active else "")
            f(ws,rr,c+2,f'=IF(ISNUMBER({col(c)}{e}),{col(c)}{e}+{tshift}*\'Entradas\'!$B$4/2,"")',float(curve["t"][k]+tshift*r["dtmin"]/2) if active else "")
    ws.print_area(0,0,m+2,14)

    # 4. Unión de nodos H: temperaturas reales y fracciones en cada tramo.
    title("Tramos_red","Tramos de intercambio y división de corrientes","Cada tramo tiene un intervalo de temperaturas caliente y frío. Las fracciones CP definen las ramas.")
    ws=sh["Tramos_red"]
    baseheads=["H sin ordenar","H ordenada","ΔH tramo","H media","Índice hot","Índice cold","CP hot","CP cold","Th baja","Th alta","Tc baja","Tc alta","Q proceso","Q refrigeración","Q calentamiento","ΔT extremo 1","ΔT extremo 2","ΔT log media","Área proceso"]
    head(ws,2,baseheads+["fh "+a.id for a in s]+["fc "+a.id for a in s])
    unsorted=np.r_[r["hot"]["h"],r["cold"]["h"]]
    for k in range(2*m):
        rr=3+k; e=rr+1; sourcek=k if k<m else k-m
        f(ws,rr,0,f"='Compuestas'!{'H' if k<m else 'J'}{4+sourcek}",float(unsorted[k]))
        f(ws,rr,1,f'=SMALL($A$4:$A${2*m+3},{k+1})',float(r["energies"][k]))
        if k==z: continue
        b=r["blocks"][k]; hm=b["hot"]; cm=b["cold"]
        f(ws,rr,2,f'=B{e+1}-B{e}',b["q"])
        f(ws,rr,3,f'=(B{e}+B{e+1})/2',(b["elo"]+b["ehi"])/2)
        for c,energy,p in [(4,"H",hm),(5,"J",cm)]:
            f(ws,rr,c,f'=IF(C{e}<=\'Resumen\'!$B$14,0,SUMPRODUCT(--(\'Compuestas\'!${energy}$4:${energy}${m+2}<=D{e}),--(\'Compuestas\'!${energy}$5:${energy}${m+3}>D{e}),\'Compuestas\'!$O$4:$O${m+2}))',p["index"] if p else 0)
        for c,idx,cpcol,p in [(6,"E","D",hm),(7,"F","E",cm)]:
            f(ws,rr,c,f'=IF({idx}{e}>0,INDEX(\'Compuestas\'!${cpcol}$4:${cpcol}${m+2},{idx}{e}),0)',p["cp"] if p else 0)
        for c,idx,cpcol,hcol,p in [(8,"E","G","H",hm),(10,"F","H","J",cm)]:
            for delta in [0,1]:
                hpoint=f'B{e+delta}'
                form=f'=IF({idx}{e}>0,INDEX(\'Compuestas\'!$B$4:$B${m+2},{idx}{e})+({hpoint}-INDEX(\'Compuestas\'!${hcol}$4:${hcol}${m+2},{idx}{e}))/{cpcol}{e},0)'
                f(ws,rr,c+delta,form,p["low" if delta==0 else "high"] if p else 0)
        for c,condition,key in [(12,f'AND(E{e}>0,F{e}>0)',"process"),(13,f'AND(E{e}>0,F{e}=0)',"cooling"),(14,f'AND(E{e}=0,F{e}>0)',"heating")]:
            f(ws,rr,c,f'=IF({condition},C{e},0)',b[key])
        d1=hm["high"]-cm["high"] if hm and cm else 0
        d2=hm["low"]-cm["low"] if hm and cm else 0
        from pinch import lmtd
        lm=lmtd(d1,d2) if hm and cm else 0
        f(ws,rr,15,f'=IF(M{e}>0,J{e}-L{e},0)',d1)
        f(ws,rr,16,f'=IF(M{e}>0,I{e}-K{e},0)',d2)
        f(ws,rr,17,f'=IF(M{e}=0,0,IF(MIN(P{e},Q{e})<=0,0,IF(ABS(P{e}-Q{e})<0.00000001,(P{e}+Q{e})/2,(P{e}-Q{e})/LN(P{e}/Q{e}))))',lm)
        area=b["process"]/(r["u"]*lm) if r["u"] and lm>0 else "sin U" if not r["u"] else "infinita" if b["process"] else 0
        f(ws,rr,18,f'=IF(M{e}=0,0,IF(\'Entradas\'!$B$5=0,"sin U",IF(R{e}=0,"infinita",M{e}/(\'Entradas\'!$B$5*R{e}))))',area if b["process"] else 0)
        for i,a in enumerate(s):
            sr=11+i
            for c,kind,idx,low,high,cpv,val in [(19+i,"HOT","E","I","J","G",b["fh"][i]),(19+n+i,"COLD","F","K","L","H",b["fc"][i])]:
                form=f'=IF(AND({idx}{e}>0,\'Entradas\'!E{sr}="{kind}",\'Entradas\'!H{sr}<=({low}{e}+{high}{e})/2,\'Entradas\'!I{sr}>=({low}{e}+{high}{e})/2),\'Entradas\'!B{sr}/{cpv}{e},0)'
                f(ws,rr,c,form,float(val),"pct")
    ws.autofilter(2,0,z+2,18+2*n); ws.print_area(0,0,z+3,18)

    # 5. Matriz viva de cargas entre corrientes + servicios + balance.
    title("Matriz_Q","Intercambios entre corrientes","Q en kJ/h. Filas: donantes calientes. Columnas: receptoras frías. La matriz suma todos los tramos.")
    ws=sh["Matriz_Q"]
    head(ws,2,["Donante / receptora"]+[a.id for a in s]+["Refrigeración","Calentamiento","|Q| corriente","Error balance"])
    ws.set_column(0,0,23); ws.set_column(n+1,n+4,20)
    for j,a in enumerate(s): f(ws,2,j+1,f"='Entradas'!A{11+j}",a.id,"head")
    for i,a in enumerate(s):
        rr=3+i; e=rr+1
        f(ws,rr,0,f"='Entradas'!A{11+i}",a.id,"text")
        for j in range(n):
            f(ws,rr,j+1,f'=SUMPRODUCT({tr(12)},{tr(19+i)},{tr(19+n+j)})',float(r["matrix"][i,j]))
        f(ws,rr,n+1,f'=SUMPRODUCT({tr(13)},{tr(19+i)})',float(r["cooling"][i]))
        f(ws,rr,n+2,f'=SUMPRODUCT({tr(14)},{tr(19+n+i)})',float(r["heating"][i]))
        f(ws,rr,n+3,f"='Entradas'!F{11+i}",a.duty)
        f(ws,rr,n+4,f'=IF(\'Entradas\'!E{11+i}="HOT",SUM(B{e}:{col(n)}{e})+{col(n+1)}{e},SUM({col(i+1)}$4:{col(i+1)}${n+3})+{col(n+2)}{e})-{col(n+3)}{e}',r["balance_errors"][i])
    ws.conditional_format(3,1,n+2,n,{"type":"3_color_scale","min_color":"#F5F9FD","mid_color":"#8BBDEB","max_color":"#2275C9"})
    ws.conditional_format(3,n+4,n+2,n+4,{"type":"cell","criteria":"not between","minimum":-r["tolerance"]*10,"maximum":r["tolerance"]*10,"format":fmt["bad"]})
    ws.print_area(0,0,n+3,n+4)

    # 6. Equipos: incluye las combinaciones con Q=0, para que una edición pueda
    # activar parejas nuevas sin volver a Python. Filtrar Q>0 para leer la red.
    title("Equipos","Intercambiadores de ramas","Filtra Q > 0 para ver los equipos activos. Si cambias entradas, vuelve a aplicar el filtro. Las fracciones se refieren a cada corriente original.")
    ws=sh["Equipos"]
    head(ws,2,["ID candidato","Tramo","Caliente","Fría","Q (kJ/h)","Th entrada","Th salida","Tc entrada","Tc salida","Rama hot","Rama cold","CP rama hot","CP rama cold","ΔT extremo 1","ΔT extremo 2","ΔT log media","Área (m²)"])
    erow=3; eq_count=0
    for k,b in enumerate(r["blocks"]):
        trr=4+k
        for i,a in enumerate(s):
            for j,cold in enumerate(s):
                if i==j: continue
                e=erow+1; eq_count+=1
                ws.write(erow,0,f"B{k+1}-P{i+1}-{j+1}",fmt["text"]); ws.write(erow,1,k+1)
                f(ws,erow,2,f"='Entradas'!A{11+i}",a.id,"text")
                f(ws,erow,3,f"='Entradas'!A{11+j}",cold.id,"text")
                q=b["process"]*b["fh"][i]*b["fc"][j]
                hf=col(19+i); cf=col(19+n+j)
                f(ws,erow,4,f"='Tramos_red'!M{trr}*'Tramos_red'!{hf}{trr}*'Tramos_red'!{cf}{trr}",float(q))
                for c,source,key,which in [(5,"J","high","hot"),(6,"I","low","hot"),(7,"K","low","cold"),(8,"L","high","cold")]:
                    val=b[which][key] if q>r["tolerance"] else ""
                    f(ws,erow,c,f'=IF(E{e}>\'Resumen\'!$B$14,\'Tramos_red\'!{source}{trr},"")',val)
                f(ws,erow,9,f'=IF(E{e}>\'Resumen\'!$B$14,\'Tramos_red\'!{cf}{trr},0)',float(b["fc"][j]) if q>r["tolerance"] else 0,"pct")
                f(ws,erow,10,f'=IF(E{e}>\'Resumen\'!$B$14,\'Tramos_red\'!{hf}{trr},0)',float(b["fh"][i]) if q>r["tolerance"] else 0,"pct")
                f(ws,erow,11,f"=J{e}*'Entradas'!B{11+i}",a.cp*b["fc"][j] if q>r["tolerance"] else 0)
                f(ws,erow,12,f"=K{e}*'Entradas'!B{11+j}",cold.cp*b["fh"][i] if q>r["tolerance"] else 0)
                for c,src in [(13,"P"),(14,"Q"),(15,"R")]:
                    val=0
                    if q>r["tolerance"]:
                        d1=b["hot"]["high"]-b["cold"]["high"]; d2=b["hot"]["low"]-b["cold"]["low"]
                        val=d1 if c==13 else d2 if c==14 else lmtd(d1,d2)
                    f(ws,erow,c,f'=IF(E{e}>\'Resumen\'!$B$14,\'Tramos_red\'!{src}{trr},0)',val)
                area=q/(r["u"]*lmtd(d1,d2)) if q>r["tolerance"] and r["u"] and lmtd(d1,d2)>0 else "sin U" if q>r["tolerance"] and not r["u"] else "infinita" if q>r["tolerance"] else 0
                f(ws,erow,16,f'=IF(E{e}<=\'Resumen\'!$B$14,0,IF(\'Entradas\'!$B$5=0,"sin U",IF(P{e}=0,"infinita",E{e}/(\'Entradas\'!$B$5*P{e}))))',area)
                erow+=1
    ws.autofilter(2,0,erow-1,16)
    ws.conditional_format(3,4,erow-1,4,{"type":"cell","criteria":">","value":r["tolerance"],"format":wb.add_format({"bg_color":"#E3F2EE","bold":True})})
    ws.print_area(0,0,erow-1,16)

    # 7. Sensibilidad en Excel: cada escenario tiene su propia cascada visible.
    title("Sensibilidad","Efecto de ΔT mínimo","Escenarios recalculados mediante cascadas independientes. Cambia los ΔT azules o las corrientes de Entradas.")
    ws=sh["Sensibilidad"]
    head(ws,2,["ΔT mínimo (K)","QH mínimo","QC mínimo","Q recuperada"])
    from pinch import targets
    dtvalues=np.linspace(0,max(40,2*r["dtmin"]),9)
    for a,dt in enumerate(dtvalues):
        rr=3+a; e=rr+1; start=16+a*(m+4); first=start+2; last=first+m-1
        t=targets(s,float(dt)); ws.write(rr,0,float(dt),fmt["input"])
        f(ws,rr,1,f'=MAX(0,-MIN(H{first}:H{last}))',t["qh"])
        f(ws,rr,2,f'=H{last}+B{e}',t["qc"])
        f(ws,rr,3,f"='Resumen'!$B$5-C{e}",t["recovered"])
        ws.merge_range(start-1,0,start-1,7,f"Cascada del escenario {a+1} · ΔT en A{e}",fmt["head"])
        head(ws,start,["T* sin ordenar","T* ordenada","ΔT intervalo","CP hot","CP cold","CP neto","ΔH","Residual bruto"])
        for k in range(m):
            rw=start+1+k; ex=rw+1; sr=11+k//2; lowhigh="H" if k%2==0 else "I"
            shift=(-dt/2 if s[k//2].hot else dt/2)
            f(ws,rw,0,f'=\'Entradas\'!{lowhigh}{sr}+IF(\'Entradas\'!E{sr}="HOT",-$A${e}/2,$A${e}/2)',(s[k//2].low if k%2==0 else s[k//2].high)+shift)
            f(ws,rw,1,f'=LARGE($A${first}:$A${last},{k+1})',float(t["levels"][k]))
            if k<m-1:
                f(ws,rw,2,f'=B{ex}-B{ex+1}',float(t["levels"][k]-t["levels"][k+1]))
                for c,typ,sgn,cache in [(3,"HOT","-",t["hot_cp"][k]),(4,"COLD","+",t["cold_cp"][k])]:
                    form=f'=SUMPRODUCT(--({ir("E")}="{typ}"),--({ir("H")}{sgn}$A${e}/2<=(B{ex}+B{ex+1})/2),--({ir("I")}{sgn}$A${e}/2>=(B{ex}+B{ex+1})/2),{ir("B")})'
                    f(ws,rw,c,form,float(cache))
                f(ws,rw,5,f'=D{ex}-E{ex}',float(t["hot_cp"][k]-t["cold_cp"][k]))
                f(ws,rw,6,f'=C{ex}*F{ex}',float(t["dh"][k]))
            f(ws,rw,7,'=0' if k==0 else f'=H{ex-1}+G{ex-1}',float(t["raw"][k]))
    ws.data_validation("A4:A12",{"validate":"decimal","criteria":">=","value":0,"ignore_blank":False})
    ws.print_area("A1:P33")

    # 8. Resumen y gráficos nativos ligados a las fórmulas anteriores.
    title("Resumen","PINCH LAB","Integración energética · Cp constante · Contracorriente · División de corrientes por intervalos")
    ws=sh["Resumen"]; ws.set_column("A:A",31); ws.set_column("B:B",20); ws.set_column("C:C",13)
    ws.set_column("D:D",3); ws.set_column("E:T",10)
    f(ws,2,0,"='Entradas'!A7","Datos válidos","text")
    entries=[(4,"Calor disponible",f'=SUMIFS({ir("F")},{ir("E")},"HOT")',r["qhot"]),
             (5,"Demanda de calor",f'=SUMIFS({ir("F")},{ir("E")},"COLD")',r["qcold"]),
             (6,"Calentamiento mínimo",f'=MAX(0,-MIN(\'Cascada\'!H4:H{m+3}))',r["qh"]),
             (7,"Refrigeración mínima",f"='Cascada'!H{m+3}+B7",r["qc"]),
             (8,"Calor recuperado",'=B5-B8',r["recovered"]),
             (9,"Demanda cubierta",'=IF(B6=0,0,B9/B6)',r["recovered"]/r["qcold"] if r["qcold"] else 0),
             (10,"ΔT mínimo", "='Entradas'!B4",r["dtmin"]),
             (11,"Pinch Th (primer nodo)",f'=INDEX(\'Cascada\'!K4:K{m+3},MATCH(MIN(\'Cascada\'!I4:I{m+3}),\'Cascada\'!I4:I{m+3},0))',r["pinches"][0]["T_hot"]),
             (12,"Pinch Tc (primer nodo)",f'=INDEX(\'Cascada\'!L4:L{m+3},MATCH(MIN(\'Cascada\'!I4:I{m+3}),\'Cascada\'!I4:I{m+3},0))',r["pinches"][0]["T_cold"]),
             (13,"Tolerancia numérica",'=MAX(1,B5,B6)*0.000000001',r["tolerance"]),
             (14,"Balance global",'=B5+B7-B6-B8',0),
             (15,"Equipos de proceso",f'=COUNTIFS(\'Equipos\'!E4:E{erow},">"&B14)',len(r["exchangers"]))]
    for rr,label,formula,value in entries:
        ws.write(rr,0,label,fmt["text"]); f(ws,rr,1,formula,value,"pct" if rr==9 else "metric" if rr in [6,7,8] else "num")
        ws.write(rr,2,"kJ/h" if rr in [4,5,6,7,8,13,14] else "K" if rr==10 else "°C" if rr in [11,12] else "",fmt["sub"])
    ws.write_formula("B14","=MAX(1,B5,B6)*0.000000001",wb.add_format({"num_format":"0.00E+00","font_color":"#52657D"}),r["tolerance"])
    ws.write("A18","Conversión a kW",fmt["head"])
    for rr,label,src,val in [(18,"Calentamiento","B7",r["qh"]),(19,"Refrigeración","B8",r["qc"]),(20,"Recuperación","B9",r["recovered"])]:
        ws.write(rr,0,label,fmt["text"]); f(ws,rr,1,f'={src}/3600',val/3600); ws.write(rr,2,"kW",fmt["sub"])
    ws.merge_range("A24:C28","Para editar: Entradas, celdas azules. El pinch es un resultado. Encontrarás todos sus nodos en Cascada, columnas J:L. En umbrales o zonas pinch puede haber varios nodos.",fmt["sub"])
    ws.merge_range("A30:C34","Red ideal de máxima recuperación con ramas y mezclas a igual temperatura. No minimiza equipos ni costes. Área solo si introduces U. Servicios externos sin temperaturas especificadas: se calcula su carga, no su área.",fmt["sub"])
    ws.merge_range("A36:C40","Cambiar el número de corrientes requiere regenerar el Excel desde Python. En este libro sí puedes cambiar ID, CP, Tin, Tout, ΔT y U. No insertes filas dentro de los cálculos.",fmt["sub"])
    ws.merge_range("A42:C45","Calor sensible y CP constante por corriente. Para cp variable, discretiza por intervalos. No introducir cambios de fase como Tin=Tout.",fmt["sub"])
    def chart(title_,series,xlabel="Carga térmica (kJ/h)",ylabel="Temperatura (°C)"):
        ch=wb.add_chart({"type":"scatter","subtype":"straight_with_markers"})
        for series_ in series: ch.add_series(series_)
        ch.set_title({"name":title_,"name_font":{"name":"Aptos","size":13,"color":"#142D4E"}})
        ch.set_x_axis({"name":xlabel,"num_format":"#,##0","major_gridlines":{"visible":False},"name_font":{"size":10}})
        ch.set_y_axis({"name":ylabel,"num_format":"0","major_gridlines":{"visible":True,"line":{"color":"#E7EDF3"}},"name_font":{"size":10}})
        ch.set_legend({"position":"bottom","font":{"size":9}})
        ch.set_chartarea({"border":{"none":True}}); ch.set_plotarea({"border":{"none":True}})
        ch.set_size({"width":650,"height":370}); ch.show_blanks_as("gap")
        return ch
    def ser(name,sheet,x,y,count,start=3,color="#2275C9"):
        return {"name":name,"categories":[sheet,start,x,start+count-1,x],"values":[sheet,start,y,start+count-1,y],"line":{"color":color,"width":2.25},"marker":{"type":"circle","size":3,"border":{"color":color},"fill":{"color":color}}}
    for pos,title_,ycold,yhot,xcold in [("E4","Compuestas ajustadas",11,10,9),("N4","Compuestas desplazadas",13,12,9),("E23","Compuestas sin ajuste",11,10,8)]:
        ch=chart(title_,[ser("Caliente","Compuestas",7,yhot,m,color="#D64D44"),ser("Fría","Compuestas",xcold,ycold,m)],ylabel="T* (°C)" if "desplazadas" in title_ else "T (°C)")
        ws.insert_chart(pos,ch)
    ch=chart("Gran curva compuesta",[ser("Calor residual","Cascada",8,1,m,color="#148A83")],"Calor residual (kJ/h)","T* (°C)")
    ws.insert_chart("N23",ch)
    # Datos individuales en una zona visible de Compuestas debajo de la tabla.
    iw=sh["Compuestas"]; base=m+7
    head(iw,base,["Corriente","H local","T real"])
    individual=[]
    for i,a in enumerate(s):
        for k in [0,1]:
            rr=base+1+2*i+k; sr=11+i
            f(iw,rr,0,f"='Entradas'!A{sr}",a.id,"text")
            f(iw,rr,1,f'=IF(\'Entradas\'!E{sr}="HOT",{1-k},{k})*\'Entradas\'!F{sr}',a.duty*(1-k if a.hot else k))
            f(iw,rr,2,f"='Entradas'!{'C' if k==0 else 'D'}{sr}",a.tin if k==0 else a.tout)
        # Color de identidad, independiente del tipo: Tin/Tout pueden invertirlo.
        palette=["#142D4E","#95693C","#6D5AA0","#148A83"]
        se=ser(a.id,"Compuestas",1,2,2,base+1+2*i,palette[i%len(palette)])
        se["name"]=f"='Entradas'!$A${11+i}"; individual.append(se)
    ws.insert_chart("E42",chart("Curvas individuales",individual))
    sense=chart("Sensibilidad a ΔT mínimo",[ser("Calentamiento","Sensibilidad",0,1,9,color="#D64D44"),ser("Refrigeración","Sensibilidad",0,2,9),ser("Recuperación","Sensibilidad",0,3,9,color="#148A83")],"ΔT mínimo (K)","Carga térmica (kJ/h)")
    ws.insert_chart("N42",sense)
    ws.freeze_panes(3,0); ws.print_area("A1:W61"); ws.set_paper(8); ws.fit_to_pages(1,1)

    # 9. Método documentado dentro del propio libro.
    title("Guia","Cómo leer y modificar el modelo","Este archivo contiene fórmulas editables. No requiere ejecutar Python para cambiar valores de las corrientes existentes.")
    ws=sh["Guia"]; ws.set_column("A:A",27); ws.set_column("B:H",15)
    notes=[
      ("1. Datos","CP=m·cp en kJ/(h·K), temperaturas en °C. Q en kJ/h. Para trabajar en kW, divide CP y Q por 3600; este archivo conserva las unidades originales."),
      ("2. ΔT mínimo y pinch","Eliges ΔT mínimo; el pinch se obtiene de los ceros de la cascada ajustada. No es una temperatura de entrada independiente. Revisa Cascada J:L."),
      ("3. Cascada","Ordenar T*; ΔH=(ΣCP_hot−ΣCP_cold)ΔT*. QHmin=max(0,−min residual bruto); QCmin=residual final+QHmin."),
      ("4. Curvas","Compuestas integra CP por intervalos. La curva fría se desplaza QCmin sobre el eje H. La gran compuesta representa el residual de la cascada frente a T*."),
      ("5. Red por tramos","Se unen los nodos de entalpía de ambas curvas. En cada intervalo se calculan Th y Tc en los dos extremos. Se asigna Qij=ΔH·fh_i·fc_j."),
      ("6. Fracciones de rama","fh_i=CP_i/ΣCP_hot; fc_j=CP_j/ΣCP_cold. La rama caliente i hacia j usa fracción fc_j. La rama fría j desde i usa fracción fh_i. Consultar Equipos J:M."),
      ("7. Servicios externos","Las colas no solapadas determinan refrigeración y calentamiento. La matriz muestra servicios por corriente. Elegir utilidades reales exige temperaturas compatibles con ΔT mínimo."),
      ("8. Área opcional","A=Q/(U·ΔTlm). U es un supuesto del usuario y debe ser compatible con las unidades. ΔTlm=(ΔT1−ΔT2)/ln(ΔT1/ΔT2). Si ambas diferencias son iguales se usa ese valor. Con ΔT extremo cero, área infinita."),
      ("9. Modificaciones","Cambiar celdas azules en Entradas actualiza tablas, equipos y gráficos. Para más/menos corrientes, regenerar desde Python. Si aplicaste filtros, vuelve a aplicarlos tras modificar los datos."),
      ("10. Límites","CP constante, régimen estacionario, calor sensible, sin pérdidas térmicas ni restricciones de mezcla de ramas de una misma corriente. La red minimiza servicios, no número de equipos ni coste total."),
      ("11. Excel","Fórmulas convencionales; gráficos XY nativos. Usa cálculo automático. En Excel de escritorio puedes forzar recálculo con Ctrl+Alt+F9. Vista previa sin motor de cálculo puede mostrar solo la caché inicial."),
      ("12. Fuente","Datos y ΔT mínimo: PDF del usuario, diapositivas 27/30. Método de cascada y temperaturas desplazadas explicado en el código y en LEEME.md. No se supone un U ni temperaturas de servicios no aportados."),
    ]
    for i,(a,b) in enumerate(notes):
        rr=3+i*3; ws.merge_range(rr,0,rr+1,0,a,fmt["head"]); ws.merge_range(rr,1,rr+1,7,b,fmt["sub"])
        ws.set_row(rr,25); ws.set_row(rr+1,25)
    ws.print_area(0,0,3+len(notes)*3,7)
    wb.close()
    return sink.getvalue() if destination is None else destination
