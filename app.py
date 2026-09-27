"""Interfaz web LOCAL de Pinch Lab, sin servidores externos ni cuenta.

Ejecutar desde esta carpeta: python app.py
La interfaz se abre en http://127.0.0.1:8765. Para detenerla, Ctrl+C.
Solo usa la biblioteca estándar para HTTP; el cálculo está en pinch.py.
"""
from __future__ import annotations
import argparse
import base64
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
from pathlib import Path
import threading
import webbrowser

import matplotlib.pyplot as plt
from pinch import Stream, EXAMPLE, analyse, tables, figures
from excel_export import export_excel

ROOT=Path(__file__).resolve().parent
LOCK=threading.Lock()  # Matplotlib no es seguro con renderizados simultáneos.


class Handler(BaseHTTPRequestHandler):
    def send(self,body,content_type="application/json; charset=utf-8",status=200,filename=None):
        self.send_response(status)
        self.send_header("Content-Type",content_type)
        self.send_header("Content-Length",str(len(body)))
        self.send_header("Cache-Control","no-store")
        if filename: self.send_header("Content-Disposition",f'attachment; filename="{filename}"')
        self.end_headers(); self.wfile.write(body)

    def do_GET(self):
        if self.path=="/":
            self.send((ROOT/"interfaz.html").read_bytes(),"text/html; charset=utf-8")
        elif self.path=="/example":
            self.send(json.dumps({"streams":[dict(id=s.id,cp=s.cp,tin=s.tin,tout=s.tout) for s in EXAMPLE],"dtmin":10,"u":0}).encode())
        else:
            self.send(b'{"error":"No encontrado"}',status=404)

    def do_POST(self):
        try:
            length=int(self.headers.get("Content-Length","0"))
            if length<=0 or length>100000:
                raise ValueError("Solicitud vacía o demasiado grande.")
            # El servidor solo escucha en loopback. Comprobar origen evita que
            # páginas ajenas lancen trabajos desde el navegador del usuario.
            origin=self.headers.get("Origin")
            allowed=f"http://127.0.0.1:{self.server.server_port}"
            if origin and origin not in (allowed,f"http://localhost:{self.server.server_port}"):
                raise ValueError("Origen no permitido.")
            data=json.loads(self.rfile.read(length))
            rows=data["streams"]
            if not 1<=len(rows)<=20:
                raise ValueError("Introduce entre 1 y 20 corrientes.")
            streams=[Stream(str(s["id"]).strip(),float(s["cp"]),float(s["tin"]),float(s["tout"])) for s in rows]
            r=analyse(streams,float(data["dtmin"]),float(data.get("u",0)))
            if self.path=="/excel":
                self.send(export_excel(r),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",filename="Pinch_Analisis.xlsx")
                return
            if self.path!="/calculate":
                self.send(b'{"error":"No encontrado"}',status=404); return
            with LOCK:
                image_data={}
                for name,fig in figures(r).items():
                    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=115); plt.close(fig)
                    image_data[name]="data:image/png;base64,"+base64.b64encode(buf.getvalue()).decode()
            tab={name:df.to_html(index=name=="matriz_Q",classes="data-table",border=0,
                                float_format=lambda x:f"{x:,.3f}",na_rep="No calculada",escape=True)
                 for name,df in tables(r).items()}
            payload={"summary":{k:r[k] for k in ["qh","qc","qhot","qcold","recovered","dtmin","pinches"]},
                     "equipment_count":len(r["exchangers"]),"images":image_data,"tables":tab,
                     "balance_error":max(abs(v) for v in r["balance_errors"])}
            self.send(json.dumps(payload,ensure_ascii=False,allow_nan=False).encode("utf-8"))
        except (KeyError,TypeError,ValueError,ArithmeticError) as exc:
            self.send(json.dumps({"error":str(exc)},ensure_ascii=False).encode("utf-8"),status=400)
        except Exception as exc:
            self.send(json.dumps({"error":f"Error de cálculo: {type(exc).__name__}: {exc}"}).encode(),status=500)


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--port",type=int,default=8765)
    parser.add_argument("--no-browser",action="store_true")
    args=parser.parse_args()
    server=ThreadingHTTPServer(("127.0.0.1",args.port),Handler)
    url=f"http://127.0.0.1:{args.port}"
    print(f"Pinch Lab: {url}\nPara detenerlo: Ctrl+C.",flush=True)
    if not args.no_browser:
        webbrowser.open(url)
    try: server.serve_forever()
    except KeyboardInterrupt: pass
    finally: server.server_close()


if __name__=="__main__": main()
