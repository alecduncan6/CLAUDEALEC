#!/usr/bin/env python3
"""Lee el Excel de pedidos del almacén y lo resume por día de pedido.

    python contabilidad/importar_almacen.py <export_almacen.xlsx> [--excluir 3503014,...] [--json salida.json]
                                            [--incidencias incidencias.json]

Solo saca totales por día (nunca nombres, teléfonos ni direcciones):
    p1, p2            pedidos confirmados de cada pack
    ent1, ent2        entregados de cada pack
    dev1, dev2        devueltos de cada pack
    inc               en incidencia (siguen pendientes hasta que pasen a Devuelto)
    envio_ent         coste de envío real (sin IVA) de los entregados
    envio_pend        coste de envío real (sin IVA) de los pendientes (preparado, en ruta, incidencia)
    carritos          carritos abandonados (no son pedidos)
    cancelados        rechazados / cancelados antes de enviarse
"""
import argparse
import datetime as dt
import json
from collections import defaultdict

from openpyxl import load_workbook

ENTREGADO = {"entregado"}
DEVUELTO = {"devuelto", "devolución", "devolucion"}
INCIDENCIA = {"incidencia"}
PENDIENTE = {"confirmado", "pedido confirmado", "preparado", "en ruta", "en reparto", "pendiente"}
CANCELADO = {"rechazado", "cancelado", "anulado"}
CARRITO = {"carrito abandonado"}

ap = argparse.ArgumentParser()
ap.add_argument("export")
ap.add_argument("--excluir", default="", help="IDs de pedido a ignorar (tests), separados por comas")
ap.add_argument("--json", help="guardar el resumen en este fichero")
ap.add_argument("--incidencias", help="guardar los pedidos por estado para el tablero de incidencias "
                                      "(lleva nombre y teléfono: solo al scratchpad, nunca al repo)")
args = ap.parse_args()
excluir = {s.strip() for s in args.excluir.split(",") if s.strip()}

ws = load_workbook(args.export, read_only=True, data_only=True).worksheets[0]
filas = ws.iter_rows(values_only=True)
cab = [str(c).strip() if c is not None else "" for c in next(filas)]
col = {nombre: i for i, nombre in enumerate(cab)}
for necesario in ("ESTADO", "UD", "PRECIO", "NOMBRE COMPLETO", "ID PEDIDO", "COSTE DE ENVÍO (SIN IVA)",
                  "Fecha del pedido", "¿Es carrito?"):
    if necesario not in col:
        raise SystemExit(f"Falta la columna '{necesario}' en el export del almacén")


def fecha_iso(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    d, m, y = str(v).strip().split("-")
    return f"{y}-{m}-{d}"


dias = defaultdict(lambda: defaultdict(float))
ignorados, avisos = [], []
estados = {}  # ID pedido -> datos para el tablero de incidencias
for f in filas:
    if not f or f[col["ID PEDIDO"]] is None:
        continue
    pid = str(f[col["ID PEDIDO"]]).strip()
    estado = str(f[col["ESTADO"]] or "").strip().lower()
    nombre = str(f[col["NOMBRE COMPLETO"]] or "").strip().lower()
    dia = dias[fecha_iso(f[col["Fecha del pedido"]])]
    if pid in excluir or nombre.startswith("test"):
        ignorados.append(pid)
        continue
    if estado in CARRITO or str(f[col["¿Es carrito?"]]).strip() == "1":
        dia["carritos"] += 1
        continue
    if estado in CANCELADO:
        dia["cancelados"] += 1
        continue
    uds = int(f[col["UD"]] or 0)
    pack = {1: "1", 2: "2"}.get(uds)
    if pack is None:
        avisos.append(f"pedido {pid}: {uds} uds no encaja con ningún pack (se cuenta como Pack 2)")
        pack = "2"
    estados[pid] = {
        "pedido": pid, "estado_almacen": str(f[col["ESTADO"]] or "").strip(), "fecha_pedido": fecha_iso(f[col["Fecha del pedido"]]),
        "pack": int(pack), "pack_nombre": "Pack 1 ud" if pack == "1" else "Pack 2 uds", "importe": float(f[col["PRECIO"]] or 0),
        "tracking": str(f[col["TRACKING"]] or "").strip() if "TRACKING" in col else "",
        "nombre": str(f[col["NOMBRE COMPLETO"]] or "").strip(),
        "telefono": str(f[col["TELF"]] or "").strip() if "TELF" in col else "",
        "ciudad": str(f[col["CIUDAD"]] or "").strip() if "CIUDAD" in col else "",
        "motivo": str(f[col["Motivo incidencia (si tiene)"]] or "").strip() if "Motivo incidencia (si tiene)" in col else "",
    }
    envio = float(f[col["COSTE DE ENVÍO (SIN IVA)"]] or 0)
    dia[f"p{pack}"] += 1
    if estado in ENTREGADO:
        dia[f"ent{pack}"] += 1
        dia["envio_ent"] += envio
    elif estado in DEVUELTO:
        dia[f"dev{pack}"] += 1
    else:
        if estado in INCIDENCIA:
            dia["inc"] += 1
        elif estado not in PENDIENTE:
            avisos.append(f"pedido {pid}: estado '{estado}' desconocido (se cuenta como pendiente)")
        dia["envio_pend"] += envio

CAMPOS = ("p1", "p2", "ent1", "ent2", "dev1", "dev2", "inc", "envio_ent", "envio_pend", "carritos", "cancelados")
resumen = {}
for fecha in sorted(dias):
    d = dias[fecha]
    resumen[fecha] = {k: (round(d[k], 2) if k.startswith("envio") else int(d[k])) for k in CAMPOS}

print(f"{'fecha':<11}{'P1':>4}{'P2':>4}{'Ent':>5}{'Dev':>5}{'Inc':>5}{'Pend':>6}{'Envío ent':>11}{'Envío pend':>12}{'Carr':>6}{'Canc':>6}")
for fecha, d in resumen.items():
    pend = d["p1"] + d["p2"] - d["ent1"] - d["ent2"] - d["dev1"] - d["dev2"]
    print(f"{fecha:<11}{d['p1']:>4}{d['p2']:>4}{d['ent1'] + d['ent2']:>5}{d['dev1'] + d['dev2']:>5}{d['inc']:>5}"
          f"{pend:>6}{d['envio_ent']:>11.2f}{d['envio_pend']:>12.2f}{d['carritos']:>6}{d['cancelados']:>6}")
print(f"Ignorados (test): {', '.join(ignorados) or 'ninguno'}")
for a in avisos:
    print("AVISO:", a)
if args.incidencias:
    with open(args.incidencias, "w", encoding="utf-8") as fh:
        json.dump(estados, fh, ensure_ascii=False, indent=2)
if args.json:
    with open(args.json, "w", encoding="utf-8") as fh:
        json.dump(resumen, fh, ensure_ascii=False, indent=2)
