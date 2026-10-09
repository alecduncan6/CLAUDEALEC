#!/usr/bin/env python3
"""Lee el Excel de pedidos del almacén y lo resume por día de pedido.

    python contabilidad/importar_almacen.py <export_almacen.xlsx> [--excluir 3503014,...] [--json salida.json]
                                            [--incidencias incidencias.json]

Solo saca totales por día (nunca nombres, teléfonos ni direcciones):
    p1, p2, p3        pedidos confirmados de cada pack (1, 2 y 3 unidades; 4 o más cuentan como Pack 3)
    ent1, ent2, ent3  entregados de cada pack
    dev1, dev2, dev3  devueltos de cada pack
    inc               en incidencia (siguen pendientes hasta que pasen a Devuelto)
    envio_ent         coste de envío real (sin IVA) de los entregados
    envio_pend        coste de envío real (sin IVA) de los pendientes (preparado, en ruta, incidencia);
                      los que aún no tienen envío en el export (0) cuentan con --envio-std
    shopify_ids       IDs de Shopify de todas las filas del día (para no contar dos veces los pedidos de Shopify)
    carritos          carritos abandonados sin recuperar (no son pedidos)
    recuperados       carritos que el equipo ha recuperado: ya tienen estado de pedido y cuentan en p1/p2
    cancelados        rechazados / cancelados antes de enviarse
"""
import argparse
import datetime as dt
import json
from collections import defaultdict

from openpyxl import load_workbook

# El almacén añade sufijos al estado ("Rehusado - en tránsito", "Confirmado - Pendiente de preparación"): se compara
# por prefijo, igual que la página. "Rehusado" = el cliente lo rechazó en la puerta y vuelve: es devolución.
# "Rechazado" = anulado antes de salir del almacén: no es pedido.
ENTREGADO = {"entregado"}
DEVUELTO = {"devuelto", "devolución", "devolucion", "rehusado"}
INCIDENCIA = {"incidencia"}
PENDIENTE = {"confirmado", "pedido confirmado", "pedido nuevo", "preparado", "enviado", "en ruta", "en reparto",
             "en tránsito", "en transito", "pendiente"}
CANCELADO = {"rechazado", "cancelado", "anulado"}


def es(lista, e):
    return any(e == k or e.startswith(k + " ") or e.startswith(k + "-") for k in lista)


ap = argparse.ArgumentParser()
ap.add_argument("export")
ap.add_argument("--excluir", default="", help="IDs de pedido a ignorar (tests), separados por comas")
ap.add_argument("--json", help="guardar el resumen en este fichero")
ap.add_argument("--incidencias", help="guardar los pedidos por estado para el tablero de incidencias "
                                      "(lleva nombre y teléfono: solo al scratchpad, nunca al repo)")
ap.add_argument("--envio-std", type=float, default=8.06,
                help="envío (sin IVA) para los pedidos que aún no tienen coste de envío en el export (salen a 0)")
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
ids_shopify = defaultdict(list)  # IDs de Shopify que conoce el almacén (incluye tests y cancelados)
ids_test = defaultdict(list)     # IDs de Shopify de los pedidos de test
for f in filas:
    if not f or f[col["ID PEDIDO"]] is None:
        continue
    pid = str(f[col["ID PEDIDO"]]).strip()
    estado = str(f[col["ESTADO"]] or "").strip().lower()
    nombre = str(f[col["NOMBRE COMPLETO"]] or "").strip().lower()
    dia = dias[fecha_iso(f[col["Fecha del pedido"]])]
    if "ID pedido Shopify" in col and f[col["ID pedido Shopify"]]:
        ids_shopify[fecha_iso(f[col["Fecha del pedido"]])].append(str(f[col["ID pedido Shopify"]]).strip())
    if pid in excluir or nombre.startswith("test"):
        ignorados.append(pid)
        if "ID pedido Shopify" in col and f[col["ID pedido Shopify"]]:
            ids_test[fecha_iso(f[col["Fecha del pedido"]])].append(str(f[col["ID pedido Shopify"]]).strip())
        continue
    # "¿Es carrito?" = 1 marca el origen; si el equipo lo recupera, Dropi le pone estado de pedido y cuenta como venta
    es_carrito = str(f[col["¿Es carrito?"]]).strip() == "1"
    if "carrito" in estado or (es_carrito and es(CANCELADO, estado)):
        dia["carritos"] += 1
        continue
    if es(CANCELADO, estado):
        dia["cancelados"] += 1
        continue
    uds = int(f[col["UD"]] or 0)
    pack = {1: "1", 2: "2", 3: "3"}.get(uds)
    if pack is None:
        pack = "1" if uds < 1 else "3"
        avisos.append(f"pedido {pid}: {uds} uds no encaja con ningún pack (se cuenta como Pack {pack})")
    estados[pid] = {
        "pedido": pid, "estado_almacen": str(f[col["ESTADO"]] or "").strip(), "fecha_pedido": fecha_iso(f[col["Fecha del pedido"]]),
        "pack": int(pack), "pack_nombre": "Pack 1 ud" if pack == "1" else f"Pack {pack} uds", "importe": float(f[col["PRECIO"]] or 0),
        "tracking": str(f[col["TRACKING"]] or "").strip() if "TRACKING" in col else "",
        "nombre": str(f[col["NOMBRE COMPLETO"]] or "").strip(),
        "telefono": str(f[col["TELF"]] or "").strip() if "TELF" in col else "",
        "ciudad": str(f[col["CIUDAD"]] or "").strip() if "CIUDAD" in col else "",
        "motivo": str(f[col["Motivo incidencia (si tiene)"]] or "").strip() if "Motivo incidencia (si tiene)" in col else "",
    }
    # hasta que no sale del almacén el envío viene a 0: se usa la tarifa estándar para no inflar el beneficio
    envio = float(f[col["COSTE DE ENVÍO (SIN IVA)"]] or 0) or args.envio_std
    dia[f"p{pack}"] += 1
    if es_carrito:
        dia["recuperados"] += 1
    if es(ENTREGADO, estado):
        dia[f"ent{pack}"] += 1
        dia["envio_ent"] += envio
    elif es(DEVUELTO, estado):
        dia[f"dev{pack}"] += 1
    else:
        if es(INCIDENCIA, estado):
            dia["inc"] += 1
        elif not es(PENDIENTE, estado):
            avisos.append(f"pedido {pid}: estado '{estado}' desconocido (se cuenta como pendiente)")
        dia["envio_pend"] += envio

CAMPOS = ("p1", "p2", "p3", "ent1", "ent2", "ent3", "dev1", "dev2", "dev3", "inc", "envio_ent", "envio_pend", "carritos", "recuperados", "cancelados")
resumen = {}
for fecha in sorted(dias):
    d = dias[fecha]
    resumen[fecha] = {k: (round(d[k], 2) if k.startswith("envio") else int(d[k])) for k in CAMPOS}
    resumen[fecha]["shopify_ids"] = ids_shopify.get(fecha, [])
    resumen[fecha]["shopify_test"] = ids_test.get(fecha, [])

print(f"{'fecha':<11}{'P1':>4}{'P2':>4}{'P3':>4}{'Ent':>5}{'Dev':>5}{'Inc':>5}{'Pend':>6}{'Envío ent':>11}{'Envío pend':>12}{'Carr':>6}{'Recu':>6}{'Canc':>6}")
for fecha, d in resumen.items():
    pend = sum(d[f"p{n}"] - d[f"ent{n}"] - d[f"dev{n}"] for n in (1, 2, 3))
    print(f"{fecha:<11}{d['p1']:>4}{d['p2']:>4}{d['p3']:>4}{d['ent1'] + d['ent2'] + d['ent3']:>5}"
          f"{d['dev1'] + d['dev2'] + d['dev3']:>5}{d['inc']:>5}"
          f"{pend:>6}{d['envio_ent']:>11.2f}{d['envio_pend']:>12.2f}{d['carritos']:>6}{d['recuperados']:>6}{d['cancelados']:>6}")
print(f"Ignorados (test): {', '.join(ignorados) or 'ninguno'}")
for a in avisos:
    print("AVISO:", a)
if args.incidencias:
    with open(args.incidencias, "w", encoding="utf-8") as fh:
        json.dump(estados, fh, ensure_ascii=False, indent=2)
if args.json:
    with open(args.json, "w", encoding="utf-8") as fh:
        json.dump(resumen, fh, ensure_ascii=False, indent=2)
