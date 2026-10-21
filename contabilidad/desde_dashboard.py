#!/usr/bin/env python3
"""Convierte una exportación de la base de datos del dashboard en datos.json.

    python contabilidad/desde_dashboard.py <carpeta_export> [datos.json]

<carpeta_export> es el out_dir de ArtifactData (list/get con out_dir), con:
    config/main.json, dias/<fecha>.json, gastos/<id>.json
    y, si está, incidencias/<pedido>.json (solo se guardan los recuentos de cerradas, nada de clientes)
"""
import datetime as dt
import json
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
origen = Path(sys.argv[1])
destino = Path(sys.argv[2]) if len(sys.argv) > 2 else AQUI / "datos.json"


def leer(p):
    return json.loads(p.read_text(encoding="utf-8"))


previo = leer(destino) if destino.exists() else {}
config = leer(origen / "config" / "main.json") if (origen / "config" / "main.json").exists() else previo.get("config", {})
dias = sorted((leer(p) for p in (origen / "dias").glob("*.json")), key=lambda d: d["fecha"])
dias = [d.get("data", d) if "fecha" not in d else d for d in dias]
gastos = [{"id": p.stem, **leer(p)} for p in sorted((origen / "gastos").glob("*.json"))]

# Mismo cálculo que el dashboard: incidencias cerradas y % de pedidos que acaban en incidencia (q), corrigiendo los días
# recientes con la parte de incidencias que ya se ve a los N días laborables del pedido (INC_CDF).
INC_CDF = [0.14, 0.30, 0.70, 0.92, 0.98]
FESTIVOS = {"2026-10-12", "2026-12-08", "2026-12-25", "2027-01-01", "2027-01-06", "2027-03-26", "2027-10-12", "2027-11-01",
            "2027-12-06", "2027-12-08"}


def lag_hab(desde, hasta):
    d, n = dt.date.fromisoformat(desde), 0
    while d < dt.date.fromisoformat(hasta):
        d += dt.timedelta(1)
        n += d.weekday() < 5 and d.isoformat() not in FESTIVOS
    return n


def leer_doc(p):
    d = leer(p)
    return d.get("data", d) if isinstance(d, dict) and "data" in d and isinstance(d["data"], dict) else d


incidencias = previo.get("incidencias", {})
if (origen / "incidencias").is_dir():
    todas = [leer_doc(p) for p in (origen / "incidencias").glob("*.json")]
    cerr = [i for i in todas if not i.get("abierta") and i.get("resultado") in ("entregado", "devuelto")]
    por_dia = {}
    for i in todas:
        if i.get("fecha_pedido"):
            por_dia[i["fecha_pedido"]] = por_dia.get(i["fecha_pedido"], 0) + 1
    num = den = 0.0
    for d in dias:
        hasta = (d.get("actualizado") or dt.date.today().isoformat())[:10]
        a = lag_hab(d["fecha"], hasta) if d["fecha"] < hasta else 0
        f = 1 if a >= len(INC_CDF) else INC_CDF[a]
        if f < 0.7:
            continue
        num += max(por_dia.get(d["fecha"], 0), d.get("inc") or 0)
        den += sum(d.get(f"p{n}") or 0 for n in (1, 2, 3)) * f
    # Igual que el panel: una incidencia abierta en el tablero sigue siéndolo aunque el almacén la haya vuelto a poner "en ruta"
    abiertas = {}
    for i in todas:
        if i.get("abierta") and i.get("fecha_pedido"):
            abiertas[i["fecha_pedido"]] = abiertas.get(i["fecha_pedido"], 0) + 1
    for d in dias:
        pend = sum((d.get(f"p{n}") or 0) - (d.get(f"ent{n}") or 0) - (d.get(f"dev{n}") or 0) for n in (1, 2, 3))
        if abiertas.get(d["fecha"], 0) > (d.get("inc") or 0):
            d["inc"] = min(max(0, pend), abiertas[d["fecha"]])
    incidencias = {"cerradas": len(cerr), "entregadas": sum(i["resultado"] == "entregado" for i in cerr),
                   "q": round(num / den, 4) if den >= 20 else None,
                   "fecha_export": max((d.get("actualizado") or "")[:10] for d in dias) or None}

destino.write_text(json.dumps({"config": config, "dias": dias, "gastos": gastos, "incidencias": incidencias}, ensure_ascii=False, indent=2) + "\n",
                   encoding="utf-8")
print(f"{destino}: {len(dias)} días, {len(gastos)} gastos, incidencias {incidencias}")
