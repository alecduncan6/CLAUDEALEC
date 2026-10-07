#!/usr/bin/env python3
"""Convierte una exportación de la base de datos del dashboard en datos.json.

    python contabilidad/desde_dashboard.py <carpeta_export> [datos.json]

<carpeta_export> es el out_dir de ArtifactData (list/get con out_dir), con:
    config/main.json, dias/<fecha>.json, gastos/<id>.json
"""
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
gastos = [{"id": p.stem, **leer(p)} for p in sorted((origen / "gastos").glob("*.json"))]

destino.write_text(json.dumps({"config": config, "dias": dias, "gastos": gastos}, ensure_ascii=False, indent=2) + "\n",
                   encoding="utf-8")
print(f"{destino}: {len(dias)} días, {len(gastos)} gastos")
