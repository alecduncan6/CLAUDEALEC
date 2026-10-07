# CeraLux · sistema de números

Marca de e-commerce de Alec: spray reparación de arañazos 120ml, venta contra reembolso (COD) en España.
Ads en Meta. Almacén/plataforma COD con estados: pedido confirmado → en ruta → entregado / devuelto.

## Piezas

- **Dashboard en vivo (fuente de verdad):** https://claude.ai/artifact/A45KKiFp9hR3ZavkXX1be2
  - Fuente: `dashboard/ceralux_numeros.html` (republicar con el tool Artifact pasando `url`).
  - Base de datos (ArtifactData, mismo `url`):
    - `config/main`: precios, costes y supuestos (mismas claves que `contabilidad/datos.json` → `config`).
    - `dias/<YYYY-MM-DD>`: `{fecha, p1, p2, ent, dev, ads, nota, actualizado}` por **fecha de pedido**.
      `p1`/`p2` = pedidos confirmados de cada pack (null = aún no reportado). `ent`/`dev` = de ESOS pedidos,
      cuántos entregados / devueltos. En ruta se calcula (p1+p2−ent−dev).
    - `gastos/<id>`: `{fecha, concepto, categoria, importe}`.
  - La página lee el gasto de Meta en directo (conector "Meta ADS", tool `ads_get_ad_entities`).
- **Excel de contabilidad:** `contabilidad/CeraLux_Contabilidad.xlsx`, generado por `contabilidad/generar_excel.py`
  desde `contabilidad/datos.json`. Todo con fórmulas (Dashboard, Diario, Gastos, Config, Guía).

## Meta Ads

- Cuenta: HeatioShop, `ad_account_id` **26290722800581380**. Solo cuentan campañas cuyo nombre contiene **"CeraLux"**.
- Gasto diario: `ads_get_ad_entities` con `level: "campaign"`, `fields: ["amount_spent"]`,
  `filtering: [{"field":"campaign.name","operator":"CONTAIN","value":["CeraLux"]}]`,
  `time_range: {"since":..,"until":..}`, `time_increment: "1"`. Sumar `amount_spent.value` por `date_start`.

## Rutina diaria (cuando Alec diga "hoy X pack 1, Y pack 2")

1. Saca el gasto de Meta de ese día (consulta de arriba).
2. `ArtifactData get dias/<fecha>`; luego `set` (con `if_version` si existe) conservando `ent`/`dev`/`nota` previos:
   `{fecha, p1, p2, ent, dev, ads, nota, actualizado}`.
3. Responde con el resumen del día: pedidos, gasto, CPA vs CPA break-even (~13 € con supuestos actuales),
   beneficio proyectado y semáforo (ESCALAR ≤ 75% del break-even, VIGILAR ≤ 100%, CORTAR > 100%).

Actualización de estados ("del 08/10: 11 entregados, 2 devueltos") → `update dias/<fecha>` con `ent`/`dev`.
Comprueba que ent + dev ≤ p1 + p2.

## Regenerar el Excel con los datos al día

1. `ArtifactData list` de `dias` y `gastos` y `get config/main`, todos con `out_dir=<scratchpad>/export`.
2. `python contabilidad/desde_dashboard.py <scratchpad>/export`
3. `python contabilidad/generar_excel.py` y recalcular con el `recalc.py` del skill xlsx (0 errores).
4. Commit + push.

## Modelo (igual en Excel y dashboard)

- Coste pedido = uds × coste_ud × 1,21 + logística (8,06 €) → 10,47 € (Pack 1) / 12,88 € (Pack 2), cuadra con el panel.
- Beneficio por entregado = PVP − coste (19,52 € / 27,11 €). IVA neto a liquidar por entregado ≈ 3,39 € / 4,71 €.
- Devuelto = −coste_devolucion (8,06 €, supuesto) + IVA recuperable.
- Proyección: en ruta × tasa de entrega (80% estimada hasta 30 pedidos resueltos; luego la real).

Supuestos pendientes de confirmar con Alec: coste real de una devolución, si Meta le cobra IVA (ROI), si factura con IVA.
