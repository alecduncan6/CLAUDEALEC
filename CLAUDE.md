# CeraLux · sistema de números

Marca de e-commerce de Alec: spray reparación de arañazos 120ml, venta contra reembolso (COD) en España.
Ads en Meta. Almacén COD con estados: Preparado → En ruta → Entregado / Devuelto, más INCIDENCIA
(no se pudo entregar a la primera; sigue pendiente, NO es venta perdida hasta que pasa a Devuelto),
Rechazado/Cancelado (no cuenta como pedido) y Carrito Abandonado (no es pedido).

## Piezas

- **Dashboard en vivo (fuente de verdad):** https://claude.ai/artifact/A45KKiFp9hR3ZavkXX1be2
  - Fuente: `dashboard/ceralux_numeros.html` (republicar con el tool Artifact pasando `url`).
  - Base de datos (ArtifactData, mismo `url`):
    - `config/main`: precios, costes y supuestos (mismas claves que `contabilidad/datos.json` → `config`).
    - `dias/<YYYY-MM-DD>` por **fecha de pedido**:
      `{fecha, p1, p2, ent1, ent2, dev1, dev2, inc, envio_ent, envio_pend, carritos, cancelados, ads, nota, fuente, actualizado}`
      - `p1`/`p2`: pedidos de cada pack (null = aún no reportado).
      - `ent*`/`dev*`: de ESOS pedidos, entregados / devueltos por pack. Pendientes = el resto.
      - `inc`: cuántos de los pendientes están en incidencia.
      - `envio_ent`/`envio_pend`: suma del "Coste de envío (sin IVA)" real de entregados / pendientes.
        Sin ellos se usa la tarifa estándar de config (8,06 €).
    - `gastos/<id>`: `{fecha, concepto, categoria, importe}`.
    - `incidencias/<ID pedido>`: tablero del equipo. `{pedido, tracking, fecha_pedido, pack, pack_nombre, importe,
      nombre, telefono, ciudad, motivo, estado_almacen, abierta, primera_vez, resultado?, cerrada?}`.
      Nombre y teléfono SOLO mientras está abierta: al cerrarse (Entregado/Devuelto) se borran.
    - `acciones/<auto>`: gestiones del equipo `{pedido, tipo, nota, en, por (id usuario), autor (alias)}`;
      tipo ∈ whatsapp | no_contesta | acordada | almacen | rechaza.
    - `importaciones/ultima`: `{en, por, autor, pedidos, nuevas, cerradas}`.
  - Reglas de acceso: todo es solo del dueño (read/write `owner`) salvo `incidencias`, `acciones` e
    `importaciones` (read/write `interact`). El equipo se añade como Colaborador o Editor desde Compartir:
    ve solo la pestaña Incidencias, nunca los números. Un Lector no ve nada de clientes.
  - La página importa el Excel del almacén ella misma (botón "Subir Excel del almacén", SheetJS en el navegador):
    actualiza `dias` (solo si lo sube el dueño) e `incidencias` (abre nuevas, cierra las entregadas/devueltas).
  - La página lee el gasto de Meta en directo (conector "Meta ADS", tool `ads_get_ad_entities`).
- **Excel de contabilidad:** `contabilidad/CeraLux_Contabilidad.xlsx`, generado por `contabilidad/generar_excel.py`
  desde `contabilidad/datos.json`. Todo con fórmulas (Dashboard, Diario, Gastos, Config, Guía).

## Meta Ads

- Cuenta: HeatioShop, `ad_account_id` **26290722800581380**. Solo cuentan campañas cuyo nombre contiene **"CeraLux"**.
- Gasto diario: `ads_get_ad_entities` con `level: "campaign"`, `fields: ["amount_spent"]`,
  `filtering: [{"field":"campaign.name","operator":"CONTAIN","value":["CeraLux"]}]`,
  `time_range: {"since":..,"until":..}`, `time_increment: "1"`. Sumar `amount_spent.value` por `date_start`.

## Cuando Alec pasa el Excel del almacén

Lo normal es que lo suba él mismo en la pestaña Incidencias del dashboard. Si te lo pasa a ti:

1. Cópialo a una carpeta propia del scratchpad (tiene datos personales de clientes: NUNCA al repo; a la base de
   datos solo nombre/teléfono de incidencias abiertas).
2. `python -I contabilidad/importar_almacen.py <export.xlsx> --excluir 3503014 --json <scratchpad>/resumen.json`
   (excluye tests por ID y cualquier cliente cuyo nombre empiece por "Test"; revisa los AVISO de estados desconocidos).
3. Saca el gasto de Meta de esas fechas.
4. `ArtifactData list dias` para tener las versiones y escribe con `batch` (`if_version` en los que existen),
   conservando `nota` y poniendo `ads` de Meta y `fuente: "export almacén <fecha>"`.
   Incidencias: añade `--incidencias <scratchpad>/estados.json` al importador y aplica la misma lógica que la
   página (abrir las INCIDENCIA nuevas; cerrar las abiertas que estén Entregado/Devuelto quitando nombre y teléfono).
5. Regenera el Excel (abajo) y verifica: la "Liquidación almacén" debe cuadrar con Σ(PRECIO − COSTE TOTAL PEDIDO)
   de los entregados del export (diferencias de 1-2 céntimos por redondeo del almacén).

## Cuando Alec solo dice "hoy X pack 1, Y pack 2"

`get dias/<fecha>` y `set` con `p1`/`p2` nuevos y el gasto de Meta, conservando el resto. Responde con pedidos,
gasto, CPA vs CPA break-even (~11,8 € con los datos actuales), beneficio proyectado y semáforo
(ESCALAR ≤ 75% del break-even, VIGILAR ≤ 100%, CORTAR > 100%). Recuérdale las incidencias pendientes.

## Regenerar el Excel con los datos al día

1. `ArtifactData list` de `dias` y `gastos` y `get config/main`, todos con `out_dir=<scratchpad>/export`.
2. `python contabilidad/desde_dashboard.py <scratchpad>/export`
3. `python contabilidad/generar_excel.py` y recalcular con el `recalc.py` del skill xlsx (0 errores).
4. Commit + push.

## Modelo (igual en Excel y dashboard)

- Coste del pedido (almacén) = uds × 1,99 × 1,21 + envío real sin IVA (6,85-8,06 €) → 10,47 € / 12,88 € con 8,06 €.
- Liquidación almacén = PVP − coste del pedido de lo entregado (lo que paga el almacén).
- IVA: se repercute el 21% del PVP y solo se deduce el IVA del producto. El envío viene "sin IVA" y el almacén
  no le suma IVA → no hay IVA de envío que deducir (`logistica_con_iva = false`).
- Devuelto = −8,06 € (supuesto: se pierde el envío, el producto vuelve).
- Proyección: pendientes (incluidas incidencias) × tasa de entrega (80% estimada hasta 30 pedidos resueltos; luego la real).

Supuestos pendientes de confirmar con Alec: coste real de una devolución, si Meta le cobra IVA (ROI), si factura con IVA.
