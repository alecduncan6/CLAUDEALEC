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
      `{fecha, p1, p2, p3, ent1, ent2, ent3, dev1, dev2, dev3, inc, envio_ent, envio_pend, carritos, recuperados, cancelados, ads, nota,
      fuente, actualizado}`
      - `p1`/`p2`/`p3`: pedidos de cada pack (null = aún no reportado). Pack = unidades: 1 → Pack 1, 2 → Pack 2, 3 o más → Pack 3.
      - `ent*`/`dev*`: de ESOS pedidos, entregados / devueltos por pack. Pendientes = el resto.
      - `inc`: cuántos de los pendientes están en incidencia.
      - `carritos`: carritos abandonados sin recuperar (estado "Carrito Abandonado", o carrito cancelado).
        `recuperados`: filas con "¿Es carrito?" = 1 que ya tienen estado de pedido (el equipo los ha recuperado por teléfono):
        cuentan en `p1`/`p2` como pedidos normales y no tienen pedido en Shopify (su "ID pedido Shopify" es del carrito).
      - `envio_ent`/`envio_pend`: suma del "Coste de envío (sin IVA)" real de entregados / pendientes.
        Sin ellos se usa la tarifa estándar de config (8,06 €). Los pedidos que aún no han salido vienen con envío 0 en el
        export: los importadores (página y `importar_almacen.py --envio-std`) les ponen la tarifa estándar.
      - `shopify_ids`: IDs de Shopify (legacyResourceId) de TODAS las filas del Excel del almacén ese día
        (incluidos tests y cancelados). `shopify_test`: los de pedidos de test. Sirven para no contar dos veces.
    - `gastos/<id>`: `{fecha, concepto, categoria, importe}`.
    - `incidencias/<ID pedido>`: tablero del equipo. `{pedido, tracking, fecha_pedido, pack, pack_nombre, importe,
      nombre, telefono, ciudad, motivo, estado_almacen, abierta, primera_vez, resultado?, cerrada?}`.
      Nombre y teléfono SOLO mientras está abierta: al cerrarse (Entregado/Devuelto) se borran.
    - `acciones/<auto>`: gestiones del equipo `{pedido, tipo, nota, en, por (id usuario), autor (alias)}`;
      tipo ∈ whatsapp | no_contesta | acordada | almacen | rechaza.
    - `importaciones/ultima`: `{en, por, autor, pedidos, nuevas, cerradas}`.
    - `caja/<YYYY-MM-DD>`: saldo del panel de Dropi `{disponible, prestamo, en}` (el último manda). Bloque "Tu dinero real hoy":
      (disponible − préstamo) − Meta Ads desde el inicio − otros gastos = posición real; − IVA de lo entregado = "tuyo de verdad".
      No usa supuestos del modelo salvo el IVA. Solo dueño.
  - Reglas de acceso: todo es solo del dueño (read/write `owner`) salvo `incidencias`, `acciones` e
    `importaciones` (read/write `interact`). El equipo se añade como Colaborador o Editor desde Compartir:
    ve solo la pestaña Incidencias, nunca los números. Un Lector no ve nada de clientes.
  - La página importa el Excel del almacén ella misma (botón "Subir Excel del almacén", SheetJS en el navegador):
    actualiza `dias` (solo si lo sube el dueño) e `incidencias` (abre nuevas, cierra las entregadas/devueltas).
  - La página lee el gasto de Meta en directo (conector "Meta ADS", tool `ads_get_ad_entities`), con compras del píxel por día.
  - Periodos: Hoy, Ayer, 7 días, Este mes, Mes pasado, Todo y "Día" (selector de fecha, periodo `dia:AAAA-MM-DD`,
    también en Creatividades; no se guarda en el navegador, al volver se ve el periodo habitual).
  - **Shopify en directo** (solo dueño): conector "Shopify", tool `graphql_query`, pedidos desde `fecha_inicio`
    (máx. 60 días) con `legacyResourceId, createdAt, cancelledAt, test, subtotalLineItemsQuantity` (sin datos de clientes).
    Fecha = día local de `createdAt`. Pack = 1 unidad → Pack 1; 2 → Pack 2; 3 o más → Pack 3. Los pedidos de Shopify cuyo ID no está
    en ningún `shopify_ids` se suman al día como pendientes ("aún no en almacén"); los de `shopify_test` se ignoran.
    Tabla "Cuadre": Shopify vs Meta (píxel) vs almacén por día y % que registra Meta. El #1158 es el pedido de test.
  - Pestaña **Creatividades** (solo dueño): lee de Meta en directo `ads_get_ad_entities` a nivel `ad` (gasto, impresiones,
    clics, visitas, compras, vídeo al 25%, ThruPlays, tiempo medio) y filtra por `meta_filtro_campana`.
    Corrige las compras de Meta con pedidos reales del periodo (×pedidos/compras), calcula CPA real y beneficio estimado
    (pedidos × beneficio esperado por pedido − gasto) y da veredicto por anuncio con el CPA break-even (BE):
    ESCALAR (≥3 compras y CPA ≤ 75% BE) · MANTENER (≥3 y ≤ BE) · PROMETEDOR (<3 y ≤ BE) · VIGILAR · APAGAR
    (0 compras con gasto ≥ 1,5×BE, o CPA > BE con gasto ≥ 2×BE) · APRENDIENDO (gasto < 0,5×BE sin compras).
    Diagnóstico frente a la media de la campaña: gancho (vídeo al 25% ÷ impresiones), CTR, llegada a la web, compra/clic.
- **Excel de contabilidad:** `contabilidad/CeraLux_Contabilidad.xlsx`, generado por `contabilidad/generar_excel.py`
  desde `contabilidad/datos.json`. Todo con fórmulas (Dashboard, Diario, Gastos, Config, Guía).
- **Excel de creativos (aparte de la app):** `creativos/CeraLux_Creativos.xlsx`, generado por
  `creativos/generar_creativos.py` desde `creativos/biblioteca.json` (anuncios con ángulo, cuerpo, tipo/concepto/texto
  del hook; cuerpos partidos en apertura · solución · mecanismo · prueba · oferta · CTA; ajustes) y
  `creativos/historico.json` (Meta por anuncio y día + `pedidos_reales` por día de Shopify).
  Hojas: Panel (periodo DESDE/HASTA, KPIs, veredictos, Top 5, por ángulo/cuerpo/concepto), Mapa (ángulo → cuerpo →
  hooks), Creativos, Cuerpos, Matriz (concepto de hook × cuerpo; "—" = sin probar), Histórico, Pedidos, Config, Guía.
  Mismos veredictos que la pestaña Creatividades. Al regenerar, el script recoge antes lo editado a mano en el Excel
  (hooks, cuerpos, notas, ajustes, pedidos) y lo guarda en `biblioteca.json`; los anuncios nuevos de Meta se añaden solos.

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

## Cuando Alec solo dice "hoy X pack 1, Y pack 2, Z pack 3"

`get dias/<fecha>` y `set` con `p1`/`p2`/`p3` nuevos y el gasto de Meta, conservando el resto. Responde con pedidos,
gasto, CPA vs CPA break-even (~11,2 € con los datos actuales), beneficio proyectado y semáforo
(ESCALAR ≤ 75% del break-even, VIGILAR ≤ 100%, CORTAR > 100%). Recuérdale las incidencias pendientes.

## Regenerar el Excel con los datos al día

1. `ArtifactData list` de `dias` y `gastos` y `get config/main`, todos con `out_dir=<scratchpad>/export`.
2. `python contabilidad/desde_dashboard.py <scratchpad>/export`
3. `python contabilidad/generar_excel.py` y recalcular con el `recalc.py` del skill xlsx (0 errores).
4. Commit + push.

## Cuando Alec dice "actualiza el Excel de creativos"

1. Si te pasa su copia del Excel con cambios, cópiala encima de `creativos/CeraLux_Creativos.xlsx` (el script recoge
   sus ediciones). Si trae guiones nuevos, añádelos a `biblioteca.json` (cuerpo nuevo = ID nuevo C4, C5…).
2. Meta por anuncio y día desde el último día de `historico.json` (rehaz siempre el último, puede estar a medias):
   `ads_get_ad_entities` con `level: "ad"`, `filtering: [{"field":"campaign_id","operator":"IN","value":[<IDs de las
   campañas con "CeraLux">]}]`, `time_increment: "1"`, `limit: 500`, `fields: ["id","name","adset_name","campaign_name",
   "effective_status","created_time","amount_spent","impressions","link_click","landing_page_view","omni_add_to_cart",
   "omni_initiated_checkout","omni_purchase","video_p25_watched_actions","video_p50_watched_actions",
   "video_thruplay_watched_actions","video_avg_time_watched_actions","frequency"]`. Sustituye en `filas` las de esas
   fechas (`fecha, ad_id, nombre, conjunto, campana, gasto, impr, clics, lpv, atc, checkout, compras, v25, v50, thru,
   tmedio, freq`).
3. `pedidos_reales` por día: pedidos de Shopify no cancelados ni test (misma consulta que el dashboard).
4. `python creativos/generar_creativos.py`, recalcular con `recalc.py` (0 errores), commit + push y pásale el xlsx.
   Comprueba que Σ gasto del Panel = gasto de Meta del periodo.

## Modelo (igual en Excel y dashboard)

- Packs: Pack 1 = 1 ud a 29,99 €, Pack 2 = 2 uds a 39,99 €, Pack 3 = 3 uds a 49,99 € (`p3_*` en config; primera venta el 08/10, #1193).
- Coste del pedido (almacén) = uds × 1,99 × 1,21 + envío real sin IVA (6,85-8,06 €) → 10,47 € / 12,88 € / 15,28 € con 8,06 €.
  Envío estándar del Pack 3 = 8,06 € (supuesto hasta ver el real en el export).
- Excel de contabilidad, Diario: B-D pedidos P1-P3, E-G entregados, H-J devueltos, K incidencias, L-M envío real, N Ads, O notas,
  P-Y cálculos (W liquidación, X cobrado, Y proyectado), AA-AO auxiliares. Config: packs en C/D/E, notas en F.
- Liquidación almacén = PVP − coste del pedido de lo entregado (lo que paga el almacén).
- IVA: se repercute el 21% del PVP y solo se deduce el IVA del producto. El envío viene "sin IVA" y el almacén
  no le suma IVA → no hay IVA de envío que deducir (`logistica_con_iva = false`).
- Devuelto = −8,06 € (confirmado con Dropi el 08/10: el envío inicial ya incluye la vuelta; se pierde el envío y el producto vuelve).
- Comisión del almacén (Dropi): `comision_pedido` = 0,95 € por pedido confirmado, se entregue o se devuelva
  (no va en el "coste total pedido" del export, así que no entra en la liquidación; se resta aparte en cobrado y proyectado).
- Dropi adelanta un préstamo de 600 €: el "Disponible" de su panel lo incluye, no es dinero propio.
- Pedidos #1154–#1157 (3 y 4 de octubre, CeraLux, sincronizados con Dropi) quedan fuera porque `fecha_inicio` es el 05/10.
  Pendiente que Alec confirme si son reales (entonces `fecha_inicio` = 2026-10-03 y exportar desde el 03/10) o pruebas.
- Proyección: pendientes (incluidas incidencias) × tasa de entrega (80% estimada hasta 30 pedidos resueltos; luego la real).

Supuestos pendientes de confirmar con Alec: si Meta le cobra IVA (ROI), si factura con IVA, si la comisión de 0,95 € lleva IVA,
si Dropi le devuelve el coste del producto cuando un devuelto vuelve al almacén. Excel de creativos: `cpa_be`/`ben_pedido` de `biblioteca.json` = break-even del dashboard.
