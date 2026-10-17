# CeraLux · sistema de números

Marca de e-commerce de Alec: spray reparación de arañazos 120ml, venta contra reembolso (COD) en España.
Ads en Meta. Almacén COD con estados: Preparado → En ruta → Entregado / Devuelto, más INCIDENCIA
(no se pudo entregar a la primera; sigue pendiente, NO es venta perdida hasta que pasa a Devuelto),
Rechazado/Cancelado (no cuenta como pedido) y Carrito Abandonado (no es pedido).
Los estados llevan sufijos y se comparan por prefijo (página y `importar_almacen.py` igual): "Rehusado - en tránsito" = el
cliente lo rechazó en la puerta → Devuelto; "Rechazado" = anulado antes de salir → no es pedido; "Pedido nuevo",
"Confirmado - Pendiente de preparación", "Enviado", "En ruta"… → pendiente. Un carrito ("¿Es carrito?" = 1) en
"Pendiente de confirmación" aún no lo ha aceptado el cliente → sigue siendo carrito, no pedido.
Desde el export del 14/10 Dropi escribe "COSTO" en vez de "COSTE" en las cabeceras de costes: los dos importadores aceptan ambas.
Envío real por transportista (sin IVA): Correos Express 6,85 · CTT 7,20 · MRW 24/48H 7,39 · MRW 14H 8,06 (columna sin
título al final del export). Motivo "Pendiente de recoger en Franquicia destino" = el paquete espera en la oficina de la
transportista a que el cliente lo recoja.

## Cómo trabajar con Alec

- Está a tiempo completo con la tienda: los planes son de jornada completa y van a escalar (creativos, oferta, mercados).
- Las incidencias las gestiona Goverify: no ponérselas a Alec como tareas.
- El precio del Pack 2 (39,99 vs 39,90) no se vuelve a sacar; los guiones nuevos dicen "dos botes por menos de 40 €".

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
        Sin ellos se usa la tarifa estándar de config (7,55 €, el envío medio que cobra Dropi). Los pedidos que aún no han salido vienen con envío 0 en el
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
    - `caja/<YYYY-MM-DD>`: saldo de Dropi `{disponible, prestamo, en, fuente, tx?}` (el último manda). Se rellena con el botón
      "Subir transacciones de Dropi" (export de la cartera: suma los movimientos con ACTUALIZO WALLET = SI y estado Confirmado y
      guarda `tx = {ingresos, n_ent, envios, n_envios, envio_medio, producto, goverify, otros, desde, hasta}`;
      disponible = saldo propio + préstamo) o con "Guardar saldo a mano". Bloque "Tu dinero real hoy" (solo dueño):
      posición real = (disponible − préstamo) − Meta Ads desde el inicio − otros gastos; + lo que entrará de los pendientes
      (`entra` de `calcDia`: (PVP − producto) × entregas esperadas − envío + Go Verify de los que Dropi aún no ha confirmado),
      corregido con la cartera: − `yaAbonadas` (entregas que la cartera ya abonó y los `dias` aún no tienen) + `yaConfirmados`
      (pedidos solo-Shopify cuyo envío Dropi ya cobró) = "al cerrar los pedidos de hoy, antes de IVA"; − IVA de entregados y
      esperados = después de IVA. El 17/10: −507 € hoy, 701 € antes de IVA / 269 € después al cerrar (modelo: 667 / 235).
      Arriba del todo, bloque **"Tu saldo en Dropi"** (lo que Alec mira para retirar; sin IVA ni Meta, que se paga con tarjeta):
      saldo hoy (`disponible`), por llegar, saldo final esperado y retirable = saldo − 600 € (el préstamo, que se queda siempre en
      Dropi), con tabla por fecha estimada (`renderSaldo`): pedidos nuevos (− envío + Go Verify, `calcDia.cargo`), con retraso,
      un día por fila y las incidencias aparte sin fecha (`entraI`). Lo de cada día de pedido (`entraN`) se reparte con el plazo
      pedido → abono en **días laborables** (sin fines de semana ni `FESTIVOS` nacionales): `tx.plazos` si la cartera trae ≥10,
      si no `PLAZO_DEF` = {1:1, 2:13, 3:20, 4:3} (37 entregas sin incidencia, medido el 17/10: casi todo llega en 2-3 días
      laborables). Solo cuentan los plazos que aún pueden pasar (> días laborables transcurridos). El bloque anterior (posición con
      Meta e IVA) sigue debajo, plegado.
    - `stats/pedidos`: `{en, fechas: {<ID pedido Dropi>: <fecha de pedido>}}` (sin datos de clientes). Lo amplía el importador del
      almacén (solo dueño) y lo usa "Subir transacciones de Dropi" para medir `tx.plazos` con los "Reembolso por pedido
      entregado #ID" (sin las del tablero de incidencias). Sembrado con el export del 15/10.
    - `stats/transportistas`: `{en, fuente, por: {<transportista>: {n, ent, dev, inc, pend, envio}}}`. Lo escribe el importador
      del almacén (solo dueño; columna sin título o TRANSPORT). Tabla "Entrega por transportista" en Resumen. 15/10: Correos
      Express 72,7% entrega y 1 incidencia de 15; MRW 14H 37,5% (20/40); CTT 40% (9/24, rehúsan); MRW 24/48 50% (4/10).
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
    La consulta pide también `name` y `tags`: Dropi etiqueta cada pedido que importa con "Sync OK - Dropi PRO". Aviso rojo arriba
    si un pedido de Shopify de hace más de 1 h no tiene esa etiqueta ni está en ningún `shopify_ids` (no ha llegado a Dropi).
    `config/main` → `sync_manual`: IDs de Shopify metidos en Dropi a mano o por "Importar CSV Shopify" (no reciben la etiqueta);
    el aviso los ignora. Hoy: #1242–#1245 (15/10); el export del 15/10 ya los trae enlazados (sin doble conteo).
  - Pestañas **Embudo** (mismos datos de Meta que Creatividades, con el mismo selector de campaña: impresión → clic → carga →
    carrito → compra, dinero perdido por escalón y palancas a testear; clic→carga sano ≥45%, en móvil lo normal es 35-45%) y **Horas** (pedidos de Shopify por hora/franja/día de la semana + gasto
    horario de Meta si lo devuelve; no repartir presupuesto por franjas hasta ~300 pedidos). Solo dueño.
  - Pestaña **Creatividades** (solo dueño): lee de Meta en directo `ads_get_ad_entities` a nivel `ad` (gasto, impresiones,
    clics, visitas, compras, vídeo al 25%, ThruPlays, tiempo medio) y filtra por `meta_filtro_campana`.
    **Selector de campaña** (compartido con Embudo, se recuerda en el navegador): "Todas" o una campaña (por `campaign_id`).
    La corrección con pedidos reales y las medias de comparación usan SIEMPRE todas las campañas (los pedidos de Shopify no
    dicen de qué campaña vienen); lo que se enseña es la campaña elegida. Lista aparte "Encendidos, sin impresiones en este
    periodo": anuncios activos/en revisión/con problemas sin datos (recién lanzados, programados o sin entrega).
    Corrige las compras de Meta con pedidos reales del periodo (×pedidos/compras), calcula CPA real y beneficio estimado
    (pedidos × beneficio esperado por pedido − gasto) y da veredicto por anuncio con el CPA break-even (BE):
    ESCALAR (≥3 compras y CPA ≤ 75% BE) · MANTENER (≥3 y ≤ BE) · PROMETEDOR (<3 y ≤ BE) · VIGILAR · APAGAR
    (0 compras con gasto ≥ 1,5×BE, o CPA > BE con gasto ≥ 2×BE) · APRENDIENDO (gasto < 0,5×BE sin compras).
    Diagnóstico frente a la media de la campaña: gancho (vídeo al 25% ÷ impresiones), CTR, llegada a la web, compra/clic.
    **Prueba en curso** (arriba de Creatividades): `config/main` → `pruebas` = `[{id, nombre, campaign_id, inicio, decision,
    presupuesto, corte, nota, cerrada?}]`. Lee Meta desde `inicio` hasta hoy (sea cual sea el periodo) y compara con el resto
    de campañas en esas fechas: gasto vs previsto, CPA real, llegada a la web, tabla por anuncio con veredicto y qué hacer
    (apagar a `corte` € sin ventas; ganador = ≥2 ventas con CPA ≤ BE → a la principal). Al terminar, `cerrada: true`.
    Hoy: `rip1` = campaña `CeraLux ES — TESTING 2 — RIP COMPETENCIA` (120255201430200271), 16/10 → 19/10, 30 €/día,
    FINISH RIP 1-4 (vídeos de la competencia adaptados) + IMAGEN 1-3.
    `c10` = campaña `CeraLux ES — TESTING 3 — C10 PUERTA` (120255232148100271, conjunto 120255232148340271, mismos ajustes
    que TESTING 2, 20 €/día CBO), 18/10 → 20/10, corte 18 €: vídeo propio "La raya de la puerta" con hook y sin hook
    (`BROAD_C10_PUERTA_<HOOK>_V1` y `BROAD_C10_PUERTA_SINHOOK_V1`). Alec sube los dos anuncios en Ads Manager.
- **Excel de contabilidad:** `contabilidad/CeraLux_Contabilidad.xlsx`, generado por `contabilidad/generar_excel.py`
  desde `contabilidad/datos.json`. Todo con fórmulas (Dashboard, Diario, Gastos, Config, Guía).
- **Excel de creativos (aparte de la app):** `creativos/CeraLux_Creativos.xlsx`, generado por
  `creativos/generar_creativos.py` desde `creativos/biblioteca.json` (anuncios con ángulo, cuerpo, tipo/concepto/texto
  del hook; cuerpos partidos en apertura · solución · mecanismo · prueba · oferta · CTA; ajustes) y
  `creativos/historico.json` (Meta por anuncio y día + `pedidos_reales` por día de Shopify).
  Hojas: Panel (periodo DESDE/HASTA, KPIs, veredictos, Top 5, por ángulo/cuerpo/concepto), Mapa (ángulo → cuerpo →
  hooks), Creativos, Cuerpos, Matriz (concepto de hook × cuerpo; "—" = sin probar), Histórico, Pedidos, Config, Guía.
  Mismos veredictos que la pestaña Creatividades. Nomenclatura de anuncios (y del vídeo en el PC):
  `ÁNGULO_CUERPO_VISUAL_HOOK_VN`, p. ej. `BROAD_C4_BOLARDO_PRECIO_V1` (VISUAL = concepto del primer plano; HOOK = lo que
  dice la voz, `MUDO` si nada; textos de los hooks en `biblioteca.json` → `hooks`). Con ese nombre el script rellena
  cuerpo, formato, concepto, tipo y hook vacíos, y toma siempre el último nombre de Meta. Al regenerar, el script recoge antes lo editado a mano en el Excel
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
gasto, CPA vs CPA break-even (~11,8 € sin IVA con los datos del 17/10, el que marque el dashboard), beneficio proyectado y semáforo
(ESCALAR ≤ 75% del break-even, VIGILAR ≤ 100%, CORTAR > 100%). Recuérdale las incidencias pendientes.

## Regenerar el Excel con los datos al día

1. `ArtifactData list` de `dias`, `gastos` e `incidencias` y `get config/main`, todos con `out_dir=<scratchpad>/export`.
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

- Packs: Pack 1 = 1 ud a 29,99 €, Pack 2 = 2 uds a 39,99 €, Pack 3 = 3 uds + aspirador de regalo (SKU `ASPIRADORPEQ`, 4,19 € sin IVA,
  sin coste logístico extra) a 59,99 € desde el 15/10 (`p3_desde`): `p3_coste_ud` = (3 × 1,99 + 4,19) / 3 = 3,3867.
  Los pedidos de Pack 3 anteriores (3 uds a 49,99 €, primera venta el 08/10, #1193) usan `p3_antes_pvp`/`p3_antes_coste_ud`
  (página: `unitDia`; Excel: columna G de Config y `IF($A<P3_DESDE,…)` en el Diario) para que la liquidación siga cuadrando con Dropi.
  En EasySell el aspirador es el producto de Shopify "Aspirador de coche (REGALO)" a 0 €: revisar en el siguiente export
  cómo lo pinta Dropi en UDS/PRODUCTOS (si cuenta 4 uds, el importador lo sigue clasificando como Pack 3).
- Coste del pedido (almacén) = uds × 1,99 × 1,21 + envío real sin IVA (6,75-8,06 €) → 10,47 € / 12,88 € / 20,35 € (Pack 3 + aspirador; 15,28 € el de antes) con 8,06 €.
  `p*_logistica` = 7,55 € (envío medio real de la cartera, 112 envíos al 17/10) para los días sin envío real.
- Excel de contabilidad, Diario: B-D pedidos P1-P3, E-G entregados, H-J devueltos, K incidencias, L-M envío real, N Ads, O notas,
  P-Y cálculos (W liquidación, X cobrado, Y proyectado), AA-AO auxiliares. Config: packs en C/D/E, notas en F.
- Liquidación almacén = PVP − coste del pedido de lo entregado (lo que paga el almacén).
- IVA: se repercute el 21% del PVP y solo se deduce el IVA del producto. El envío viene "sin IVA" y el almacén
  no le suma IVA → no hay IVA de envío que deducir (`logistica_con_iva = false`).
- Devuelto = −`coste_devolucion` = −7,55 € (el envío ya cobrado al confirmar; la vuelta no se cobra aparte y el producto vuelve).
- Cartera de Dropi (export de transacciones del 17/10, cuadra al céntimo con el saldo propio 339,21 €): al CONFIRMAR un pedido
  Dropi cobra ya el envío real ("Coste de logística en pedido confirmado", 6,75-8,06 €) y Go Verify 0,95 €; al ENTREGARLO
  abona el PVP entero ("Reembolso por pedido entregado") y cobra el producto con IVA (2,41 / 4,82 / 7,22 €). Un devuelto no
  tiene más cargos: pierdes el envío ya cobrado. No hay retraso en los abonos: el saldo es bajo porque el envío de los
  pendientes ya está pagado. Carritos recuperados: Go Verify cobra 0,95 € + 0,10 € de gestión. El préstamo de 600 € no sale
  en las transacciones.
- Comisión por pedido (`comision_pedido`) = 0,95 € de Go Verify por pedido confirmado, se entregue o se devuelva
  (no va en el "coste total pedido" del export, así que no entra en la liquidación; se resta aparte en cobrado y proyectado).
- Dropi adelanta un préstamo de 600 €: el "Disponible" de su panel lo incluye, no es dinero propio.
- Pedidos #1154–#1157 (3 y 4 de octubre, CeraLux, sincronizados con Dropi) quedan fuera porque `fecha_inicio` es el 05/10.
  Pendiente que Alec confirme si son reales (entonces `fecha_inicio` = 2026-10-03 y exportar desde el 03/10) o pruebas.
  Meta gastó 0 € el 3 y 4 de octubre: no vienen de los anuncios de esta cuenta.
- Proyección (dashboard, desde el 17/10): los pendientes normales se entregan a la tasa de los resueltos SIN incidencia
  (resueltos − incidencias cerradas del tablero; 80% estimada hasta 30 resueltos) y las incidencias abiertas (`dias.inc`,
  por pack según el tablero `incidencias`) a la tasa de incidencias = (entregadas + `tasa_incidencias` × 10) ÷ (cerradas + 10):
  la estimación (35%) pesa como 10 incidencias y se va corrigiendo con cada una que se cierra (40,6% el 17/10, 3 de 6).
  El break-even usa la tasa final esperada = (entregados + pendientes × su tasa) ÷ todo lo enviado (65,1% el 17/10 →
  CPA máx. ~11,8 € sin IVA). Motivo: la tasa de lo ya resuelto (85%) está inflada porque las incidencias tardan en resolverse.
  **IVA: fuera de TODO el panel** (decisión firme de Alec el 17/10: lo que llega al saldo de Dropi es lo que retira, y punto).
  `aplicar_iva = false` en config y en el Excel (Config C20 = NO): beneficios, CPA máximo, semáforo, veredictos y Excel sin IVA.
  No volver a sacar el IVA en el panel ni restarlo salvo que él lo pida.
  **Pedidos recientes** (desde el 17/10): un pendiente sin incidencia todavía puede caer en incidencia. q = % de pedidos que
  acaban en incidencia = incidencias del tablero ÷ (pedidos × parte ya visible), solo días con ≥70% visible; `INC_CDF` =
  [0,14, 0,30, 0,70, 0,92, 0,98] a 0-4 días laborables (medido con 50 incidencias; a 5, todas). Tasa del pendiente =
  tasa sin incidencia + (tasa incidencias − tasa sin incidencia) × q(1−F)/(1−qF) (`tasaNormal`). El 17/10: q = 54,6%, pedido
  nuevo 63,3%, tasa final 65,1%, CPA máx. 11,79 €, beneficio proyectado desde el inicio ~490 € (antes 667 €: se daban
  los pedidos recientes al 90,5%, como si ya hubieran pasado sin incidencia). Excel: Config filas 60-65 (C61 = q, lo calcula
  `desde_dashboard.py`; C62 fecha del último export; C63:G63 INC_CDF; C64:L64 festivos) y Diario AR-AT.
  **Escenarios** (Resumen, debajo de los KPIs, sigue el periodo elegido): beneficio del periodo si la entrega FINAL fuese
  55/60/65/75/80/85% (85% = competencia) + la previsión actual. Lo resuelto no cambia; los pendientes se entregan a la tasa
  que haga falta (`margPend`/`margF` de `calcDia`); CPA máx. de cada fila = break-even con esa tasa.
  Excel de contabilidad igual: Config filas 49-58 (C52/C53 = incidencias cerradas/entregadas, las rellena `desde_dashboard.py`
  si el export trae `incidencias/`; TASA_INC, TASA_BE) y en el Diario AP = incidencias pendientes, AQ = tasa de sus pendientes.
  Al exportar para regenerar, baja también `incidencias` (lista) al mismo `out_dir`: datos personales, solo al scratchpad.

Supuestos pendientes de confirmar con Alec: si Meta le cobra IVA (ROI), si factura con IVA, si la comisión de 0,95 € lleva IVA,
si Dropi le devuelve el coste del producto cuando un devuelto vuelve al almacén. Excel de creativos: `cpa_be`/`ben_pedido` de `biblioteca.json` = break-even del dashboard.
