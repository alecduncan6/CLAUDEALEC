# Traspaso de la asesoría · CeraLux (Alec)

**Fecha del traspaso:** jueves 22/10/2026, 15:30 (hora de Madrid).
**Periodo cubierto:** del 03/10 (primeros pedidos) al 22/10/2026.
**Para:** el nuevo asesor de e-commerce que sigue con Alec en Claude Code.

Este documento recoge todo lo que pasó en la asesoría anterior: el negocio, las campañas, las decisiones, las reglas que se acordaron, lo que queda pendiente y cómo se trabajaba cada día.

Úsalo junto con `CLAUDE.md`, en la raíz del repositorio. `CLAUDE.md` es la referencia técnica del sistema: panel, base de datos, importadores y modelo. Este documento es la historia y el contexto de negocio. Si una cifra no coincide, manda la más reciente del panel.

> Convenciones
> - Fechas en formato dd/mm, todas de octubre de 2026 salvo que se diga otra cosa.
> - "Compras" son las del píxel de Meta. "Pedidos" son los reales, de Shopify o del almacén.
> - CPA = gasto ÷ pedidos.
> - "Break-even" o "CPA máximo" = lo máximo que se puede pagar por pedido sin perder dinero, con la tasa de entrega esperada.
> - Todo va **sin IVA**: es una decisión firme de Alec (ver §3).

---

## Índice

0. [Resumen en 1 minuto](#0--resumen-en-1-minuto)
1. [El negocio](#1--el-negocio)
2. [Historia de las campañas](#2--historia-de-las-campañas-desde-el-inicio)
3. [Decisiones importantes y su porqué](#3--decisiones-importantes-y-su-porqué)
4. [Reglas, umbrales y aprendizajes](#4--reglas-umbrales-y-aprendizajes)
5. [Problemas abiertos y próximos pasos](#5--problemas-abiertos-y-próximos-pasos)
6. [Cómo trabajábamos el día a día](#6--cómo-trabajábamos-el-día-a-día)
7. [Anexos: series de datos](#7--anexos-series-de-datos)

---

## 0 · Resumen en 1 minuto

**Quién es Alec**
- 23 años. Filmmaker con equipo audiovisual profesional: graba y monta él mismo (DaVinci, ElevenLabs).
- Objetivo: **10.000 €/mes de beneficio** con e-commerce.
- Al empezar tenía un trabajo de 10–14 h y de 17–20 h. **Desde el 15/10 dice que está a tiempo completo con la tienda**: los planes son de jornada completa.

**Qué vende**
- **CeraLux**, un spray reparador de arañazos de coche (120 ml). Se vende **contra reembolso (COD) en España**.
- Anuncios en **Meta**. Tienda en **Shopify** (podaris.store) con formulario **EasySell**.
- Almacén y envíos con **Dropi**. **Go Verify** confirma los pedidos por teléfono, recupera carritos y gestiona las incidencias.

**Dónde está (22/10)**

| Métrica | Valor |
|---|---|
| Pedidos desde el 05/10 | 153 (18 días, ~8,5/día) |
| Gasto en Meta | ~1.324 € |
| CPA medio | 8,52 € |
| Resueltos | 64 entregados y 29 devueltos (68,8 %) |
| Incidencias | 31 abiertas; de las 33 cerradas solo se entregaron 8 |
| Entrega final esperada | **62,8 %** |
| CPA máximo | **11,78 €** |
| Beneficio esperado desde el inicio | **~528 €** |
| Saldo en Dropi | 1.380 € (780 € retirables) |

**Lo que más pesa**
1. **La tasa de entrega.** Las incidencias acaban casi siempre devueltas, sobre todo con MRW 14H. Alec está cambiando de transportista.
2. **El desgaste de los creativos.** Los dos ganadores (FINISH RIP 1 y C4 TIENES) bajan desde el 20/10, y C4 dice "hasta el 31 de octubre", así que caduca. Hacen falta creativos nuevos ya.
3. **El ticket medio**, que sube con el Pack 3 + aspirador. Es la palanca que más está ayudando.

**Pendiente de respuesta de Alec (22/10)**
- Su "dale" para pausar FINISH RIP 2 y C10 HOOK.
- Que él baje TESTING 2 de 36 a 30 €/día.

---

## 1 · El negocio

### 1.1 Producto, precios y packs

**El producto**
- Spray reparador de arañazos de 120 ml, marca **CeraLux**.
- Solo funciona en arañazos superficiales, los que **no enganchan la uña**. La "prueba de la uña" es el eje de la comunicación y de la honestidad de la marca.

**Los packs** (se venden en EasySell; en Shopify el producto tiene una sola variante)

| Pack | Contenido | PVP | Desde |
|---|---|---|---|
| Pack 1 | 1 bote | **29,99 €** | inicio |
| Pack 2 | 2 botes | **39,99 €** | inicio. Va marcado por defecto como "MÁS VENDIDO" |
| Pack 3 (antiguo) | 3 botes | 49,99 € | 08/10 (primera venta, #1193) hasta el 14/10 |
| **Pack 3 (actual)** | **3 botes + aspirador de coche de REGALO** | **59,99 €** | **15/10** (primero con aspirador, #1264, 17/10) |

**El envío es gratis para el cliente**: todos los pedidos de Shopify lo tienen a 0 €.

**El aspirador de regalo**
- SKU `ASPIRADORPEQ`. Cuesta **4,19 € sin IVA** y no suma coste de envío. Es **negro**: en anuncios y formulario hay que usar fotos del real.
- En Shopify es el producto "Aspirador de coche (REGALO)" a 0 €: estado **Unlisted** (no Draft, porque EasySell no podría añadirlo), sin seguimiento de stock y con vendor "Dropi PRO".
- En el Excel del almacén el Pack 3 sale con 4 unidades. El importador lo clasifica igual como Pack 3 (salen avisos "4 uds no encaja", que son normales).
- Antes de cerrarlo se barajaron otros precios del aspirador: 2,69 € + IVA (reunión con Dropi), 8,75 € y 10,99 €.

**Tema cerrado: el precio del Pack 2**
- El anuncio C4 dice "dos botes por 39,90 €" y la tienda cobra 39,99 €.
- **Alec pidió el 15/10 que no se vuelva a mencionar.** Los guiones nuevos dicen "**dos botes por menos de 40 €**" o "2º bote a solo 10 €".

**Formulario EasySell (revisado el 15/10)**
- Pack 2 por defecto con "20 € cada unidad". Precios tachados 59,98 / 89,97 € (reales: lo que cuestan las unidades sueltas). Botón "PAGO CONTRA ENTREGA".
- Cambios que se pidieron:
  - quitar "valorado en 70 €" (prohibido y engañoso);
  - usar la foto real del aspirador (negro, no blanco);
  - cambiar "OFERTA DE OCTUBRE" por "PACK COMPLETO" u "OFERTA COCHE LIMPIO";
  - poner "MEJOR OFERTA" en vez de "LO ESTÁ ARRASANDO";
  - corregir la tilde de "Llévate".
- No está verificado que se hicieran todos.

### 1.2 Costes y lo que deja cada pedido

| Concepto | Importe |
|---|---|
| Producto | 1,99 €/bote sin IVA. Dropi lo cobra con IVA al entregar: 2,41 € (P1), 4,82 € (P2), 7,22 € (P3 antiguo), ~12,29 € (P3 con aspirador) |
| Envío real sin IVA, por transportista | Correos Express 24H 6,85 € · GLS 19h 6,75 € · CTT 24/48h 7,20 € · MRW 24/48H 7,39 € · MRW 14H 8,06 € |
| Envío medio | 7,55 € según la cartera de Dropi (112 envíos al 17/10). El panel usa el real del export (7,30 € el 22/10) y 7,55 € como tarifa estándar |
| Go Verify | **0,95 € por pedido confirmado**, se entregue o no. Carritos recuperados: 0,95 € + 0,10 € |
| Devuelto | **−7,55 €**: se pierde el envío ya cobrado. No hay cargo extra por la vuelta y el producto vuelve |

**Neto por pedido entregado** (PVP − producto − envío medio), según el panel del 22/10:

| Pack | Neto por entrega | Diferencia con Pack 1 |
|---|---|---|
| Pack 1 | 20,29 € | — |
| Pack 2 | 27,88 € | +7,59 € |
| Pack 3 + aspirador | 40,40 € | +20,11 € |

**Cómo mueve el dinero Dropi** (cuadrado al céntimo con la cartera el 17/10)
- **Al confirmar un pedido** cobra ya el envío real y Go Verify cobra 0,95 €.
- **Al entregarlo** abona el PVP entero y cobra el producto con IVA.
- **Un devuelto** no tiene más cargos.
- No hay retraso en los abonos. El saldo es bajo porque **el envío de los pendientes ya está pagado**.

**El préstamo de Dropi: 600 €**
- El "Disponible" de Dropi lo incluye, pero no es dinero de Alec y no sale en las transacciones.
- **Se queda siempre en Dropi**: retirable = saldo − 600 €.

**Plazo de abono:** de 37 entregas medidas el 17/10, 33 llegaron en 2–3 días laborables desde el pedido. Dropi no abona en fin de semana ni festivos.

### 1.3 Almacén, confirmación y transportistas

**Dropi** es el almacén y la logística contra reembolso.
- Contacto: **Andrés**.
- Lo que dijo en la reunión del 13/10:
  - tasa de entrega media del spray en Dropi: 75 % (573 pedidos del 23 al 30/09);
  - Alec es el 2º o 3º vendedor del spray; el primero vende ~2.500 uds/semana;
  - hay 45.000 uds pedidas hasta diciembre, así que hay que avisarle antes de escalar;
  - los vendedores rentables ganan 200–400 €/día;
  - propuso el Pack 3 con aspirador y abrir Portugal;
  - "no me importa si eres rentable mientras saques pedidos": él gana con el volumen.
- Andrés recomendó no cambiar de transportista. El asesor anterior insistió en cambiarlo con datos (ver §2.7).

**Go Verify (Goverify)**
- Confirma cada pedido por teléfono.
- Recupera carritos abandonados: 13 de 31, un 41,9 % al 22/10.
- **Gestiona las incidencias. No son tarea de Alec**: no se las pongas como tareas.

**Estados del almacén**
- Preparado → En ruta → Entregado / Devuelto.
- **INCIDENCIA** = no se pudo entregar a la primera. Sigue pendiente; no es venta perdida hasta que pasa a Devuelto.
- "Rehusado - en tránsito" = el cliente lo rechazó en la puerta → cuenta como devuelto.
- "Rechazado" = anulado antes de salir → no es pedido.
- "Carrito Abandonado" → no es pedido. Un carrito en "Pendiente de confirmación" sigue siendo carrito.

**Transportistas en uso:** Correos Express, CTT, MRW 24/48H, MRW 14H y, desde el 18/10, **GLS** (Dropi empezó a usarlo sin avisar).

### 1.4 Shopify y web

- **Tienda:** podaris.store, plan Basic. Nombre interno "CL Store" o tema "PODARIS". Antes tuvo otros 3 productos: Podaris Step, Mini Aire y Limpiador.
- **Producto:** `https://www.podaris.store/products/spray-reparacion-de-aranazos-120ml`.
- **Página de producto:** es un bloque HTML `#cl-pdp` dentro de la descripción. Convierte muy bien: un **8,6 % de las visitas compra** (12–20/10). No se tocan los textos.
- **Apps:** EasySell (formulario contra reembolso), Essential Sticky Cart (duplica botón), GemPages (carga en la cabecera aunque no se use), "Facebook e Instagram".
- **Problemas de velocidad detectados el 07/10:**
  - fuentes de Google extra (Archivo, Manrope);
  - GemPages;
  - fotos de 1,7–2,3 MB (una es un PNG de 2.400 px);
  - portadas de vídeo rotas.
  - El asesor ofreció arreglarlo en una copia del tema; Alec no respondió.
  - Desde este entorno no se puede abrir podaris.store (lo bloquea el proxy): para medir la velocidad, Alec tiene que pasar una captura de pagespeed.web.dev.
- **Sincronización con Dropi:**
  - Dropi etiqueta cada pedido que importa con "✅ Sync OK - Dropi PRO".
  - El 14/10 a las 15:09 se cortó la sincronización: #1242–#1245 entraron a mano por "Importar CSV Shopify" y están en `config/main → sync_manual`.
  - El panel avisa en rojo si un pedido de más de 1 h no llega a Dropi.
- **Pedidos que no cuentan:**
  - #1158 es el pedido de test (ID de almacén 3503014), siempre excluido.
  - El 3565715 del 21/10 es otro test: el cliente se llama "Test" y el importador lo excluye solo.
  - #1154–#1157 (3 y 4/10) están fuera porque el panel empieza el 05/10. **Sin confirmar si eran reales.**
- **Ojo con el checkout:** un carrito de 36,98 € (29,99 € + 6,99 € de envío) entró por el checkout normal de Shopify y no por EasySell. Hay que comprobar que ningún botón lleve ahí.

### 1.5 Meta

| Elemento | Dato |
|---|---|
| Cuenta publicitaria | **HeatioShop**, `26290722800581380` |
| Píxel | `1486649376319596` |
| Página de Facebook | `1016093914923135` |
| Atribución | 7 días tras clic y 1 día tras visualización |

- **Solo cuentan las campañas cuyo nombre contiene "CeraLux"**; el panel filtra por eso. En la cuenta hay campañas viejas de Podaris (Humidificador) en pausa.
- **El píxel es fiable:** 113 compras frente a 115 pedidos (17/10); 36 frente a 35 (16–18/10); 104 % de los pedidos de Shopify (21/10). Las diferencias son carritos recuperados por teléfono, que no pasan por Shopify.
- **Ajustes estándar de los conjuntos:**
  - España peninsular (sin Baleares, Canarias, Ceuta ni Melilla), 18–65 años;
  - **Advantage+ de público desactivado** y **sin Audience Network**;
  - optimización a Compra con el píxel, puja de mayor volumen.
- **Ajustes estándar de los anuncios:**
  - mejoras creativas de Advantage+ todas desactivadas y botón **"Comprar"**;
  - UTMs `utm_source=facebook&utm_medium=paid&utm_campaign={{campaign.name}}&utm_content={{ad.name}}`.
- El 60 % del gasto va a mayores de 45. El mejor tramo es **45–54 años**: CPA de 5,37–6,0 € frente a 10 € en 65+.

---

## 2 · Historia de las campañas desde el inicio

### 2.1 Línea de tiempo

| Fecha | Qué pasó |
|---|---|
| 03–04/10 | Pedidos #1154–#1158 (#1158 de test). Meta gastó 0 € esos días |
| **05/10** | **Arranca** "CeraLux ES — Test 3 ángulos — Oct 2026" (CBO, 45 €/día) con 3 conjuntos: AS1 RABIA, AS2 DINERO y AS3 BROAD |
| 07/10 | Se monta el sistema de números: Excel de contabilidad, panel en vivo, tablero de incidencias y Excel de creativos |
| 08/10 | Récord: **16 pedidos con 50 €** (CPA 3,15 €). Primer Pack 3 (#1193). Se añade la comisión de 0,95 € |
| 09/10 | Se crea "CeraLux ES — TESTING — BROAD C4 BOLARDO" (20 €/día, arranca el 10/10) con 3 hooks sobre el cuerpo C4: PRECIO, MIRA y TIENES. Nomenclatura nueva |
| 10/10 | La principal sube de 45 a **54 €/día** (lo hizo el asesor y Meta la pausó; Alec la volvió a encender). **El escalado no funcionó** (ver §2.3) |
| 11–12/10 | Alec de vacaciones. Domingo 11/10: 89,51 € para 6 pedidos |
| 12/10 | Veredicto del test C4: **TIENES ganador** (5,66 €/venta), PRECIO flojo, MIRA no entrega |
| 13/10 | Reunión con Dropi (Andrés). Caja baja (~200 €) |
| 14/10 | Auditoría: Dropi cambia "COSTE" por "COSTO" en el export. Se corta la sincronización Shopify→Dropi |
| **15/10** | Pack 3 + aspirador a 59,99 €. Con OK de Alec se **pausan** RABIA, DINERO, CARRITO_H7 y MIRA. Alec baja la principal a 45 €. Por la noche se pausa LLAVES_H2. Se crea "TESTING 2 — RIP COMPETENCIA" (30 €/día CBO) |
| 16/10 | Arranca TESTING 2 (FINISH RIP 1–4 + IMAGEN 1–3). Muy buen día: 100 € y 12–13 pedidos |
| **17/10** | Cartera de Dropi cuadrada. Modelo corregido (incidencias y pedidos recientes). **IVA fuera de todo el panel.** Se crea "TESTING 3 — C10 PUERTA" (20 €/día) |
| 18/10 | Se publica C10 (con hook y sin hook). Dropi empieza con GLS. Se evalúa la bandolera. Respuesta al proveedor Yohn. Por la noche, **auditoría y principal pausada** |
| 19/10 | Presupuestos nuevos (los pone Alec): TESTING 2 de 30 a **36 €** y TESTING C4 de 20 a **24 €**. TESTING 3 en **20 €**. Total **80 €/día** |
| 20/10 | Día malo (5 pedidos y 71,96 €). Noche de Champions. Alec dice "ya corrijo transportista" |
| **21/10** | Se descubre un **fallo de Meta** (§2.8) y se arregla el panel. Fecha de decisión de las pruebas RIP y C10 |
| 22/10 | Análisis completo: los ganadores se desgastan. Propuestas pendientes de OK (§5.1) |

### 2.2 Campaña principal "Test 3 ángulos" (05/10 → pausada el 18/10 por la noche)

- **ID** 120254979743130271. CBO, puja de mayor volumen.
- **Presupuesto:** 45 €/día (05/10) → 54 € (10/10) → 45 € (15/10, lo bajó Alec) → **pausada** la noche del 18/10.
- **Total:** 730,50 € y 84 compras (8,70 €) hasta el 19/10.

| Conjunto | Cuerpo | Ángulo | Gasto total | Compras | €/compra | Estado |
|---|---|---|---|---|---|---|
| AS3 · BROAD — hooks visuales | C3 (demo real sin cortes) | Demostración | 627,42 € | 76 | 8,26 € | campaña pausada |
| AS1 · RABIA — no fue culpa mía | C1 (bote animado tipo Pixar) | Rabia por un rayón ajeno | 51,70 € | 4 | 12,93 € | pausado 15/10 |
| AS2 · DINERO — voy a vender el coche | C2 (bote animado tipo Pixar) | Vender el coche | 51,38 € | 4 | 12,85 € | pausado 15/10 |

**Anuncios** (totales del 05/10 al 22/10, píxel)

| Anuncio | Gasto | Compras | €/compra | Notas |
|---|---|---|---|---|
| **BROAD_H1 BOLARDO** | 479,15 € | 60 | 7,99 € | **El primer gran ganador**: 6,77 €/compra la 1.ª semana. Se quemó: 17,40 € del 16 al 18/10. ~100.000 impresiones |
| BROAD_H2 LLAVES | 68,36 € | 9 | 7,60 € | Muy bueno el día 1 (5 ventas). Pausado el 15/10 por 22,94 € y 1 venta en los últimos días |
| BROAD_H3 LLAVES | 27,28 € | 2 | 13,64 € | |
| BROAD_H4 PERRO | 13,03 € | 2 | 6,52 € | Prometedor, pero Meta no le dio presupuesto |
| BROAD_H7 CARRITO | 37,19 € | 3 | 12,40 € | El mejor gancho (25,8 %) y la mitad de compra/visita. Pausado el 15/10 |
| BROAD_H5 PERRO / H6 CARRITO | 0,80 € / 1,61 € | 0 | — | Sin entrega |
| PIXARVENTA_H1 | 42,85 € | 4 | 10,71 € | DINERO |
| PIXARVENTA_H2 / H3 | 4,11 € / 4,42 € | 0 | — | |
| PIXARSUPER_H1 | 31,24 € | 3 | 10,41 € | RABIA. Fue el mejor gancho al principio (26 %) |
| PIXARSUPER_H2 | 20,46 € | 1 | 20,46 € | |

### 2.3 El escalado fallido del 10/10

Se subió la principal a 54 € y se lanzó el test C4 a 20 € (unos 75 €/día en total).

| | 05–09/10 | 10–13/10 |
|---|---|---|
| Gasto | 55 €/día | 75 €/día |
| Pedidos | 9,2/día | 8,5/día |
| CPA | 6,03 € | 8,82 € |
| Beneficio estimado | ~50 €/día | ~22 €/día |

**Diagnóstico:** el test usaba el mismo visual (BOLARDO) y el mismo público, así que **canibalizaba a la principal**. Con un solo ganador, el techo era de ~8–9 pedidos/día.

**Aprendizaje:** escalar es abrir **visuales nuevos**, no meter más dinero al mismo.

### 2.4 TESTING · BROAD C4 BOLARDO (desde el 10/10)

- **IDs:** campaña 120255084246940271, conjunto 120255084249730271 (`BROAD · C4 · BOLARDO x3 hooks · ES 18-65`). Copia exacta de AS3.
- **Presupuesto:** 20 €/día → **24 €/día** desde el 19/10.
- **Cuerpo C4:** demo + prueba de la uña + oferta de octubre. Cierra con "solo hasta el 31 de octubre, dos botes por 39,90…" → **caduca el 31/10**.

| Anuncio | Hook (voz) | Gasto total | Compras | €/compra | Estado |
|---|---|---|---|---|---|
| **BROAD_C4_BOLARDO_TIENES_V1** | "¿Tienes un arañazo como este? No lo lleves a pulir todavía." | 181,21 € | 24 | **7,55 €** | Activo. El mejor gancho de la cuenta (15–29 %) |
| BROAD_C4_BOLARDO_PRECIO_V1 | "Pulir el coche en España te puede costar 400 €. Esto, menos de 30." | 78,47 € | 8 | 9,81 € | Activo. Meta casi no lo usa desde el 21/10 |
| BROAD_C4_BOLARDO_MIRA_V1 | "Mira lo que pasa cuando echas este spray en los rayones de tu coche." | 13,02 € | 0 | — | Pausado el 15/10 |

**Evolución de TIENES**

| Periodo | Gasto | Compras | €/compra |
|---|---|---|---|
| 10–12/10 | 39,61 € | 7 | 5,66 € |
| 16–19/10 | 66,52 € | 10 | 6,65 € |
| 20–22/10 | 37,82 € | 3 | 12,61 € (bajando) |

### 2.5 TESTING 2 · RIP COMPETENCIA (desde el 16/10)

- **IDs:** campaña 120255201430200271, conjunto 120255201435740271 (`BROAD · RIP x4 · ES 18-65 · sin AN`). La creó el asesor el 15/10.
- **Presupuesto:** 30 €/día CBO → **36 €/día** desde el 19/10.
- **Contenido:** vídeos de la competencia (Nano Car Shine) **traducidos y adaptados** por Alec (FINISH RIP 1–4) y 3 imágenes. Mismos textos que TIENES y botón "Comprar".

| Anuncio | Gasto total | Compras | €/compra | Lectura |
|---|---|---|---|---|
| **FINISH RIP 1** | 160,75 € | 18 | 8,93 € | Ganador de la prueba: 7,44 € del 16 al 19/10. **Bajando**: 16,19 € del 20 al 22/10 |
| FINISH RIP 2 | 48,28 € | 3 | 16,09 € | **Por encima del máximo con más de 2× su gasto: toca apagar** |
| FINISH RIP 3 | 2,34 € | 0 | — | Sin entrega |
| FINISH RIP 4 | 3,62 € | 1 | — | Sin entrega |
| IMAGEN 1 / 2 / 3 | 0,09 € / 2,95 € / 8,82 € | 0 / 1 / 1 | — | Gastan muy poco |

- RIP 1 tiene el gancho más bajo de la cuenta (~7 %), pero una conversión muy alta los primeros días (24 % de compra por visita).
- El 18/10 Alec **rechazó pausar RIP 2, 3, 4 e IMAGEN 1 y 3** ("gastan poco, que gasten un poco más"). Ahora RIP 2 ya ha gastado lo suficiente para cumplir la regla de apagado.

### 2.6 TESTING 3 · C10 PUERTA (desde el 18/10)

- **IDs:** campaña 120255232148100271, conjunto 120255232148340271. 20 €/día CBO, mismos ajustes que TESTING 2. La creó el asesor el 17/10.
- **Contenido:** vídeo propio "**La raya de la puerta**": hombre de 45–55 años contando su historia en tono de bar. Alec lo aprobó ("está perfecto").
  - `BROAD_C10_SENOR_HOOK_V1`: 44,1 s, ID 120255232153490271.
  - `BROAD_C10_SENOR_SINHOOK_V1`: 42,7 s, ID 120255232171640271.
  - Publicados el 18/10 a las 00:56. Alec los había llamado "SEÑOR HOOK" y "SEÑOR SIN HOOK".
- **Defectos:** el de hook salió sin titular y los dos llevan "Ver detalles" en vez de "Comprar". Se ofreció rehacerlos y no se hizo.
- **Al publicar** salió el error "Se requiere un número de WhatsApp (#2446880)": era una opción de WhatsApp activada en el anuncio. Se arregla dejando el destino solo en "Sitio web".

| Anuncio | Gasto total | Compras | €/compra | Lectura |
|---|---|---|---|---|
| **C10 SIN HOOK** | 87,76 € | 8 | 10,97 € | **El único que mejora**: 13,82 € del 16 al 19/10 → 9,04 € del 20 al 22/10. Compran el 10,6 % de sus visitas |
| C10 HOOK | 6,18 € | 0 | — | Meta no lo entrega. Pausar |

### 2.7 Embudo, transportistas y tasa de entrega

**Embudo de todas las campañas CeraLux (12–20/10)**

```
2.844 clics → 894 cargas de página (31 %) → 168 formularios abiertos (19 %) → 77 compras (46 %)
```

- Clic → carga: en móvil lo normal es 35–45 %. Una parte de lo que se pierde son clics sin querer.
- **La página convierte bien: 8,6 % de las visitas compran.** Se pierde antes de que cargue y en el formulario.
- Por anuncio (clic→carga / visita→compra): C10 SIN HOOK 43 % / 6,7 % · RIP 1 37 % / 10,7 % · H1 31 % / 7,3 % · TIENES 25 % / 13,1 % · PRECIO 21 % / 14,6 % · MIRA 9 % / —.
- **Del 20 al 22/10 la compra por visita cayó del 11,7 % al 6,3 %** en los anuncios de RIP y C4. No es frecuencia (1,1–1,5). Coincide con la subida de presupuesto del 19/10 y con el gasto que Meta movió a RIP 2.

**Transportistas**

| Transportista | 15/10 (90 pedidos) | 22/10 (export del almacén) |
|---|---|---|
| Correos Express | 15 pedidos, 73 % de entrega, 1 incidencia | 19 pedidos: 14 entregados, 4 devueltos, **0 incidencias (78 %)** |
| CTT | 24 pedidos, 40 %, 9 incidencias (rechazos) | 28: 15 entregados, 10 devueltos, 2 incidencias (60 %) |
| MRW 24/48H | 10 pedidos, 50 % | 16: 8 entregados, 3 devueltos, 4 incidencias |
| **MRW 14H** | 40 pedidos, 38 %, 20 incidencias | 50: 19 entregados, **12 devueltos, 18 incidencias** (la peor) |
| GLS | — | 40: 8 entregados, 0 devueltos, 4 incidencias, **28 en camino** |

- De las incidencias de MRW 14H, la mayoría son "pendiente de recoger": no reintenta la entrega.
- Los pedidos de Correos van a 14 provincias distintas, así que no es por zona. La probabilidad de que la diferencia sea casualidad es menor de 1 entre 100.

**Evolución de la tasa de entrega**

| Fecha | Dato |
|---|---|
| 09/10 | 82 % de lo resuelto (9 entregados, 2 devueltos) |
| 14/10 | 83 % de lo resuelto (19 / 4). 21 en incidencia |
| 15/10 | 86 % de 37 resueltos. El asesor avisó de que estaba inflada |
| 17/10 | Dropi daba un 47,6 %. Modelo nuevo: final esperada **65,1 %** (sin incidencia 90,5 %, incidencias 40,6 %, 54,6 % de los pedidos acaban en incidencia) |
| 21/10 | 72,4 % de lo resuelto (63 / 24). Incidencias 30,3 % (8 de 28). Final 63,5 % |
| **22/10** | **68,8 % de lo resuelto (64 / 29). Sin incidencia 93,3 %. Incidencias 26,7 % (8 de 33). El 45,4 % acaba en incidencia. Final esperada 62,8 %** |

- Las incidencias se cierran casi siempre como "Rehusado - en tránsito": 5 de 5 el 22/10 y 14 de 18 el 21/10.
- Motivos al 20/10: 15 en oficina ("pendiente de recoger en franquicia"), 10 "no acepta el envío", 9 ausentes y 9 rehusados.

### 2.8 El fallo de Meta (descubierto el 21/10)

- Desde ~el 21/10, **cuando se pide el gasto día a día (`time_increment: "1"`), Meta devuelve gasto e impresiones a 0 en cada campaña o anuncio que tuvo compras ese día.** Pedido como total de un periodo, o por horas, sale bien.
- El panel llegó a mostrar **~67 € de gasto desde el inicio en vez de 1.240 €**: beneficio de 1.803 € en vez de 562 € y CPA de 0 €. Alec lo notó ("me estás dando números raros").
- **Arreglo en el panel:** detecta esos días, los vuelve a pedir como total del día y cuadra la suma con el total del rango.
- **A mano:** pide siempre totales por periodo, o día a día con `since` = `until` y sin `time_increment`. El procedimiento del Excel de creativos (CLAUDE.md) usa `time_increment` y está afectado.

### 2.9 Mezcla de packs y ticket medio

| Periodo | Pack 1 | Pack 2 | Pack 3 | Ticket medio |
|---|---|---|---|---|
| 05–07/10 | ~74 % | ~26 % | 0 % | ~32,6 € |
| 12–14/10 | 73 % | 23 % | 4 % | 33,07 € |
| 18–20/10 | 45 % | 36 % | 18 % | 39,08 € |
| 16–22/10 (7 días) | 37,5 % | 48,2 % | 14,3 % | **39,10 €** |
| 21 y 22/10 | 1 de 9 | 6 de 9 | 2 de 9 | 42–45 € |

**El aspirador de regalo y el Pack 2 por defecto funcionan.** Cada cliente que pasa de Pack 1 a Pack 3 deja +20 €.

### 2.10 Creativos escritos y producidos (biblioteca)

**Cuerpos** (`creativos/biblioteca.json`)

| ID | Nombre | Formato | Estado |
|---|---|---|---|
| C1 | Rabia · el bote habla | Animación tipo Pixar | Probado y apagado (AS1) |
| C2 | Dinero · vender el coche | Animación tipo Pixar, sin llamada a la acción | Probado y apagado (AS2) |
| C3 | Demostración real · sin cortes | Demostración grabada | Ganador inicial (H1 BOLARDO), quemado |
| C4 | Demo · uña + oferta octubre | Demo + voz | En uso (TIENES y PRECIO). **Caduca el 31/10** |
| C5 | Demo · uña + oferta sin fecha | Igual que C4 con cierre "dos botes por menos de 40 €" | **Escrito, sin grabar.** Es el recambio de C4 |
| C6 | Presentador en directo | Persona real (45–55) dice el texto de C5 | Escrito |
| C7 | Cuerpo de la competencia (RIP) | — | Al final se lanzó como FINISH RIP 1–4 |
| C8 | Aspirador (cierre con los 3 botes y el aspirador; unboxing de 28 s) | — | Escrito |
| C9 | Paisano en croma · cuñado + Pack 3 | Busto parlante en croma | **Alec lo rechazó por "cutre y poco natural"** |
| C10 | "La raya de la puerta" | Hombre contando su historia | **En el aire** (TESTING 3). Hay versión de 30 s y versión "TIMADA" ("yo pensaba que me iban a meter una timada…") |

Hay además una versión de mujer de "La raya de la puerta" ("Antes de que lo vea mi marido"), sin grabar.

**Hooks escritos** (`biblioteca.json → hooks`)
- **TIENES** (el mejor), PRECIO, MIRA (no funciona).
- **Sin grabar:** TALLER ("En el taller me pedían doscientos euros por esta raya. Y me la he quitado yo por menos de treinta."), UNA ("Antes de llevar el coche a pulir, haz una cosa: pasa la uña por la raya.") y PARKING ("Me la hicieron en el parking del súper. Ni una nota, ni nada.").

**Imágenes con Nano Banana**
- Tres tandas: la 1.ª Alec la vio "muy cutre"; luego HERO_REFLEJO, PARKING_HOMBRE50, FLATLAY_OFERTA, PARKING_RABIA, LUPA_UÑA y COMPARATIVA_TALLER.
- Estilo: ámbar #F2B705, Montserrat ExtraBold y Bebas Neue, formato 4:5.

**Banner fijo en los vídeos:** "2º BOTE A SOLO 10 € / ENVÍO GRATIS · PAGAS AL RECIBIR", a unos dos tercios de la altura. Variante: "LLÉVATE 3 + ASPIRADOR DE REGALO".

**Copy de los anuncios con aspirador (17/10)**
- Titular: "3 botes + aspirador de REGALO · 59,99 €".
- Descripción: "Pagas al recibirlo en casa".
- Texto principal: "¿Una raya en la puerta y en el taller te piden 180 €? Prueba esto antes. / Pasa la uña… si se engancha, es de chapista (te lo decimos claro). / 🎁 Pack de 3 botes + aspirador… 59,99 €. / ¿Solo quieres probarlo? Dos botes te salen por menos de 40 €. / 🚚 Pagas en efectivo cuando te llega a casa."

### 2.11 Otros proyectos y estudios

**Bandolera antirrobo** (SKU de Dropi `BOLSOANTIRROBO`, 6,29 € sin IVA), 18/10
- Puntuación 58/100: "ahora no, o test pequeño con condiciones".
- Competencia: Concept Global (Travel Sling Bag 2.0, 30 anuncios en España desde junio, 2x1), Nomalero y Opulenco "Trekka" (29,99 € contra reembolso).
- Plan propuesto:
  - oferta 1 × 29,99 € / 2 × 44,99 €;
  - 30 €/día durante 2 días con 3 vídeos propios; apagar a 15 € sin venta o con CPA > 8 € tras 3 pedidos;
  - campaña **sin "CeraLux" en el nombre**;
  - **antes del primer pedido, filtrar el panel por SKU**, o se contará como Pack 1 de CeraLux.
- **Sin "dale" de Alec. No se ha empezado.**

**Proveedor chino "Yohn"** (18/10)
- Respuesta final en inglés, tras dos correcciones de Alec: sin hablar del producto actual, solo pidiendo su catálogo (car care, accesorios de coche, gadgets de casa), MOQ, branding y envío DDP. Volumen indicado: 10–15 pedidos/día escalando a 30–50.
- Consejo: no decirle el coste actual (1,99 €), no enviar pedido a pedido desde China.
- Filtro de catálogo: mismo cliente, que se demuestre en vídeo en 5 s, venta a 25–40 € con coste menor de 3–4 € puesto en España.
- **Esperando su catálogo.**

**Competencia del spray**
- Nano Car Shine: 69 anuncios en 9 países de Europa del Este + Italia. Landing propia, sin Shopify.
- En Rumanía: Smania/Autogloss, vigoshop, Pingo Express y otros. Allí nadie usa la prueba de la uña ni el ángulo de los 400 €.
- Un competidor con el que Alec habló por teléfono hace ~100 pedidos/día y entrega el 85 %.
- Alec tiene un **compañero que hace exactamente lo mismo** y con quien compara resultados.

**Mercados**
- Portugal: lo propuso Andrés, con voz de ElevenLabs en portugués. Está en la cola; también Italia.
- Rumanía: "ahora no".
- Regla: entrar cuando España entregue más del 75 % durante una semana y haya 2–3 ganadores.

---

## 3 · Decisiones importantes y su porqué

| Fecha | Decisión | Por qué |
|---|---|---|
| 07/10 | Contabilidad **por fecha de pedido**, con beneficio "cobrado" (caja) y "proyectado" (para decidir) | Para comparar el gasto de un día con lo que acaban dejando los pedidos de ese día |
| 07/10 | Pestaña Incidencias para el equipo; nombre y teléfono solo mientras está abierta | Alec no quiere actualizarlas a mano. Protección de datos |
| 07/10 | Seguir con el panel como artifact de claude.ai, sin hosting propio | Suficiente hasta varios productos o 80–100 pedidos/día |
| 08/10 | El envío de los pedidos que aún no han salido se cuenta a la tarifa estándar | Dropi los pone a 0 € e inflaban los días recientes (~19 € el 07/10) |
| 08/10 | Carritos recuperados por teléfono = pedidos. Tres packs en todo el sistema | — |
| 09/10 | Los tests van **en campaña aparte**, nunca dentro de la principal ni del conjunto ganador | Añadir anuncios reinicia el aprendizaje; un conjunto nuevo en CBO no recibe dinero |
| 09/10 | Nomenclatura `ÁNGULO_CUERPO_VISUAL_HOOK_VN` | El Excel de creativos clasifica por nombre |
| 09/10 | **El asesor no toca nada en Meta sin OK explícito** | — |
| 10/10 | **Los presupuestos de campañas activas los cambia siempre Alec** | Cuando los cambia el asesor por la herramienta, Meta pausa la campaña (pasó el 10/10) |
| 15/10 | Pausar RABIA, DINERO, CARRITO_H7 y MIRA (con OK); por la noche, LLAVES_H2 | Desde el 10/10 los cuatro primeros sumaban 85 € para 3 pedidos (28,5 €). LLAVES_H2: 22,94 € y 1 venta |
| 15/10 | Pack 3 = 3 botes + aspirador a 59,99 € | Sube el ticket. Lo propuso Andrés. Un paquete más grande se acepta mejor en la puerta |
| 15/10 | **Alec corrige tres cosas:** no volver a mencionar el 39,90; las incidencias son de Goverify; está a tiempo completo | Quedó en `CLAUDE.md` |
| 15/10 | Sin Advantage+ de público, sin Audience Network, sin estrategia de ciclo de vida del cliente | Píxel con pocas compras; en un test solo cambia una cosa; Meta no ve los rechazos del contra reembolso |
| 15/10 | Desaconsejado: segunda tienda, stock privado, cuenta de agencia, LLC, USDT | Foco. Lo fiscal, con un gestor |
| **17/10** | **IVA fuera de TODO el panel** (beneficio, CPA máximo, semáforo, veredictos y Excel) | **Decisión firme de Alec:** "lo que me llega de saldo es lo que retiro y va a mi cuenta, y punto". El asesor primero dejó el CPA máximo con IVA y Alec lo corrigió. **No volver a sacarlo salvo que él lo pida** |
| 17/10 | Modelo de entrega corregido: los pedidos recientes aún pueden caer en incidencia | Antes se daban al 90,5 % y el beneficio salía inflado (667 → ~490 € ese día) |
| 17/10 | Bloque "Tu saldo en Dropi" y escenarios por tasa de entrega en el panel | Lo pidió Alec para saber qué puede retirar |
| 17/10 | C10 se lanza en campaña aparte (TESTING 3) | Para no estropear las pruebas RIP y C4 |
| 17/10 | Bajar la principal: el asesor propuso 30 € | Alec le corrigió ("analiza bien las métricas, creo que te estás equivocando"): había usado mal una ventana de 3 días. Corregido a 35 € |
| **18/10 noche** | **Pausar la principal** (lo hizo el asesor con OK). TESTING 2 a 36 € y TESTING a 24 € (los puso Alec) | H1 BOLARDO quemado (17,40 €). Los ganadores estaban en las pruebas. Simulación: sin la principal, más beneficio con menos gasto |
| 18/10 | Alec rechaza pausar RIP 2, 3, 4 e IMAGEN 1 y 3 | "Gastan poco, que gasten un poco más" |
| 20/10 | No tocar nada en Meta por un día malo; juzgar en bloques de 3 días | El 20/10 fue mala suerte más noche de Champions: tráfico normal, menos compras |
| 20/10 | Alec: "ya corrijo transportista; centrémonos en lo que podemos cambiar" | Creativos, web, formulario y packs |
| 21/10 | Arreglo del panel por el fallo de Meta | §2.8 |
| 22/10 | Propuesta: apagar RIP 2 y C10 HOOK; TESTING 2 de 36 a 30 € | **Pendiente de OK** (§5.1) |

---

## 4 · Reglas, umbrales y aprendizajes

### 4.1 CPA máximo (break-even) y semáforo

- **Fórmula:** beneficio esperado por pedido antes de anuncios = tasa final × neto medio por entrega − (1 − tasa) × 7,55 − 0,95. Con la mezcla de packs real.
- **Valor actual: 11,78 €** (22/10, entrega final esperada del 62,8 %). Lo calcula el panel: usa siempre el que marque.

**Semáforo**

| Estado | CPA | Acción |
|---|---|---|
| ESCALAR | ≤ 75 % del máximo (hoy ≤ 8,84 €) | +20 % cada 48 h |
| VIGILAR | ≤ 100 % (hoy ≤ 11,78 €) | No tocar |
| CORTAR | > 100 % | Recortar |

**Cómo ha ido cambiando el CPA máximo**

| Fecha | CPA máximo | Motivo |
|---|---|---|
| 07/10 | ~13,08 € | Envío con IVA, que estaba mal |
| 08/10 | 11,76 € → 10,81 € | Al añadir la comisión |
| 08–13/10 | 11,2 – 11,6 € | Según la mezcla de packs |
| 14/10 | 11,53 € | |
| 15/10 | 13,2 – 13,6 € | Inflado por el 86 % de lo resuelto |
| 17/10 | 8,35 € → 9,45 € → **11,79 €** | Modelo de incidencias; con IVA y luego sin IVA |
| 21/10 | 13,13 € → 12,51 € → 11,72 € | Unificar incidencias y export nuevo |
| **22/10** | **11,78 €** | |

### 4.2 Escalar y cortar

- **Subir presupuesto:** +20 % cada 48 h, solo si el CPA de los **últimos 3 días** está por debajo del 75 % del máximo. Más del 20–30 % de golpe devuelve el conjunto a aprendizaje.
- **Recortar:** si el CPA de 3 días supera el máximo, volver al presupuesto anterior. Nunca recortar todas las campañas a la vez: cada cambio reinicia el aprendizaje.
- **No subir presupuesto mientras la entrega sea mala**: cada pedido deja poco y el CPA siempre sube al escalar.
- Al escalar, fijarse en el **CPA marginal** (cuántos pedidos trae el dinero extra), no en la media.
- **Caja:**
  - Meta se paga hoy y Dropi abona días después, así que contra reembolso siempre va en negativo unos días.
  - No gastar en anuncios más de lo que hay en la cartera y la caja propias: **el préstamo no es presupuesto**.
  - A 35 pedidos/día hace falta un colchón de 4.000–5.000 €.
- Los conjuntos ganadores **no se tocan**: ni anuncios nuevos, ni edad, ni ubicaciones.
- Para escalar ganadores sin reiniciar: usar la publicación existente (post ID) en una campaña de escalado cuando se llegue a ~120 €/día solo con ganadores.

### 4.3 Veredictos por anuncio

Son los del panel y del Excel de creativos (BE = CPA máximo).

| Veredicto | Condición |
|---|---|
| **ESCALAR** | ≥ 3 compras y CPA ≤ 75 % BE |
| **MANTENER** | ≥ 3 compras y CPA ≤ BE |
| **PROMETEDOR** | < 3 compras y CPA ≤ BE |
| **VIGILAR** | Resto |
| **APAGAR** | 0 compras con gasto ≥ 1,5 × BE (~18 €), o CPA > BE con gasto ≥ 2 × BE (~24 €) |
| **APRENDIENDO** | Gasto < 0,5 × BE sin compras |

- **Pruebas:** un ganador necesita ≥ 2 ventas con CPA ≤ BE. Se apaga a `corte` € sin ventas (18–20 €). Se decide a las 48–72 h y nunca en fin de semana.
- No hace falta apagar anuncios que Meta ya no entrega dentro de un CBO, aunque se pueden pausar por limpieza.
- No juzgar con 3–4 ventas ni por un solo día. Bloques de **3 días**.

### 4.4 Cómo se analizan los creativos

**Métricas**
- **Gancho:** % que ve el 25 % del vídeo ÷ impresiones. Meta no da las reproducciones de 3 s.
- **CTR de enlace.**
- **Llegada a la web:** visitas ÷ clics.
- **Compra por visita:** la que manda.
- **CPA real:** con pedidos de Shopify, porque Meta no dice de qué campaña viene cada pedido.

**Lo que se ha visto**
- **Juzgar por compra por visita, no por gancho ni CTR.**
  - CARRITO enganchaba el doble que BOLARDO y vendía la mitad.
  - C10 tuvo el mejor CTR (4,2 %) y al principio el peor CPA.
  - PRECIO tiene la peor llegada a la web (21 %) y la mejor conversión (14,6 %).
- **Lo que vende es la demostración** (el arañazo y el paño en los primeros segundos, la prueba de la uña) y, últimamente, **la historia en primera persona de una persona real** (C10 SIN HOOK convierte el 10,6 % de las visitas).
- **El visual decide a quién llega el anuncio; el hook decide si se para.** Con creativos parecidos, Meta llega a la misma gente: escalar es abrir visuales nuevos.
- **Un creativo dura unos 10 días antes de quemarse.** Ritmo necesario: 3–5 creativos nuevos por semana. El test más barato son hooks nuevos sobre un cuerpo validado.
- **Los vídeos adaptados de la competencia funcionan unos días y se queman. Lo propio aguanta más.**
- **Un anuncio publicado no se puede editar**: hay que crear uno nuevo. Renombrar sí se puede y no reinicia el aprendizaje.

**Nomenclatura:** `ÁNGULO_CUERPO_VISUAL_HOOK_VN`, en mayúsculas, sin tildes ni espacios. `MUDO` = sin hook de voz. Mismo nombre en el PC, en DaVinci y en Meta, y la campaña debe llevar "CeraLux". Carpetas del PC: 1_BRUTOS, 2_VOCES, 3_DAVINCI, 4_EXPORTS/fecha_TEST_xx.

**Reglas de grabación y de guion**
- Solo arañazos reales y superficiales. Vertical 9:16, el arañazo en el primer segundo y con subtítulos.
- Coches y sitios españoles. Voz para 45–54 años.
- Lenguaje del cliente: "raya", "rayón", "lo del súper", "el taller te clava", "chapuza". Nada de "barniz", "fórmula", "pulimento" ni "aplicar". "Nanopolímeros" solo en texto en pantalla.
- **Nunca:**
  - "valorado en X €";
  - precios tachados que no sean reales (directiva Omnibus);
  - "este mes" u "octubre" en piezas nuevas;
  - antes/después generados con IA;
  - prometer "cualquier arañazo" sin la condición de la uña;
  - cifras inventadas (el "400 €" del taller va con "te puede costar", nunca "de media").
- **En contra reembolso no se dice "si no funciona, no lo pagas"**: invita a rechazar en la puerta.
- Sí avisar de que llamarán para confirmar, y que el aspirador del vídeo sea el real (negro).
- **Alec quiere guiones naturales, con humor e historia en primera persona.** Rechaza lo que suene "cutre" o de teletienda.
- **Los guiones van en el chat**, no en un .txt aparte.

### 4.5 El modelo de beneficio (panel y Excel)

- **Por cada pedido:**
  - **entregado:** PVP − producto − envío real;
  - **devuelto:** −7,55 €;
  - **en todos:** −0,95 € de Go Verify.
- **Pendientes:**
  - los normales se entregan a la tasa de los resueltos **sin** incidencia: 93,3 %;
  - las incidencias abiertas, a la tasa de incidencias: estimación del 35 % que pesa como 10 casos y se corrige con cada cierre (26,7 % el 22/10);
  - los pedidos recientes, a una mezcla, porque aún pueden caer en incidencia (q = 45,4 %). Un pedido nuevo vale ~63 %.
- **Incidencias de un día** = las abiertas del tablero o `dias.inc` si son más (una que vuelve a "en ruta" sigue siendo incidencia).
- **Por qué no vale la tasa de lo resuelto:** sale inflada (85 %) porque los entregados se resuelven rápido y las incidencias tardan.
- **Todo sin IVA.** Pendiente de confirmar con un gestor: IVA, ROI y régimen (autónomo o SL).

### 4.6 Lo que más pesa en el beneficio

- **La tasa de entrega pesa más que el CPA.** Cada punto de entrega vale ~0,3 € por pedido. Pasar del 65 % al 80 % equivale a bajar el CPA de 8 € a 3,4 €.
- **El ticket pesa tanto como el CPA:** subir el ticket 5 € es como bajar el CPA 5 €.
- **Una devolución borra ~25 €**: el beneficio que se pierde más el envío.
- **Para 10.000 €/mes** (~333 €/día):
  - con un margen de ~3–4 € por pedido harían falta ~90–100 pedidos/día;
  - con un 80 % de entrega y CPA de 8 €, ~40/día;
  - con un 85 % de entrega, ~32/día;
  - hoy se hacen ~8–10.

---

## 5 · Problemas abiertos y próximos pasos

### 5.1 Pendiente de respuesta de Alec (22/10, 15:30)

1. **"Dale" para pausar FINISH RIP 2 y C10 HOOK**: los pausa el asesor.
   - RIP 2: 48,28 € y 3 compras, 16,09 €/compra.
   - C10 HOOK: Meta no lo entrega.
2. **Alec baja TESTING 2 de 36 a 30 €/día**: vuelve al presupuesto con el que RIP 1 vendía a ~7 €. C4 (24 €) y C10 (20 €) no se tocan.
3. **Cerrar las pruebas `rip1` y `c10` en el panel** (`config/main → pruebas`, `cerrada: true`):
   - RIP: gana RIP 1, que se queda (bajando).
   - C10: gana SIN HOOK (10,97 € y mejorando), que se queda en su campaña.
   - La fecha de decisión era el 21/10.
4. **No subir presupuesto hasta que el CPA de 3 días baje de ~8,80 €.**

### 5.2 Urgente esta semana

- **Creativos nuevos.** Los dos ganadores (RIP 1 y TIENES) bajan a la vez. Orden recomendado:
  1. **C5:** el vídeo de TIENES (y de PRECIO) con el cierre sin fecha. Una hora de edición. **Tiene que estar en el aire antes del 31/10**, porque C4 dice "hasta el 31 de octubre": el 01/11 sería publicidad falsa.
  2. **3 hooks nuevos (TALLER, UNA, PARKING) sobre el mismo cuerpo**: solo hay que grabar los primeros 3 s.
  3. **Una pieza que venda el Pack 3 con aspirador en el formato que mejor funciona** (persona real con historia, como C10). El guion C9 en croma lo rechazó Alec: reformular en el estilo de "La raya de la puerta".
  - Todo se prueba en una campaña de test, nunca en los conjuntos ganadores.
- **Transportista:** Alec dijo el 20/10 "ya corrijo transportista". Hay que confirmar que Dropi deja de usar MRW 14H y pasa a Correos Express. **Vigilar GLS**: 28 de 40 pedidos aún en camino.

### 5.3 Abierto, menos urgente

- **Goverify:** preguntarles qué hacen con "No acepta el envío" y "Rehusado". Entre recuperar el 30 % o el 50 % de las 31 abiertas van unos 200 €. Es una petición a Goverify, no una tarea de incidencias para Alec.
- **Formulario EasySell:**
  - captura en móvil para revisarlo campo a campo (sin email, "No pagas nada ahora" junto al botón, Pack 2 por defecto);
  - confirmar los cambios del 15/10;
  - comprobar si se rompió la noche del 20/10 (5 formularios abiertos y 0 enviados);
  - mensaje de gracias explicando el pago en efectivo;
  - WhatsApp el día antes de la entrega.
- **Velocidad de la web:** captura de pagespeed.web.dev en móvil; desinstalar apps que no se usen (GemPages, Sticky Cart); fotos de menos de 300 KB.
- **C10 rehecho** con titular y botón "Comprar".
- **Transacciones de Dropi:** subir el export de la cartera de vez en cuando. Afina las fechas del bloque de saldo; el último completo es del 17/10.
- **Gastos fijos** (Shopify, EasySell, apps, dominio): nunca se apuntaron en "Otros gastos" del panel.
- **#1154–#1157** (3 y 4/10): ¿reales o pruebas? Si eran reales, `fecha_inicio` debería ser el 03/10.
- **Gestor fiscal:** IVA (modelo 303; el primer trimestre, octubre–diciembre, se paga en enero), alta en el ROI (si no, Meta cobra IVA), IRPF o sociedades, autónomo o SL, si la comisión de 0,95 € lleva IVA y si Dropi devuelve el coste del producto de un devuelto. **En el panel el IVA queda fuera por decisión de Alec.**
- **Dominio propio** (ceralux.es o .shop): recomendado el 07/10 porque el dominio podaris.store no casa con la marca. Sin hacer.
- **Bandolera:** pendiente del "dale". Antes de lanzarla, filtrar el panel por SKU.
- **Proveedor Yohn:** esperar su catálogo y filtrarlo.
- **Prueba de precio:** el Pack 1 a 34,99 € durante una semana. Propuesto el 17/10, sin decisión.
- **Calendario de ofertas, siempre con motivos verdaderos:**
  - 31/10: fin de la oferta de lanzamiento;
  - al llegar a 100 pedidos: hito real (ya superado);
  - **Black Friday el 27/11**: preparar oferta y creativos;
  - diciembre: regalo.
- **Mercados:** Portugal (con Andrés) cuando España entregue más del 75 % durante una semana con 2–3 ganadores.
- **Más adelante:** Advantage+ con 5–10 ganadores y más de 50 compras/semana; "Creative Testing" de Meta a partir de 150 €/día; repartir presupuesto por franjas horarias solo a partir de ~300 pedidos.

---

## 6 · Cómo trabajábamos el día a día

### 6.1 Cómo es Alec y qué espera

**Cómo escribe**
- Mensajes cortos, a veces dictados por voz, con erratas o en MAYÚSCULAS.
- Pide auditorías "a fondo", "instrucciones acertadas", "planes claros sin irse por las ramas".
- Cuando algo le lía, pide que se lo expliquen paso a paso y muy simple.

**El tono que quiere**
- **De cofundador**, directo, en español y de tú.
- Que le corrijas si se equivoca y que te corrijas si te equivocas. Él lo hace sin rodeos: "creo que te estás equivocando", "no me cuadra nada".
- Cuando él tiene razón, se reconoce y se rehace el análisis con los datos bien.

**Formato de respuesta** (lo pide en sus preferencias)
1. Respuesta rápida.
2. Por qué funciona (con números).
3. Pasos accionables (hoy / mañana / esta semana).
4. Errores comunes.
5. Insight avanzado.

Con tablas. Para estrategias, el marco **camino más rápido / más barato / más escalable**, y luego cuál elegirías y por qué.

**Cómo trabaja con las propuestas**
- Quiere que hagas todo lo que esté en tu mano y que le digas lo que le queda a él "sí o sí".
- Pide decisiones claras con datos y siempre el **beneficio real**: lo que llega al saldo de Dropi y puede retirar.

**No quiere**
- **El IVA** en ningún número: no volver a sacarlo salvo que lo pida.
- **El precio del Pack 2** (39,90 / 39,99): tema cerrado.
- **Las incidencias como tareas suyas**: son de Goverify.
- Planes de media jornada: está a tiempo completo.
- Diseños o guiones "cutres".
- Guiones en ficheros .txt: los quiere en el chat.
- Respuestas largas sin conclusión.

**Su estado de ánimo:** el 17/10 dijo sentirse "estancado". Se le recomendó pasar menos tiempo mirando el panel (máximo 10 minutos al día) y más grabando y hablando con Dropi.

### 6.2 Lo que pedía y lo que se hacía

| Lo que pasa o dice Alec | Lo que hacía el asesor |
|---|---|
| **Sube el Excel del almacén** en la pestaña Incidencias del panel (lo normal) **o lo pasa en el chat** (`resumen_excel_NN.xlsx`, con datos de clientes) | Lo copia al scratchpad (nunca al repo), lo pasa por el importador del panel (o `contabilidad/importar_almacen.py`), actualiza `dias`, `incidencias`, `stats/*` e `importaciones/ultima` en la base de datos con versiones (`if_version`), pone el gasto de Meta de cada día, regenera el Excel de contabilidad, comprueba que la liquidación cuadra con Σ(PRECIO − COSTE TOTAL PEDIDO) de los entregados (±0,20 €), hace commit y push, le pasa el .xlsx y le resume: saldo, retirable, tasa de entrega, CPA máximo, beneficio, incidencias y transportistas |
| **Da el saldo de Dropi** ("1380 en dropi") | Lo guarda en `caja/<fecha>` (`disponible`, `prestamo: 600`) y le da el retirable y lo que llegará por fecha |
| **Sube las transacciones de la cartera de Dropi** (`dropipro_transacciones_…xlsx`) | Botón "Subir transacciones de Dropi" del panel: cuadra el saldo y mide los plazos de abono |
| "Hoy X pack 1, Y pack 2, Z pack 3" | Actualiza `dias/<fecha>` con el gasto de Meta y responde con CPA frente al máximo, beneficio proyectado y semáforo |
| "Analiza las métricas / ¿qué apago?" | Lee Meta en directo (por periodos, no día a día: §2.8), cruza con Shopify y el almacén, compara periodos y da veredictos con las reglas de §4 |
| "Haz los cambios / dale" | Pausa o activa anuncios, crea campañas y conjuntos **en pausa** y renombra. **Nunca cambia presupuestos de campañas activas**: se los dice a Alec |
| Pide guiones, hooks o prompts | Los escribe en el chat siguiendo §4.4 y los añade a `creativos/biblioteca.json` si van a grabarse |
| "Actualiza el Excel de creativos" | Procedimiento de `CLAUDE.md`. Ojo con el fallo de Meta (§2.8): pide los datos por anuncio por periodos |
| Capturas (Dropi, EasySell, Ads Manager), mensajes de Andrés o de proveedores, enlaces de la competencia | Las revisa y da cambios concretos; redacta mensajes para Andrés o proveedores; busca en la biblioteca de anuncios de Meta |

### 6.3 Herramientas y piezas

- **Panel en vivo (la fuente de verdad):** https://claude.ai/artifact/A45KKiFp9hR3ZavkXX1be2
  - Fuente en `dashboard/ceralux_numeros.html`; se republica con `Artifact` pasando ese `url`.
  - Pestañas: Números (resumen, saldo de Dropi, escenarios, cuadre Shopify/Meta/almacén, transportistas, diario), Creatividades (con "Prueba en curso"), Embudo, Horas e Incidencias.
  - Lee Meta y Shopify en directo.
  - **Solo el dueño ve los números.** El equipo, como Colaborador, ve solo Incidencias.
- **Base de datos del panel:** colecciones `config/main`, `dias`, `incidencias`, `acciones`, `importaciones`, `caja`, `stats/pedidos`, `stats/transportistas` y `gastos`. Están todas descritas en `CLAUDE.md`.
- **Excel de contabilidad:** `contabilidad/CeraLux_Contabilidad.xlsx`, generado por `desde_dashboard.py` y `generar_excel.py`, con fórmulas. Se recalcula con el `recalc.py` del skill xlsx (0 errores).
- **Excel de creativos:** `creativos/CeraLux_Creativos.xlsx`, generado por `generar_creativos.py` desde `biblioteca.json` e `historico.json`. **El histórico solo llega al 07/10**: no se ha actualizado desde entonces.
- **Guiones:** `creativos/guiones/` (DEMO 3 hooks, C5 visuales nuevos, C9 croma).
- **Rama de trabajo:** `claude/epic-albattani-hzcxwy`. Cada commit termina con las líneas de atribución. No se abren PRs salvo que Alec lo pida.
- **Conectores:**
  - Meta Ads (`ads_get_ad_entities`, `ads_update_entity` para pausar o renombrar, `ads_create_campaign` / `ads_create_ad_set` en pausa);
  - Shopify (`graphql_query`, sin datos de clientes).
  - **No hay acceso a Dropi**: Alec no debe pasar su contraseña. Alternativa: Claude en Chrome desde la app de escritorio.
- **Privacidad:** los Excel del almacén y de la cartera tienen datos de clientes. Solo al scratchpad, nunca al repo. En la base de datos, nombre y teléfono solo de incidencias abiertas. El Python que los lee se ejecuta con `-I`.

### 6.4 Rutina que funcionaba

- **Por la mañana:** Alec sube el Excel del almacén del día y el saldo. Se actualiza todo y se da el resumen (saldo, retirable, entrega, CPA máximo, beneficio y lo que pasa con los transportistas).
- **A media tarde o por la noche:** análisis de anuncios por bloques de 3 días y decisiones (pausas con OK; presupuestos los cambia Alec, mejor de noche, entre las 23:00 y las 23:59).
- **Varias veces por semana:** guiones o hooks nuevos para grabar. Alec graba y monta; se lanzan en campaña de test.
- **Cuando algo "no cuadra":** auditoría completa cruzando Meta, Shopify, el almacén y el panel. Así se encontró el fallo de Meta del 21/10.

---

## 7 · Anexos: series de datos

### 7.1 Por día: gasto en Meta y pedidos (almacén, export del 22/10)

| Fecha | Gasto | Pedidos (P1 / P2 / P3) | € por pedido | Notas |
|---|---|---|---|---|
| 05/10 lun | 67,67 € | 9 (5/4/0) | 7,52 € | Arranque |
| 06/10 mar | 70,18 € | 8 (7/1/0) | 8,77 € | |
| 07/10 mié | 47,78 € | 7 (6/1/0) | 6,83 € | |
| 08/10 jue | 50,46 € | 16 (9/6/1) | **3,15 €** | Récord. Primer Pack 3 |
| 09/10 vie | 41,21 € | 5 (2/3/0) | 8,24 € | |
| 10/10 sáb | 64,81 € | 9 (6/2/1) | 7,20 € | Subida a 54 €/día + test C4 |
| 11/10 dom | 89,51 € | 6 (5/1/0) | 14,92 € | |
| 12/10 lun | 78,33 € | 14 (9/3/2) | 5,60 € | Festivo (Pilar): Dropi no entrega ni abona |
| 13/10 mar | 68,05 € | 9 (7/1/1) | 7,56 € | Reunión con Dropi |
| 14/10 mié | 73,35 € | 9 (5/4/0) | 8,15 € | Se corta la sincronización Shopify→Dropi |
| 15/10 jue | 60,80 € | 5 (2/3/0) | 12,16 € | Pack 3 + aspirador. Pausas |
| 16/10 vie | 100,40 € | 12 (5/7/0) | 8,37 € | Arranca TESTING 2 (RIP) |
| 17/10 sáb | 81,87 € | 9 (3/5/1) | 9,10 € | |
| 18/10 dom | 149,88 € | 12 (4/6/2) | 12,49 € | Arranca C10. Principal pausada esa noche |
| 19/10 lun | 85,31 € | 8 (5/1/2) | 10,66 € | Presupuestos 36/24/20 |
| 20/10 mar | 71,96 € | 6 (3/2/1) | 11,99 € | Día malo (Champions) |
| 21/10 mié | 79,53 € | 5 (1/3/1) | 15,91 € | Fallo de Meta descubierto |
| 22/10 jue (hasta las 15:30) | 38,71 € | 4 (0/3/1) | 9,68 € | |
| **Total** | **~1.320 €** | **153 (84/56/13)** | **~8,6 €** | |

En el almacén cuentan los carritos recuperados por teléfono, por eso a veces hay más pedidos que en Shopify. El gasto del 22/10 es parcial.

### 7.2 Por campaña (del 05/10 al 22/10, píxel)

| Campaña | Gasto | Compras | €/compra | Estado |
|---|---|---|---|---|
| Test 3 ángulos (principal) | 730,50 € | 84 | 8,70 € | Pausada el 18/10 por la noche |
| TESTING · BROAD C4 BOLARDO | 272,70 € | 32 | 8,52 € | Activa, 24 €/día |
| TESTING 2 · RIP COMPETENCIA | 226,77 € | 24 | 9,45 € | Activa, 36 €/día (propuesta: 30 €) |
| TESTING 3 · C10 PUERTA | 93,94 € | 8 | 11,74 € | Activa, 20 €/día |
| **Total** | **1.323,91 €** | **148** | **8,95 €** | **80 €/día** |

### 7.3 Últimos días por anuncio (16–19/10 frente a 20–22/10)

| Anuncio | 16–19/10 | 20–22/10 | Tendencia |
|---|---|---|---|
| FINISH RIP 1 | 111,64 € / 15 → 7,44 € | 48,57 € / 3 → 16,19 € | Bajando |
| FINISH RIP 2 | 15,01 € / 2 → 7,51 € | 32,09 € / 1 → 32,09 € | Apagar |
| C4 TIENES | 66,52 € / 10 → 6,65 € | 37,82 € / 3 → 12,61 € | Bajando |
| C4 PRECIO | 20,92 € / 3 → 6,97 € | 22,24 € / 1 → 22,24 € | Bajando; Meta ya casi no lo usa |
| C10 SIN HOOK | 41,46 € / 3 → 13,82 € | 45,21 € / 5 → 9,04 € | Mejorando |
| **Total de los activos** | **~275 € / 37 → 7,44 €** | **~190 € / 14 → 13,59 €** | **Por encima del máximo** |

Compra por visita: del 11,7 % al 6,3 %. Clics, CTR y frecuencia, normales.

### 7.4 Resultados y caja

**Beneficio esperado desde el inicio** (según el panel en cada momento)

| Fecha | Beneficio | Contexto |
|---|---|---|
| 07/10 | ~158 € con IVA si se entregaba todo (~90 € al 80 %) | 19 pedidos |
| 08/10 | ~173 € | 23 pedidos, con la comisión |
| 14/10 | 383 € al 80 % | ~70–200 € con escenarios realistas |
| 17/10 | 204 € → ~490 € | Al quitar el IVA y con el modelo nuevo |
| 21/10 | 562 € → **456 €** | Con el export del almacén: muchas incidencias devueltas |
| **22/10** | **528 €** | 7 días: +380 €. Últimos 3 días: más o menos en tablas |

**Saldo de Dropi**

| Fecha | Disponible | Propio (− 600 €) |
|---|---|---|
| 08/10 | 619,46 € | 19,46 € |
| 13/10 | ~200 € (aprox.) | — |
| 15/10 | 777,51 € | 177,51 € |
| 17/10 | 939,21 € | 339,21 € (cartera cuadrada al céntimo) |
| 21/10 | 1.198,34 € | 598,34 € |
| **22/10** | **1.380,00 €** | **780,00 €** |

**Previsión de saldo del 22/10**
- Por llegar: +1.082 € (818 € de pedidos normales y 264 € de incidencias).
- Saldo final esperado: ~2.462 € (retirable ~1.862 €).
- Por fecha: hoy +486 €, viernes 23/10 +146 €, lunes 26/10 +114 €, martes 27/10 +64 €, miércoles 28/10 +8 €, incidencias +264 € sin fecha.

**Liquidación del almacén** (PVP − coste de lo entregado), siempre cuadrada con el Excel de Dropi con una diferencia de ±0,20 €:

| Fecha | Liquidación |
|---|---|
| 09/10 | 202,43 € |
| 14/10 | 445,42 € |
| 15/10 | 757,75 € → 855,59 € |
| 21/10 | 1.522,24 € |
| 22/10 | 1.550,56 € |
