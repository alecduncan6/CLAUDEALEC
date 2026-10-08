#!/usr/bin/env python3
"""Genera la contabilidad de CeraLux en Excel a partir de datos.json.

    python contabilidad/generar_excel.py [datos.json] [salida.xlsx]

datos.json es una copia de la base de datos del dashboard (config, dias, gastos).
Todo el libro funciona con fórmulas: si editas el Excel a mano, se recalcula solo.
"""
import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

AQUI = Path(__file__).resolve().parent
DATOS = Path(sys.argv[1]) if len(sys.argv) > 1 else AQUI / "datos.json"
SALIDA = Path(sys.argv[2]) if len(sys.argv) > 2 else AQUI / "CeraLux_Contabilidad.xlsx"

DIAS_HOJA = 365          # filas preparadas en Diario
F0 = 6                   # primera fila de datos en Diario
FN = F0 + DIAS_HOJA - 1  # última fila de datos en Diario
G0, GN = 5, 304          # filas de datos en Gastos

# ---------- estilo ----------
FUENTE = "Arial"
TINTA = "1B2537"
ORO = "B7862B"
GRIS = "6B7686"
AZUL_INPUT = "0000FF"
VERDE_LINK = "008000"
FONDO_INPUT = "FFF4CC"
FONDO_CARD = "F3F5F8"
FONDO_SUB = "E6EAF0"
VERDE_OK, AMBAR, ROJO = "1E8449", "B9770E", "C0392B"

EUR = '#,##0.00 "€";[Red]-#,##0.00 "€";"-"'
EUR0 = '#,##0 "€";[Red]-#,##0 "€";"-"'
ENTERO = '#,##0;[Red]-#,##0;"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
ROAS = '0.00"x";[Red]-0.00"x";"-"'
FECHA = "dd/mm/yyyy"

fino = Side(style="thin", color="D5DBE3")
BORDE = Border(left=fino, right=fino, top=fino, bottom=fino)
BORDE_INF = Border(bottom=Side(style="thin", color="AEB7C4"))


def font(bold=False, color="000000", size=10, italic=False):
    return Font(name=FUENTE, bold=bold, color=color, size=size, italic=italic)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def celda(ws, ref, valor, *, bold=False, color="000000", size=10, fmt=None,
          fondo=None, align=None, wrap=False, italic=False, borde=None):
    c = ws[ref]
    c.value = valor
    c.font = font(bold, color, size, italic)
    if fmt:
        c.number_format = fmt
    if fondo:
        c.fill = fill(fondo)
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if borde:
        c.border = borde
    return c


def entrada(ws, ref, valor, fmt=None, nota=None):
    """Celda editable: texto azul sobre fondo amarillo."""
    c = celda(ws, ref, valor, color=AZUL_INPUT, fmt=fmt, fondo=FONDO_INPUT, borde=BORDE)
    if nota:
        c.comment = Comment(nota, "CeraLux")
    return c


def titulo(ws, ref, texto, sub=None, sub_ref=None):
    celda(ws, ref, texto, bold=True, color=TINTA, size=16)
    if sub:
        celda(ws, sub_ref, sub, color=GRIS, size=9, italic=True)


def seccion(ws, rango, texto):
    ini = rango.split(":")[0]
    ws.merge_cells(rango)
    celda(ws, ini, texto, bold=True, color="FFFFFF", size=10, fondo=TINTA)
    ws[ini].alignment = Alignment(horizontal="left", vertical="center", indent=1)


def cabecera(ws, fila, col_ini, textos, fondo=FONDO_SUB, color=TINTA):
    for i, t in enumerate(textos):
        c = ws.cell(row=fila, column=col_ini + i, value=t)
        c.font = font(True, color, 9)
        c.fill = fill(fondo)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE


def nombre(wb, n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


# ---------- datos ----------
datos = json.loads(DATOS.read_text(encoding="utf-8"))
cfg = datos["config"]
dias = {d["fecha"]: d for d in datos.get("dias", [])}
gastos = sorted(datos.get("gastos", []), key=lambda g: g.get("fecha", ""))
inicio = dt.date.fromisoformat(cfg["fecha_inicio"])

wb = Workbook()
dash = wb.active
dash.title = "Dashboard"
diario = wb.create_sheet("Diario")
hgastos = wb.create_sheet("Gastos")
conf = wb.create_sheet("Config")
guia = wb.create_sheet("Guía")
for ws in (dash, diario, hgastos, conf, guia):
    ws.sheet_view.showGridLines = False

# =====================================================================
# CONFIG
# =====================================================================
ws = conf
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 46
ws.column_dimensions["C"].width = 15
ws.column_dimensions["D"].width = 15
ws.column_dimensions["E"].width = 72
titulo(ws, "B1", f"Configuración · {cfg['marca']}",
       "Celdas azules con fondo amarillo = editables. Todo lo demás se calcula solo.", "B2")

cabecera(ws, 4, 2, ["Unit economics por pack", cfg["p1_nombre"], cfg["p2_nombre"], "De dónde sale"])
ws["B4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
filas_pack = [
    (5, "PVP contra reembolso (IVA incl.)", "pvp", EUR, "Importe que cobra el repartidor. Dato de Alec."),
    (6, "Unidades en el pack", "uds", ENTERO, "Dato de Alec."),
    (7, "Coste producto por unidad (sin IVA)", "coste_ud", EUR, "Export del almacén: 1,99 € por spray (sin IVA)."),
    (8, "Coste de envío estándar (sin IVA)", "logistica", EUR,
     "Tarifa máxima del almacén (8,06 €). Si el día tiene envío real del export, se usa el real (6,85-8,06 €)."),
]
for f, etiqueta, clave, fmt, nota in filas_pack:
    celda(ws, f"B{f}", etiqueta)
    entrada(ws, f"C{f}", cfg[f"p1_{clave}"], fmt)
    entrada(ws, f"D{f}", cfg[f"p2_{clave}"], fmt)
    celda(ws, f"E{f}", nota, color=GRIS, size=9, wrap=True)

calc_pack = [
    (9, "Coste producto (IVA incl.)", "={c}6*{c}7*(1+IVA_PROD)", "Uds × coste × (1 + IVA producto)."),
    (10, "Coste total del pedido", "={c}9+{c}8", "Cuadra con 'Coste total pedido' del almacén: 10,47 € y 12,88 €."),
    (11, "Beneficio dropshipper por pedido ENTREGADO", "={c}5-{c}10", "Lo que te liquida el almacén: 19,52 € y 27,11 €."),
    (12, "IVA repercutido (dentro del PVP)", "={c}5*IVA_VENTAS/(1+IVA_VENTAS)", "PVP × 21/121."),
    (13, "IVA soportado deducible", "={c}6*{c}7*IVA_PROD+IF(LOG_IVA=1,{c}8*IVA_VENTAS/(1+IVA_VENTAS),0)",
     "IVA del producto (+ IVA del envío solo si en Parámetros dices que lo lleva)."),
    (14, "IVA neto a pagar por pedido entregado", "=IF(APLICA_IVA=1,{c}12-{c}13,0)",
     "0 si en Parámetros pones NO al IVA."),
    (15, "Beneficio neto por pedido entregado (antes de Ads)", "={c}11-{c}14", "Lo que te queda de verdad por cada entrega."),
]
for f, etiqueta, formula, nota in calc_pack:
    celda(ws, f"B{f}", etiqueta, bold=(f in (11, 15)))
    for c in "CD":
        celda(ws, f"{c}{f}", formula.format(c=c), fmt=EUR, bold=(f in (11, 15)), borde=BORDE)
    celda(ws, f"E{f}", nota, color=GRIS, size=9)

cabecera(ws, 17, 2, ["Parámetros", "Valor", "", "Qué significa"])
ws["B17"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
params = [
    (18, "IVA de las ventas", cfg["iva_ventas"], PCT, "IVA general en España."),
    (19, "IVA del producto (lo aplica el almacén)", cfg["iva_producto"], PCT, "Coste producto × 1,21 en el export."),
    (20, "¿Descontar el IVA del beneficio?", "SÍ" if cfg["aplicar_iva"] else "NO", None,
     "SÍ si facturas como autónomo o SL (Hacienda se queda ~4,8 € por Pack 1 entregado). NO solo para ver el bruto."),
    (21, "¿El envío lleva IVA deducible?", "SÍ" if cfg.get("logistica_con_iva") else "NO", None,
     "NO: el export dice 'Coste de envío (SIN IVA)' y el coste total no le suma IVA, así que no hay IVA que deducir."),
    (22, "Coste por pedido devuelto", cfg["coste_devolucion"], EUR,
     "Confirmado con Dropi: el envío de ida ya incluye la vuelta, así que pierdes el envío (8,06 €) y el producto vuelve."),
    (23, "Recargo sobre el gasto en Ads", cfg["recargo_ads"], PCT,
     "0% si Meta factura sin IVA (alta en ROI) o lo deduces. 21% si te cobran IVA y no lo recuperas."),
    (24, "Tasa de entrega estimada", cfg["tasa_entrega_estimada"], PCT,
     "Para proyectar pendientes e incidencias mientras no tengas datos propios."),
    (25, "Pedidos resueltos para usar tu tasa real", cfg["min_resueltos"], ENTERO,
     "Con estos entregados + devueltos, el libro pasa a usar tu tasa de entrega real."),
    (26, "% de Pack 2 estimado (sin datos)", cfg["pct_pack2_estimado"], PCT, "Solo se usa hasta que haya pedidos."),
    (27, "Objetivo de beneficio neto mensual", cfg["objetivo_mensual"], EUR0, "La meta: 10.000 €/mes."),
    (28, "Fecha de inicio", inicio, FECHA, "Primer día del Diario y del resumen mensual."),
    (29, "Cuenta publicitaria Meta", f"{cfg['meta_cuenta']} ({cfg['meta_ad_account_id']})", None,
     "De aquí sale el gasto en Ads (conector de Meta)."),
    (30, "Filtro de campañas", cfg["meta_filtro_campana"], None,
     "Solo cuenta campañas cuyo nombre contiene este texto."),
    (31, "Comisión del almacén por pedido confirmado", cfg.get("comision_pedido", 0.95), EUR,
     "Dropi cobra 0,95 € por cada pedido confirmado, se entregue o se devuelva."),
]
for f, etiqueta, valor, fmt, nota in params:
    celda(ws, f"B{f}", etiqueta)
    entrada(ws, f"C{f}", valor, fmt)
    celda(ws, f"E{f}", nota, color=GRIS, size=9, wrap=True)
dv = DataValidation(type="list", formula1='"SÍ,NO"', allow_blank=False)
ws.add_data_validation(dv)
dv.add("C20")
dv.add("C21")

cabecera(ws, 32, 2, ["Calculado con tus datos", "Valor", "", "Cómo se calcula"])
ws.row_dimensions[32].height = 20
ws["B32"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
rng = lambda col: f"Diario!${col}${F0}:${col}${FN}"  # noqa: E731
entregados = f"(SUM({rng('D')})+SUM({rng('E')}))"
devueltos = f"(SUM({rng('F')})+SUM({rng('G')}))"
calc = [
    (33, "IVA aplicado (1 = sí)", '=IF(C20="SÍ",1,0)', ENTERO, ""),
    (34, "Envío con IVA deducible (1 = sí)", '=IF(C21="SÍ",1,0)', ENTERO, ""),
    (35, "IVA recuperable por devolución", "=IF(AND(APLICA_IVA=1,LOG_IVA=1),C22*IVA_VENTAS/(1+IVA_VENTAS),0)", EUR, ""),
    (36, "Coste neto por devolución", "=C22-C35", EUR, ""),
    (37, "Pedidos resueltos (entregados + devueltos)", f"={entregados}+{devueltos}", ENTERO, "Histórico."),
    (38, "Tasa de entrega real", f'=IF(C37>0,{entregados}/C37,"")', PCT, "Entregados ÷ (entregados + devueltos)."),
    (39, "Tasa de entrega USADA en proyecciones", '=IF(AND(C37>=C25,C38<>""),C38,C24)', PCT,
     "Real si hay suficientes pedidos resueltos; si no, la estimada."),
    (40, "% de Pack 2 real", f'=IF(SUM({rng("M")})>0,SUM({rng("C")})/SUM({rng("M")}),"")', PCT, "Upsell real."),
    (41, "% de Pack 2 usado", '=IF(C40="",C26,C40)', PCT, ""),
    (42, "Envío medio real por pedido (sin IVA)",
     f'=IF(SUM({rng("AI")})>0,(SUM({rng("I")})+SUM({rng("J")}))/SUM({rng("AI")}),"")', EUR,
     "Media de los envíos reales del export del almacén."),
    (43, "Ticket medio esperado", "=(1-C41)*C5+C41*D5", EUR, ""),
    (44, "Beneficio neto medio por pedido entregado",
     '=(1-C41)*(C5-C9-IF(C42="",C8,C42)-C14)+C41*(D5-D9-IF(C42="",D8,C42)-D14)', EUR,
     "Mezcla de packs, con tu envío medio real."),
    (45, "Beneficio esperado por pedido confirmado (antes de Ads)", "=C39*C44-(1-C39)*C36-COMISION", EUR,
     "Tasa × beneficio por entrega − (1 − tasa) × coste de devolución − comisión del almacén."),
    (46, "CPA break-even (máximo que puedes pagar por pedido)", "=C45/(1+RECARGO_ADS)", EUR,
     "Por encima de este CPA pierdes dinero."),
    (47, "ROAS break-even", '=IF(C46>0,C43/C46,"")', ROAS, "Facturación bruta ÷ Ads mínima para no perder."),
]
for f, etiqueta, formula, fmt, nota in calc:
    celda(ws, f"B{f}", etiqueta, bold=(f in (46, 47)))
    celda(ws, f"C{f}", formula, fmt=fmt, bold=(f in (46, 47)), borde=BORDE)
    celda(ws, f"E{f}", nota, color=GRIS, size=9)

for n, ref in {
    "P1_PVP": "$C$5", "P2_PVP": "$D$5", "P1_CPROD": "$C$9", "P2_CPROD": "$D$9",
    "P1_LOG": "$C$8", "P2_LOG": "$D$8", "P1_IVAN": "$C$14", "P2_IVAN": "$D$14",
    "IVA_VENTAS": "$C$18", "IVA_PROD": "$C$19", "COSTE_DEV": "$C$22", "RECARGO_ADS": "$C$23",
    "OBJETIVO_MES": "$C$27", "FECHA_INICIO": "$C$28", "APLICA_IVA": "$C$33", "LOG_IVA": "$C$34",
    "IVA_DEV": "$C$35", "COMISION": "$C$31", "TASA_USADA": "$C$39", "BEN_PEDIDO": "$C$45", "CPA_BE": "$C$46", "ROAS_BE": "$C$47",
}.items():
    nombre(wb, n, f"Config!{ref}")

# =====================================================================
# DIARIO
# =====================================================================
ws = diario
titulo(ws, "A1", "Diario de pedidos",
       "Una fila por día de PEDIDO. Lo amarillo se rellena (lo hace Claude con el export del almacén). "
       "Pendientes = preparados + en ruta + incidencias; una incidencia no es venta perdida hasta que pasa a Devuelto.", "A2")
ws["A3"].value = ("Ejemplo: el 05/10 entran 5 Pack 1 y 4 Pack 2 → B=5, C=4. Luego se entregan 3 Pack 1 y 1 Pack 2 "
                  "→ D=3, E=1. 3 de los pendientes están en incidencia → H=3.")
ws["A3"].font = font(color=GRIS, size=9, italic=True)

D_ENVIO_ENT = "IF(ISNUMBER(I{r}),I{r},D{r}*P1_LOG+E{r}*P2_LOG)"
D_ENVIO_PEND = "IF(ISNUMBER(J{r}),J{r},X{r}*P1_LOG+Y{r}*P2_LOG)"
cols = [
    # (letra, cabecera, ancho, formato, fórmula o None si es entrada)
    ("A", "Fecha del pedido", 12, FECHA, None),
    ("B", "Pedidos Pack 1", 8, ENTERO, None),
    ("C", "Pedidos Pack 2", 8, ENTERO, None),
    ("D", "Entregados Pack 1", 9, ENTERO, None),
    ("E", "Entregados Pack 2", 9, ENTERO, None),
    ("F", "Devueltos Pack 1", 9, ENTERO, None),
    ("G", "Devueltos Pack 2", 9, ENTERO, None),
    ("H", "Incidencias (pendientes)", 10, ENTERO, None),
    ("I", "Envío real entregados (€)", 10, EUR, None),
    ("J", "Envío real pendientes (€)", 10, EUR, None),
    ("K", "Gasto Ads Meta (€)", 10, EUR, None),
    ("L", "Notas", 24, None, None),
    ("M", "Pedidos totales", 8, ENTERO, "=B{r}+C{r}"),
    ("N", "Pendientes", 9, ENTERO, "=X{r}+Y{r}"),
    ("O", "Facturación bruta", 11, EUR, "=B{r}*P1_PVP+C{r}*P2_PVP"),
    ("P", "Ticket medio", 9, EUR, '=IF(M{r}>0,O{r}/M{r},"")'),
    ("Q", "CPA", 9, EUR, '=IF(M{r}>0,K{r}/M{r},"")'),
    ("R", "ROAS", 8, ROAS, '=IF(K{r}>0,O{r}/K{r},"")'),
    ("S", "Tasa de entrega", 9, PCT, '=IF(D{r}+E{r}+F{r}+G{r}>0,(D{r}+E{r})/(D{r}+E{r}+F{r}+G{r}),"")'),
    ("T", "Liquidación almacén (entregados)", 12, EUR,
     "=D{r}*(P1_PVP-P1_CPROD)+E{r}*(P2_PVP-P2_CPROD)-" + D_ENVIO_ENT),
    ("U", "Beneficio cobrado (real)", 12, EUR,
     "=T{r}-(D{r}*P1_IVAN+E{r}*P2_IVAN)-(F{r}+G{r})*(COSTE_DEV-IVA_DEV)-AH{r}-AJ{r}"),
    ("V", "Beneficio proyectado", 12, EUR, "=AC{r}-AD{r}-AE{r}-AF{r}-AG{r}-AH{r}-AJ{r}"),
    ("W", "", 2, None, None),
    ("X", "Pendientes Pack 1", 9, ENTERO, "=MAX(0,B{r}-D{r}-F{r})"),
    ("Y", "Pendientes Pack 2", 9, ENTERO, "=MAX(0,C{r}-E{r}-G{r})"),
    ("Z", "Entregados proy. P1", 9, '0.0;-0.0;"-"', "=D{r}+X{r}*TASA_USADA"),
    ("AA", "Entregados proy. P2", 9, '0.0;-0.0;"-"', "=E{r}+Y{r}*TASA_USADA"),
    ("AB", "Devueltos proy.", 9, '0.0;-0.0;"-"', "=F{r}+G{r}+(X{r}+Y{r})*(1-TASA_USADA)"),
    ("AC", "Ventas proy. (IVA incl.)", 11, EUR, "=Z{r}*P1_PVP+AA{r}*P2_PVP"),
    ("AD", "Coste producto proy.", 10, EUR, "=Z{r}*P1_CPROD+AA{r}*P2_CPROD"),
    ("AE", "Envíos proy.", 10, EUR, "=" + D_ENVIO_ENT + "+" + D_ENVIO_PEND + "*TASA_USADA"),
    ("AF", "Devoluciones proy.", 10, EUR, "=AB{r}*COSTE_DEV"),
    ("AG", "IVA neto proy.", 10, EUR, "=Z{r}*P1_IVAN+AA{r}*P2_IVAN-AB{r}*IVA_DEV"),
    ("AH", "Ads total (con recargo)", 10, EUR, "=K{r}*(1+RECARGO_ADS)"),
    ("AI", "Pedidos con envío real", 9, ENTERO, "=IF(ISNUMBER(I{r}),D{r}+E{r}+X{r}+Y{r},0)"),
    ("AJ", "Comisión almacén", 9, EUR, "=M{r}*COMISION"),
]
FMT = {l: f for l, _, _, f, _ in cols}
ENTRADAS = {"B": "p1", "C": "p2", "D": "ent1", "E": "ent2", "F": "dev1", "G": "dev2", "H": "inc",
            "I": "envio_ent", "J": "envio_pend", "K": "ads", "L": "nota"}
for letra, cab, ancho, _fmt, formula in cols:
    ws.column_dimensions[letra].width = ancho
    if not cab:
        continue
    es_in = letra in ENTRADAS
    c = ws[f"{letra}5"]
    c.value = cab
    c.font = font(True, "5A4300" if es_in else TINTA, 9)
    c.fill = fill(FONDO_INPUT if es_in else FONDO_SUB)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDE
ws.row_dimensions[5].height = 54
celda(ws, "X3", "Cálculos auxiliares (no tocar) →", color=GRIS, size=9, italic=True)

celda(ws, "A4", "TOTAL", bold=True, color=TINTA)
for letra, _, _, fmt, formula in cols:
    if letra in ("A", "L", "W") or fmt is None:
        continue
    celda(ws, f"{letra}4", f"=SUM({letra}{F0}:{letra}{FN})", bold=True, fmt=fmt, color=TINTA)
celda(ws, "P4", '=IF(M4>0,O4/M4,"")', bold=True, fmt=EUR, color=TINTA)
celda(ws, "Q4", '=IF(M4>0,K4/M4,"")', bold=True, fmt=EUR, color=TINTA)
celda(ws, "R4", '=IF(K4>0,O4/K4,"")', bold=True, fmt=ROAS, color=TINTA)
celda(ws, "S4", '=IF(D4+E4+F4+G4>0,(D4+E4)/(D4+E4+F4+G4),"")', bold=True, fmt=PCT, color=TINTA)

for i in range(DIAS_HOJA):
    r = F0 + i
    fecha = inicio + dt.timedelta(days=i)
    d = dias.get(fecha.isoformat(), {})
    celda(ws, f"A{r}", fecha, fmt=FECHA, borde=BORDE)
    for letra, clave in ENTRADAS.items():
        v = d.get(clave)
        if v in ("", None) or (letra in "DEFGH" and v == 0):
            v = None
        celda(ws, f"{letra}{r}", v, color=AZUL_INPUT, fmt=FMT[letra], borde=BORDE)
    for letra, _, _, fmt, formula in cols:
        if formula:
            celda(ws, f"{letra}{r}", formula.format(r=r), fmt=fmt, borde=BORDE,
                  color=(GRIS if letra >= "X" or len(letra) == 2 else "000000"), bold=(letra in "TUV"))

ws.conditional_formatting.add(f"B{F0}:C{FN}", FormulaRule(
    formula=[f'AND($K{F0}>0,$B{F0}="",$C{F0}="",$A{F0}<TODAY())'], fill=fill("FAD7A0")))
ws.conditional_formatting.add(f"H{F0}:H{FN}", CellIsRule(operator="greaterThan", formula=["0"], fill=fill("FAD7A0")))
ws.conditional_formatting.add(f"A{F0}:A{FN}", FormulaRule(
    formula=[f"$A{F0}=TODAY()"], fill=fill("F5E6C4"), font=Font(name=FUENTE, bold=True)))
ws.conditional_formatting.add(f"U{F0}:V{FN}", CellIsRule(
    operator="greaterThan", formula=["0"], font=Font(name=FUENTE, bold=True, color=VERDE_OK)))
ws.freeze_panes = f"B{F0}"
ws["K5"].comment = Comment("Lo trae Claude del conector de Meta Ads (cuenta HeatioShop, campañas 'CeraLux').", "CeraLux")
ws["H5"].comment = Comment("Pedidos en INCIDENCIA (p. ej. destinatario ausente). Siguen pendientes: "
                           "no son venta perdida hasta que el almacén los marca como Devuelto.", "CeraLux")
ws["I5"].comment = Comment("Suma del 'Coste de envío (sin IVA)' real de los pedidos entregados. "
                           "Vacío = se usa la tarifa estándar de Config.", "CeraLux")
ws["J5"].comment = Comment("Suma del envío real de los pedidos pendientes (preparado, en ruta, incidencia).", "CeraLux")
ws["T5"].comment = Comment("Lo que te liquida el almacén por lo entregado: PVP − coste total del pedido. "
                           "Debe cuadrar con su panel.", "CeraLux")
ws["U5"].comment = Comment("Dinero ya ganado: liquidación − IVA − devoluciones − Ads − comisión del almacén. "
                           "Los días recientes salen negativos porque aún hay pedidos pendientes.", "CeraLux")
ws["V5"].comment = Comment("Lo que dejará el día cuando se resuelvan los pendientes, "
                           "usando la tasa de entrega de Config.", "CeraLux")

# =====================================================================
# GASTOS
# =====================================================================
ws = hgastos
titulo(ws, "A1", "Otros gastos",
       "Todo lo que no es producto, envío ni Ads: Shopify, apps, dominio, muestras, gestoría...", "A2")
ws["A3"].value = "Ejemplo de fila: 01/10/2026 | Shopify Basic | Suscripciones | 36,00 € | mensual"
ws["A3"].font = font(color=GRIS, size=9, italic=True)
cabecera(ws, 4, 1, ["Fecha", "Concepto", "Categoría", "Importe (€)", "Notas"], fondo=FONDO_INPUT, color="5A4300")
for letra, ancho in zip("ABCDE", (12, 34, 18, 13, 40)):
    ws.column_dimensions[letra].width = ancho
cats = "Suscripciones,Apps,Contenido / UGC,Muestras,Herramientas,Gestoría,Comisiones,Otros"
dvc = DataValidation(type="list", formula1=f'"{cats}"', allow_blank=True)
ws.add_data_validation(dvc)
dvc.add(f"C{G0}:C{GN}")
for i in range(GN - G0 + 1):
    r = G0 + i
    g = gastos[i] if i < len(gastos) else {}
    fecha = dt.date.fromisoformat(g["fecha"]) if g.get("fecha") else None
    celda(ws, f"A{r}", fecha, color=AZUL_INPUT, fmt=FECHA, borde=BORDE)
    celda(ws, f"B{r}", g.get("concepto"), color=AZUL_INPUT, borde=BORDE)
    celda(ws, f"C{r}", g.get("categoria"), color=AZUL_INPUT, borde=BORDE)
    celda(ws, f"D{r}", g.get("importe"), color=AZUL_INPUT, fmt=EUR, borde=BORDE)
    celda(ws, f"E{r}", g.get("nota"), color=AZUL_INPUT, borde=BORDE)
ws.freeze_panes = f"A{G0}"



# =====================================================================
# DASHBOARD
# =====================================================================
ws = dash
for i, ancho in enumerate([2] + [11.5] * 12 + [2]):
    ws.column_dimensions[chr(ord("A") + i)].width = ancho

DESDE, HASTA = "$C$3", "$F$3"


def S(col):
    """Suma de una columna de Diario dentro del periodo elegido."""
    return (f"SUMIFS(Diario!${col}${F0}:${col}${FN},Diario!$A${F0}:$A${FN},\">=\"&{DESDE},"
            f"Diario!$A${F0}:$A${FN},\"<=\"&{HASTA})")


GASTOS_PERIODO = (f"SUMIFS(Gastos!$D${G0}:$D${GN},Gastos!$A${G0}:$A${GN},\">=\"&{DESDE},"
                  f"Gastos!$A${G0}:$A${GN},\"<=\"&{HASTA})")

titulo(ws, "B1", f"{cfg['marca']} · Dashboard de rentabilidad",
       f"{cfg['producto']} · contra reembolso · Ads: Meta ({cfg['meta_cuenta']})", "B2")
celda(ws, "B3", "Desde", bold=True, color=TINTA, align="right")
entrada(ws, "C3", "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)", FECHA,
        "Escribe una fecha para fijar el periodo. Por defecto: día 1 del mes actual.")
ws.merge_cells("C3:D3")
celda(ws, "E3", "Hasta", bold=True, color=TINTA, align="right")
entrada(ws, "F3", "=TODAY()", FECHA, "Por defecto: hoy.")
ws.merge_cells("F3:G3")
celda(ws, "H3", "Días", bold=True, color=TINTA, align="right")
celda(ws, "I3", f"=MAX(1,{HASTA}-{DESDE}+1)", fmt=ENTERO, bold=True)
celda(ws, "J3", "Tasa entrega usada", bold=True, color=TINTA, align="right")
ws.merge_cells("J3:K3")
celda(ws, "L3", "=TASA_USADA", fmt=PCT, bold=True, color=VERDE_LINK)

# --- tarjetas KPI (6 por fila, 2 columnas cada una) ---
# Referencias: B7 neto proy., D7 liquidación, F7 cobrado, H7 pedidos, J7 facturación, L7 ads,
# B10 CPA, D10 CPA BE, B13 entregados, D13 pendientes, F13 incidencias, H13 devueltos, L13 semáforo.
seccion(ws, "B5:M5", "RESUMEN DEL PERIODO")
TARJETAS = [
    [("Beneficio neto proyectado", "=F25", EUR),
     ("Liquidación almacén", f"={S('T')}", EUR),
     ("Beneficio cobrado (real)", f"={S('U')}+F24", EUR),
     ("Pedidos", f"={S('M')}", ENTERO),
     ("Facturación bruta", f"={S('O')}", EUR),
     ("Gasto Ads", f"={S('K')}", EUR)],
    [("CPA (coste por pedido)", '=IF(H7>0,L7/H7,"")', EUR),
     ("CPA break-even", "=CPA_BE", EUR),
     ("ROAS", '=IF(L7>0,J7/L7,"")', ROAS),
     ("ROAS break-even", "=ROAS_BE", ROAS),
     ("Ticket medio", '=IF(H7>0,J7/H7,"")', EUR),
     ("% Pack 2", f'=IF(H7>0,{S("C")}/H7,"")', PCT)],
    [("Entregados", f"={S('D')}+{S('E')}", ENTERO),
     ("Pendientes", f"={S('N')}", ENTERO),
     ("En incidencia", f"={S('H')}", ENTERO),
     ("Devueltos", f"={S('F')}+{S('G')}", ENTERO),
     ("Tasa de entrega", '=IF(B13+H13>0,B13/(B13+H13),"")', PCT),
     ("Semáforo campaña",
      '=IF(B10="","SIN DATOS",IF(B10<=D10*0.75,"ESCALAR",IF(B10<=D10,"VIGILAR","CORTAR")))', None)],
]
for fila_t, grupo in zip((6, 9, 12), TARJETAS):
    for j, (etq, formula, fmt) in enumerate(grupo):
        c1 = chr(ord("B") + 2 * j)
        c2 = chr(ord(c1) + 1)
        ws.merge_cells(f"{c1}{fila_t}:{c2}{fila_t}")
        ws.merge_cells(f"{c1}{fila_t + 1}:{c2}{fila_t + 1}")
        celda(ws, f"{c1}{fila_t}", etq.upper(), bold=True, color=GRIS, size=8, fondo=FONDO_CARD, align="center")
        celda(ws, f"{c1}{fila_t + 1}", formula, bold=True, color=TINTA, size=15, fmt=fmt,
              fondo=FONDO_CARD, align="center")
        for cc in (c1, c2):
            ws[f"{cc}{fila_t}"].border = Border(top=fino, left=fino if cc == c1 else None,
                                                right=fino if cc == c2 else None)
            ws[f"{cc}{fila_t + 1}"].border = Border(bottom=fino, left=fino if cc == c1 else None,
                                                    right=fino if cc == c2 else None)
    ws.row_dimensions[fila_t + 1].height = 30
for regla, color in (("ESCALAR", VERDE_OK), ("VIGILAR", AMBAR), ("CORTAR", ROJO)):
    ws.conditional_formatting.add("L13", CellIsRule(operator="equal", formula=[f'"{regla}"'],
                                                    fill=fill(color), font=Font(name=FUENTE, bold=True,
                                                                                color="FFFFFF", size=15)))
ws.conditional_formatting.add("B7", CellIsRule(operator="lessThan", formula=["0"],
                                               font=Font(name=FUENTE, bold=True, color=ROJO, size=15)))
ws.conditional_formatting.add("B7", CellIsRule(operator="greaterThan", formula=["0"],
                                               font=Font(name=FUENTE, bold=True, color=VERDE_OK, size=15)))
ws.conditional_formatting.add("B10", FormulaRule(formula=['AND(ISNUMBER(B10),B10>D10)'],
                                                 font=Font(name=FUENTE, bold=True, color=ROJO, size=15)))
ws.conditional_formatting.add("F13", CellIsRule(operator="greaterThan", formula=["0"],
                                                font=Font(name=FUENTE, bold=True, color=AMBAR, size=15)))

# --- cuenta de resultados ---
seccion(ws, "B15:F15", "CUENTA DE RESULTADOS (proyectada)")
pyg = [
    (16, "Ventas esperadas (IVA incl.)", f"={S('AC')}"),
    (17, "− Coste de producto", f"=-{S('AD')}"),
    (18, "− Envíos", f"=-{S('AE')}"),
    (19, "− Devoluciones", f"=-{S('AF')}"),
    (20, "− Comisión almacén (0,95 €/pedido)", f"=-{S('AJ')}"),
    (21, "− IVA a liquidar (neto)", f"=-{S('AG')}"),
    (22, "− Publicidad Meta", f"=-{S('AH')}"),
    (23, "Beneficio de campaña", "=SUM(F16:F22)"),
    (24, "− Otros gastos", f"=-{GASTOS_PERIODO}"),
    (25, "BENEFICIO NETO", "=F23+F24"),
]
celda(ws, "E16", "% ventas", color=GRIS, size=8, align="right")
for f, etq, formula in pyg:
    total = f in (23, 25)
    ws.merge_cells(f"B{f}:D{f}")
    celda(ws, f"B{f}", etq, bold=total, color=TINTA if total else "000000",
          fondo=FONDO_CARD if total else None)
    celda(ws, f"F{f}", formula, fmt=EUR, bold=total, fondo=FONDO_CARD if total else None, borde=BORDE_INF)
    if f > 16:
        celda(ws, f"E{f}", f'=IF($F$16>0,F{f}/$F$16,"")', fmt='0.0%;-0.0%;"-"', color=GRIS, size=9,
              fondo=FONDO_CARD if total else None)
ws.merge_cells("B26:D26")
celda(ws, "B26", "Margen neto sobre ventas", italic=True, color=GRIS)
celda(ws, "F26", '=IF(F16>0,F25/F16,"")', fmt=PCT, bold=True)

# --- objetivo ---
seccion(ws, "H15:M15", "CAMINO A 10.000 €/MES")
obj = [
    (16, "Objetivo de beneficio neto mensual", "=OBJETIVO_MES", EUR0),
    (17, "Beneficio neto por día (periodo)", "=F25/I3", EUR),
    (18, "Ritmo a 30 días", "=L17*30", EUR0),
    (19, "% del objetivo", '=IF(L16>0,L18/L16,"")', PCT),
    (20, "Beneficio esperado por pedido (antes de Ads)", "=BEN_PEDIDO", EUR),
    (21, "CPA actual", "=B10", EUR),
    (22, "Margen por pedido después de Ads", '=IF(ISNUMBER(L21),L20-L21*(1+RECARGO_ADS),"")', EUR),
    (23, "Pedidos/día necesarios para el objetivo",
     '=IF(NOT(ISNUMBER(L22)),"",IF(L22<=0,"Baja el CPA",ROUNDUP((L16-F24/I3*30)/30/L22,0)))', ENTERO),
    (24, "Inversión diaria en Ads a ese CPA", '=IF(ISNUMBER(L23),L23*L21,"")', EUR0),
]
for f, etq, formula, fmt in obj:
    ws.merge_cells(f"H{f}:K{f}")
    celda(ws, f"H{f}", etq, bold=(f in (19, 23)))
    ws.merge_cells(f"L{f}:M{f}")
    celda(ws, f"L{f}", formula, fmt=fmt, bold=(f in (19, 23)), align="right", borde=BORDE_INF)
ws.conditional_formatting.add("L19", CellIsRule(operator="greaterThanOrEqual", formula=["1"],
                                                font=Font(name=FUENTE, bold=True, color=VERDE_OK)))

# --- últimos 14 días ---
seccion(ws, "B27:G27", "ÚLTIMOS 14 DÍAS (hasta la fecha 'Hasta')")
cabecera(ws, 28, 2, ["Fecha", "Pedidos", "Gasto Ads", "CPA", "Beneficio proyectado", "Tasa entrega"])


def S_dia(col, celda_fecha):
    return f"SUMIFS(Diario!${col}${F0}:${col}${FN},Diario!$A${F0}:$A${FN},{celda_fecha})"


for i in range(14):
    r = 29 + i
    fd = f"$B{r}"
    ent = f"({S_dia('D', fd)}+{S_dia('E', fd)})"
    res = f"({S_dia('D', fd)}+{S_dia('E', fd)}+{S_dia('F', fd)}+{S_dia('G', fd)})"
    celda(ws, f"B{r}", f"={HASTA}-{13 - i}", fmt="ddd dd/mm", borde=BORDE, align="center")
    celda(ws, f"C{r}", f"={S_dia('M', fd)}", fmt=ENTERO, borde=BORDE)
    celda(ws, f"D{r}", f"={S_dia('K', fd)}", fmt=EUR, borde=BORDE)
    celda(ws, f"E{r}", f'=IF(C{r}>0,D{r}/C{r},"")', fmt=EUR, borde=BORDE)
    celda(ws, f"F{r}", f"={S_dia('V', fd)}", fmt=EUR, borde=BORDE, bold=True)
    celda(ws, f"G{r}", f'=IF({res}>0,{ent}/{res},"")', fmt=PCT, borde=BORDE)

graf = BarChart()
graf.type = "col"
graf.title = "Beneficio proyectado vs gasto en Ads (€)"
graf.style = 10
graf.add_data(Reference(ws, min_col=6, min_row=28, max_row=42), titles_from_data=True)
graf.set_categories(Reference(ws, min_col=2, min_row=29, max_row=42))
graf.series[0].graphicalProperties.solidFill = TINTA
linea = LineChart()
linea.add_data(Reference(ws, min_col=4, min_row=28, max_row=42), titles_from_data=True)
linea.series[0].graphicalProperties.line.solidFill = ORO
linea.series[0].graphicalProperties.line.width = 28000
linea.series[0].smooth = False
graf += linea
graf.x_axis.number_format = "dd/mm"
graf.y_axis.number_format = '#,##0 "€"'
graf.x_axis.delete = False
graf.y_axis.delete = False
graf.legend.position = "b"
graf.height, graf.width = 7.6, 15.5
ws.add_chart(graf, "H27")

# --- resumen mensual ---
seccion(ws, "B45:M45", "RESUMEN MENSUAL")
cabecera(ws, 46, 2, ["Mes", "Pedidos", "Gasto Ads", "CPA", "Facturación bruta", "Entregados", "Devueltos",
                     "Tasa entrega", "Beneficio campaña", "Otros gastos", "Beneficio neto", "% objetivo"])
ws.row_dimensions[46].height = 28


def S_mes(col, r):
    return (f"SUMIFS(Diario!${col}${F0}:${col}${FN},Diario!$A${F0}:$A${FN},\">=\"&$B{r},"
            f"Diario!$A${F0}:$A${FN},\"<=\"&EOMONTH($B{r},0))")


for i in range(12):
    r = 47 + i
    celda(ws, f"B{r}", f"=DATE(YEAR(FECHA_INICIO),MONTH(FECHA_INICIO)+{i},1)", fmt="mmm yyyy",
          borde=BORDE, align="center")
    celda(ws, f"C{r}", f"={S_mes('M', r)}", fmt=ENTERO, borde=BORDE)
    celda(ws, f"D{r}", f"={S_mes('K', r)}", fmt=EUR0, borde=BORDE)
    celda(ws, f"E{r}", f'=IF(C{r}>0,D{r}/C{r},"")', fmt=EUR, borde=BORDE)
    celda(ws, f"F{r}", f"={S_mes('O', r)}", fmt=EUR0, borde=BORDE)
    celda(ws, f"G{r}", f"={S_mes('D', r)}+{S_mes('E', r)}", fmt=ENTERO, borde=BORDE)
    celda(ws, f"H{r}", f"={S_mes('F', r)}+{S_mes('G', r)}", fmt=ENTERO, borde=BORDE)
    celda(ws, f"I{r}", f'=IF(G{r}+H{r}>0,G{r}/(G{r}+H{r}),"")', fmt=PCT, borde=BORDE)
    celda(ws, f"J{r}", f"={S_mes('V', r)}", fmt=EUR0, borde=BORDE)
    celda(ws, f"K{r}", (f"=-SUMIFS(Gastos!$D${G0}:$D${GN},Gastos!$A${G0}:$A${GN},\">=\"&$B{r},"
                         f"Gastos!$A${G0}:$A${GN},\"<=\"&EOMONTH($B{r},0))"), fmt=EUR0, borde=BORDE)
    celda(ws, f"L{r}", f"=J{r}+K{r}", fmt=EUR0, borde=BORDE, bold=True)
    celda(ws, f"M{r}", f'=IF(OBJETIVO_MES>0,L{r}/OBJETIVO_MES,"")', fmt=PCT, borde=BORDE)
celda(ws, "B59", "TOTAL", bold=True, color=TINTA, align="center", fondo=FONDO_CARD)
for letra in "CDFGHJKL":
    fmt = ENTERO if letra in "CGH" else EUR0
    celda(ws, f"{letra}59", f"=SUM({letra}47:{letra}58)", fmt=fmt, bold=True, fondo=FONDO_CARD)
celda(ws, "E59", '=IF(C59>0,D59/C59,"")', fmt=EUR, bold=True, fondo=FONDO_CARD)
celda(ws, "I59", '=IF(G59+H59>0,G59/(G59+H59),"")', fmt=PCT, bold=True, fondo=FONDO_CARD)
celda(ws, "M59", "", fondo=FONDO_CARD)
ws.conditional_formatting.add("L47:L58", CellIsRule(operator="lessThan", formula=["0"],
                                                    font=Font(name=FUENTE, bold=True, color=ROJO)))
ws.conditional_formatting.add("M47:M58", CellIsRule(operator="greaterThanOrEqual", formula=["1"],
                                                    fill=fill("D5F0DD")))

graf2 = BarChart()
graf2.type = "col"
graf2.title = "Beneficio neto mensual vs objetivo (€)"
graf2.style = 10
graf2.add_data(Reference(ws, min_col=12, min_row=46, max_row=58), titles_from_data=True)
graf2.set_categories(Reference(ws, min_col=2, min_row=47, max_row=58))
graf2.series[0].graphicalProperties.solidFill = TINTA
obj_linea = LineChart()
celda(ws, "O46", "Objetivo", color=GRIS, size=8)
for i in range(12):
    celda(ws, f"O{47 + i}", "=OBJETIVO_MES", fmt=EUR0, color=GRIS, size=8)
ws.column_dimensions["O"].hidden = True
obj_linea.add_data(Reference(ws, min_col=15, min_row=46, max_row=58), titles_from_data=True)
obj_linea.series[0].graphicalProperties.line.solidFill = ORO
obj_linea.series[0].graphicalProperties.line.dashStyle = "dash"
obj_linea.series[0].smooth = False
graf2 += obj_linea
graf2.x_axis.number_format = "mmm yy"
graf2.y_axis.number_format = '#,##0 "€"'
graf2.x_axis.delete = False
graf2.y_axis.delete = False
graf2.legend.position = "b"
graf2.height, graf2.width = 8, 31
ws.add_chart(graf2, "B61")
ws.freeze_panes = "A4"

# =====================================================================
# GUÍA
# =====================================================================
ws = guia
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 120
titulo(ws, "B1", "Cómo usar este libro")
lineas = [
    ("h", "Hojas"),
    ("t", "Dashboard: tus números del periodo que elijas (celdas Desde / Hasta). Todo se calcula solo."),
    ("t", "Diario: una fila por día de pedido, con pedidos y estados de cada pack."),
    ("t", "Gastos: Shopify, apps, muestras, gestoría... todo lo que no es producto, envío ni Ads."),
    ("t", "Config: precios, costes de tus 2 packs y supuestos. Si el almacén cambia tarifas, se cambia aquí."),
    ("h", "Rutina diaria (2 minutos)"),
    ("t", "1. Descarga el Excel de pedidos del almacén y pásaselo a Claude. Lo importa entero: pedidos por pack, "
          "estados, incidencias y coste de envío real. Sin datos personales: solo totales por día."),
    ("t", "2. Si no tienes el export a mano, basta con decir \"hoy 9 Pack 1 y 4 Pack 2\" o apuntarlo en el dashboard web."),
    ("t", "3. El gasto de Meta Ads lo trae Claude del conector (y el dashboard web lo lee en directo)."),
    ("h", "Estados del almacén"),
    ("t", "Preparado, En ruta e Incidencia = pendientes. Una incidencia NO es venta perdida hasta que pasa a Devuelto."),
    ("t", "Entregado = cobrado. Devuelto = coste de devolución. Rechazado y carritos abandonados no cuentan como pedido. "
          "Los pedidos de test se excluyen."),
    ("t", "Por qué por fecha de pedido: así el gasto de Ads de un día se compara con lo que esos pedidos acaban dejando."),
    ("h", "Los tres beneficios"),
    ("t", "Liquidación almacén: PVP − coste total del pedido de lo ya entregado. Es lo que te paga el almacén y "
          "cuadra con su panel."),
    ("t", "Beneficio cobrado (real): liquidación − IVA − devoluciones − Ads. Es caja. Los días recientes "
          "salen en negativo porque los pedidos aún están pendientes."),
    ("t", "Beneficio proyectado: lo mismo, suponiendo que los pendientes se entregan a tu tasa de entrega. "
          "Es el número para decidir si escalar o cortar."),
    ("h", "Semáforo de campaña (CPA = gasto en Ads ÷ pedidos confirmados)"),
    ("t", "ESCALAR: CPA ≤ 75% del CPA break-even. Hay margen de sobra: sube presupuesto un 20-30% cada 48 h."),
    ("t", "VIGILAR: CPA entre el 75% y el 100% del break-even. Ganas poco: prueba creatividades nuevas antes de escalar."),
    ("t", "CORTAR: CPA por encima del break-even. Cada pedido te cuesta dinero: apaga o cambia el ángulo."),
    ("h", "Supuestos que debes confirmar (Config)"),
    ("t", "• Coste de un pedido devuelto = 8,06 € (pierdes el envío, el producto vuelve). Si tu almacén cobra "
          "también el retorno, súbelo."),
    ("t", "• El envío no lleva IVA deducible (el export lo da 'sin IVA' y no le suma IVA al coste total)."),
    ("t", "• Meta te factura sin IVA (alta en ROI). Si no, pon 21% en 'Recargo sobre el gasto en Ads'."),
    ("t", "• El IVA se descuenta del beneficio (SÍ). Si aún no facturas con IVA, ponlo en NO para ver el bruto, "
          "pero ese IVA te lo reclamará Hacienda cuando te des de alta."),
    ("h", "Colores"),
    ("t", "Azul sobre amarillo = dato que escribes tú. Negro = fórmula. Verde = viene de otra hoja."),
]
fila = 3
for tipo, texto in lineas:
    if tipo == "h":
        fila += 1
        celda(ws, f"B{fila}", texto, bold=True, color=TINTA, size=11)
    else:
        celda(ws, f"B{fila}", texto, wrap=True)
    fila += 1

wb.active = 0
wb.save(SALIDA)
print(f"Guardado: {SALIDA}")
