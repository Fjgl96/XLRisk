"""Builds two ready-to-run XLRisk test workbooks (Spanish):

    examples/XLRisk_Prueba_Basica.xlsx         price x quantity, exact results known
    examples/XLRisk_Prueba_Multivariable.xlsx  8 risk inputs, correlation, 4 outputs

Each workbook already contains the XLRisk settings sheet (seed, iterations) with
its outputs registered, exactly as Setup and Add Output would leave it: open it
with XLRisk loaded and press Run. Works with the original XLRisk and with 1.1.0.

Usage:  python examples/build_examples.py      (needs openpyxl)
"""
import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName

HERE = Path(__file__).resolve().parent
ARIAL = "Arial"
NAVY = "1F3864"
F_TITLE = Font(name=ARIAL, size=14, bold=True, color=NAVY)
F_SUB = Font(name=ARIAL, size=11, bold=True, color=NAVY)
F_HDR = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=ARIAL, size=10)
F_BOLD = Font(name=ARIAL, size=10, bold=True)
F_INPUT = Font(name=ARIAL, size=10, color="0000FF")
F_LINK = Font(name=ARIAL, size=10, color="008000")
F_NOTE = Font(name=ARIAL, size=9, italic=True, color="404040")
F_TAG = Font(name=ARIAL, size=8, bold=True, color="404040")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_OUT = PatternFill("solid", fgColor="008C00")
THIN = Side(style="thin", color="808080")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

SOLES = '"S/ "#,##0;("S/ "#,##0);-'
SOLES2 = '"S/ "#,##0.00'
USD2 = '"US$ "#,##0.00'
UNITS = "#,##0"
PCT = "0.0%"


def header(ws, row, labels, col=1):
    for j, h in enumerate(labels):
        c = ws.cell(row=row, column=col + j, value=h)
        c.font, c.fill, c.border, c.alignment = F_HDR, FILL_HDR, BOX, WRAP


def put(ws, ref, value, fmt=None, font=F_BODY, fill=None, border=True):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = BOX
    return c


def inp(ws, ref, value, fmt):
    return put(ws, ref, value, fmt, F_INPUT, FILL_INPUT)


def tag(ws, ref, text):
    c = ws[ref]
    c.value, c.font, c.alignment = text, F_TAG, Alignment(horizontal="center", vertical="center")


def xlrisk_sheet(wb, outputs, inputs, seed=2026, iterations=5000):
    """Settings sheet with the same layout and names that XLRisk's Setup creates.

    outputs: list of (address like "Modelo!$C$16", name)
    inputs:  list of (address, formula without "=") shown for reference;
             XLRisk rewrites this list on every Setup and Run.
    """
    ws = wb.create_sheet("XLRisk")
    rows = [
        (1, "Simulation Settings", None, None),
        (2, "Seed", seed, "Seed"),
        (3, "Update Screen", False, "ScreenUpdate"),
        (4, "Iterations", iterations, "Iterations"),
        (5, "Latin Hypercube", True, "LatinHypercube"),
        (6, "Calculate data tables during simulation", True, "CalcDataTables"),
        (10, "Macro to run before each iteration", "", "MacroBefore"),
        (11, "Macro to run after each iteration", "", "MacroAfter"),
        (12, "Macro to run after simulation", "", "MacroAfterSimulation"),
    ]
    for r, label, value, name in rows:
        ws.cell(row=r, column=1, value=label).font = Font(name=ARIAL, size=10, bold=(r == 1), italic=(r > 1))
        if name:
            ws.cell(row=r, column=2, value=value if value != "" else None).font = F_INPUT
            ws.defined_names[name] = DefinedName(name, attr_text=f"XLRisk!$B${r}")
    ws["D1"], ws["D3"], ws["E3"] = "Simulation Inputs", "Range", "Formula"
    ws["G1"], ws["G3"], ws["H3"] = "Simulation Outputs", "Range", "Name"
    for ref in ("D1", "D3", "E3", "G1", "G3", "H3"):
        ws[ref].font = F_BOLD
    ws.defined_names["RiskInputs"] = DefinedName("RiskInputs", attr_text="XLRisk!$D$3")
    ws.defined_names["RiskOutputs"] = DefinedName("RiskOutputs", attr_text="XLRisk!$G$3")
    for i, (addr, formula) in enumerate(inputs, 4):
        ws.cell(row=i, column=4, value=addr).font = F_BODY
        c = ws.cell(row=i, column=5, value=formula)
        c.data_type, c.font = "s", F_BODY
    for i, (addr, name) in enumerate(outputs, 4):
        ws.cell(row=i, column=7, value=f"=AddressWithSheet({addr})").font = F_BODY
        ws.cell(row=i, column=8, value=name).font = F_BODY
    for col, w in zip("ABCDEFGH", (36, 10, 3, 18, 48, 3, 18, 22)):
        ws.column_dimensions[col].width = w
    ws.sheet_properties.tabColor = "808080"
    return ws


def instructions(wb, title, lines, notes):
    ws = wb.active
    ws.title = "Instrucciones"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 115
    ws["A1"], ws["A1"].font = title, F_TITLE
    r = 3
    ws.cell(row=r, column=1, value="Cómo correrlo").font = F_SUB
    for t in lines:
        r += 1
        c = ws.cell(row=r, column=1, value=t)
        c.font, c.alignment = F_BODY, WRAP
    r += 2
    ws.cell(row=r, column=1, value="Qué observar").font = F_SUB
    for t in notes:
        r += 1
        c = ws.cell(row=r, column=1, value=t)
        c.font, c.alignment = F_BODY, WRAP
    r += 2
    ws.cell(row=r, column=1, value="Celdas amarillas con texto azul: supuestos editables. Celdas verdes: outputs "
            "ya registrados en la hoja XLRisk.").font = F_NOTE
    ws.sheet_properties.tabColor = NAVY
    return ws


# ---------------------------------------------------------------------------
# 1. Basic workbook
# ---------------------------------------------------------------------------

def build_basic(path):
    wb = Workbook()
    instructions(wb, "XLRisk: prueba básica (precio x cantidad)", [
        "1. Abre este libro con el complemento XLRisk cargado (pestaña XLRisk en la cinta).",
        "2. Ve a la hoja Modelo y pulsa XLRisk > Run. Los outputs, la semilla (2026) y las iteraciones (5,000) "
        "ya están configurados en la hoja XLRisk.",
        "3. Al terminar se crea la hoja Risk Results 1. Compara la media de Utilidad con el Chequeo exacto de la "
        "hoja Modelo, y la media de Pérdida con la probabilidad de pérdida exacta.",
        "4. Cambia los supuestos amarillos y vuelve a correr: el chequeo exacto se recalcula solo.",
    ], [
        "Utilidad esperada exacta: S/ 28,850. Con 5,000 iteraciones la media simulada debería quedar a menos de "
        "S/ 2,000 de ese valor.",
        "Probabilidad de pérdida exacta: 35%. La media del output Pérdida debería quedar entre 33% y 37%.",
        "Con Show Samples apagado, las celdas de precio y cantidad simulados muestran su valor esperado "
        "(S/ 50.25 y 9,800 unidades).",
    ])

    ws = wb.create_sheet("Modelo")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (30, 14, 14, 4, 30, 16)):
        ws.column_dimensions[col].width = w
    ws["A1"], ws["A1"].font = "Utilidad = (precio - costo variable) x cantidad - costo fijo", F_TITLE

    header(ws, 3, ["Precio: escenario", "Precio (S/)", "Probabilidad"])
    for i, (n, v, p) in enumerate([("Bajo", 45, 0.25), ("Base", 50, 0.50), ("Alto", 56, 0.25)]):
        put(ws, f"A{4 + i}", n)
        inp(ws, f"B{4 + i}", v, SOLES2)
        inp(ws, f"C{4 + i}", p, "0%")
    header(ws, 8, ["Cantidad: escenario", "Unidades", "Probabilidad"])
    for i, (n, v, p) in enumerate([("Baja", 8000, 0.30), ("Base", 10000, 0.50), ("Alta", 12000, 0.20)]):
        put(ws, f"A{9 + i}", n)
        inp(ws, f"B{9 + i}", v, UNITS)
        inp(ws, f"C{9 + i}", p, "0%")
    put(ws, "A13", "Costo variable unitario (S/)")
    inp(ws, "B13", 32, SOLES2)
    put(ws, "A14", "Costo fijo (S/)")
    inp(ws, "B14", 150000, SOLES)

    header(ws, 16, ["Modelo", "Tipo", "Valor"])
    put(ws, "A17", "Precio simulado (S/)", font=F_BOLD)
    put(ws, "C17", "=RiskDiscrete(B4:B6,C4:C6)", SOLES2, F_BOLD)
    put(ws, "A18", "Cantidad simulada", font=F_BOLD)
    put(ws, "C18", "=RiskDiscrete(B9:B11,C9:C11)", UNITS, F_BOLD)
    put(ws, "A19", "Utilidad (S/)", font=F_BOLD)
    put(ws, "C19", "=(C17-B13)*C18-B14", SOLES, Font(name=ARIAL, size=10, bold=True, color="FFFFFF"), FILL_OUT)
    put(ws, "A20", "Pérdida (1 = sí)", font=F_BOLD)
    put(ws, "C20", "=IF(C19<0,1,0)", "0", Font(name=ARIAL, size=10, bold=True, color="FFFFFF"), FILL_OUT)
    for ref, t in (("B17", "INPUT"), ("B18", "INPUT"), ("B19", "OUTPUT"), ("B20", "OUTPUT")):
        tag(ws, ref, t)

    header(ws, 3, ["Chequeo exacto", "Valor"], col=5)
    checks = [
        ("Precio esperado (S/)", "=SUMPRODUCT(B4:B6,C4:C6)", SOLES2),
        ("Cantidad esperada", "=SUMPRODUCT(B9:B11,C9:C11)", UNITS),
        ("Utilidad esperada (S/)", "=(F4-B13)*F5-B14", SOLES),
        ("Probabilidad de pérdida",
         "=SUMPRODUCT(((B4:B6-B13)*B9-B14<0)*C4:C6)*C9+SUMPRODUCT(((B4:B6-B13)*B10-B14<0)*C4:C6)*C10"
         "+SUMPRODUCT(((B4:B6-B13)*B11-B14<0)*C4:C6)*C11", PCT),
        ("Suma de probabilidades: precio", "=SUM(C4:C6)", "0%"),
        ("Suma de probabilidades: cantidad", "=SUM(C9:C11)", "0%"),
    ]
    for i, (lab, f, fmt) in enumerate(checks):
        put(ws, f"E{4 + i}", lab)
        put(ws, f"F{4 + i}", f, fmt, F_BOLD)
    ws.sheet_properties.tabColor = "008C00"

    xlrisk_sheet(wb, outputs=[("Modelo!$C$19", "Utilidad"), ("Modelo!$C$20", "Pérdida")],
                 inputs=[("Modelo!$C$17", "RiskDiscrete(B4:B6,C4:C6)"),
                         ("Modelo!$C$18", "RiskDiscrete(B9:B11,C9:C11)")])
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(path)


# ---------------------------------------------------------------------------
# 2. Multi-variable workbook
# ---------------------------------------------------------------------------

# (label, distribution, parameters [(label, value, format)], unit format)
SUPUESTOS = [
    ("Volumen anual (toneladas)", "RiskNormal", [("Media", 1200, UNITS), ("Desv. estándar", 120, UNITS)], UNITS),
    ("Precio de venta (S/ por t)", "RiskTriang",
     [("Mínimo", 11500, SOLES), ("Más probable", 12000, SOLES), ("Máximo", 12800, SOLES)], SOLES),
    ("Costo FOB (US$ por t)", "RiskPert",
     [("Mínimo", 1900, USD2), ("Más probable", 2000, USD2), ("Máximo", 2300, USD2)], USD2),
    ("Tipo de cambio (S/ por US$)", "RiskLogNorm",
     [("Media", 3.75, "0.000"), ("Desv. estándar", 0.15, "0.000")], "0.000"),
    ("Flete y aduanas (% del FOB)", "RiskUniform", [("Mínimo", 0.08, PCT), ("Máximo", 0.12, PCT)], PCT),
    ("Gastos fijos anuales (S/)", "RiskPert",
     [("Mínimo", 3000000, SOLES), ("Más probable", 3200000, SOLES), ("Máximo", 3600000, SOLES)], SOLES),
    ("Pérdida de un cliente grande (1 = sí)", "RiskBernoulli", [("Probabilidad", 0.15, PCT)], "0"),
    ("Clientes que caen en mora", "RiskBinomial", [("Clientes con crédito", 40, UNITS),
                                                   ("Probabilidad de mora", 0.05, PCT)], UNITS),
]


def build_multi(path):
    wb = Workbook()
    instructions(wb, "XLRisk: prueba multivariable (presupuesto anual de una importadora)", [
        "1. Abre este libro con el complemento XLRisk cargado.",
        "2. Ve a la hoja Modelo y pulsa XLRisk > Run. Los 4 outputs, la semilla (2026) y las 5,000 iteraciones "
        "ya están configurados.",
        "3. En la hoja Risk Results revisa: la media y el P5 / P95 de la Utilidad operativa, la media de Pérdida "
        "(probabilidad de terminar el año en pérdida) y el tornado, que ordena qué supuesto mueve más el resultado.",
        "4. Prueba cambiar la correlación volumen-precio en la hoja Supuestos (por ejemplo a 0 o a -0.8) y "
        "compara el ancho del rango de la utilidad.",
    ], [
        "El modelo combina 8 supuestos inciertos con 6 distribuciones distintas: normal, triangular, PERT, "
        "lognormal, uniforme, Bernoulli y binomial.",
        "Volumen y precio están correlacionados en -0.5 con RiskCorrmat: cuando el precio sube, el volumen tiende "
        "a bajar.",
        "El tipo de cambio usa RiskLogNorm. Como XLRisk recibe la media y desviación de LN(X), la hoja Supuestos "
        "convierte media 3.75 y desviación 0.15 a esos parámetros. Con Show Samples apagado el modelo muestra "
        "3.750, lo que confirma la conversión.",
        "Con Show Samples apagado la hoja Modelo muestra el caso base, con cada supuesto en su valor esperado: "
        "utilidad operativa de unos S/ 1.07 millones.",
        "Resultado de referencia (simulación independiente de 200,000 iteraciones): utilidad media cerca de "
        "S/ 1.06 millones, P5 cerca de -S/ 0.11 millones, P95 cerca de S/ 2.24 millones y probabilidad de "
        "pérdida cerca de 7%. XLRisk con 5,000 iteraciones debería dar valores parecidos.",
    ])

    sp = wb.create_sheet("Supuestos")
    sp.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJ", (46, 16, 16, 16, 16, 3, 16, 12, 12, 3)):
        sp.column_dimensions[col].width = w
    sp["A1"], sp["A1"].font = "Supuestos inciertos", F_TITLE
    header(sp, 3, ["Supuesto", "Distribución", "Parámetro 1", "Parámetro 2", "Parámetro 3"])
    param_cells = []
    for i, (label, dist, params, _) in enumerate(SUPUESTOS):
        r = 4 + 2 * i
        put(sp, f"A{r}", label, font=F_BOLD)
        put(sp, f"B{r}", dist)
        cells = []
        for j, (plabel, value, fmt) in enumerate(params):
            col = "CDE"[j]
            put(sp, f"{col}{r}", plabel, font=Font(name=ARIAL, size=8, color="404040"))
            inp(sp, f"{col}{r + 1}", value, fmt)
            cells.append(f"Supuestos!${col}${r + 1}")
        sp.merge_cells(f"A{r}:A{r + 1}")
        sp.merge_cells(f"B{r}:B{r + 1}")
        sp[f"A{r}"].alignment = sp[f"B{r}"].alignment = Alignment(vertical="center", wrap_text=True)
        param_cells.append(cells)

    # Lognormal conversion for the exchange rate (row 10/11 holds mean and stdev)
    sp["A21"], sp["A21"].font = "Tipo de cambio: parámetros de LN(X) para RiskLogNorm", F_SUB
    put(sp, "A22", "sigma = RAIZ(LN(1 + (desv/media)^2))")
    put(sp, "B22", "=SQRT(LN(1+(D11/C11)^2))", "0.000000", F_BOLD)
    put(sp, "A23", "mu = LN(media) - sigma^2 / 2")
    put(sp, "B23", "=LN(C11)-B22^2/2", "0.000000", F_BOLD)

    sp["A25"], sp["A25"].font = "Otros parámetros", F_SUB
    put(sp, "A26", "Volumen que se pierde si se va el cliente grande")
    inp(sp, "B26", 0.15, PCT)
    put(sp, "A27", "Pérdida por cliente en mora (S/)")
    inp(sp, "B27", 25000, SOLES)

    header(sp, 3, ["Correlación", "Volumen", "Precio"], col=7)
    put(sp, "G4", "Volumen", font=F_BOLD)
    put(sp, "G5", "Precio", font=F_BOLD)
    put(sp, "H4", 1, "0.00")
    inp(sp, "I4", -0.5, "0.00")
    put(sp, "H5", "=I4", "0.00")
    put(sp, "I5", 1, "0.00")
    put(sp, "G7", "¿Matriz válida?")
    put(sp, "H7", "=RiskIsValidCorrmat(H4:I5)", None, F_BOLD)
    sp["G8"] = "Edita solo I4; H5 la copia para mantener la simetría."
    sp["G8"].font = F_NOTE
    sp.sheet_properties.tabColor = "FFC000"

    md = wb.create_sheet("Modelo")
    md.sheet_view.showGridLines = False
    for col, w in zip("ABC", (44, 10, 18)):
        md.column_dimensions[col].width = w
    md["A1"], md["A1"].font = "Presupuesto anual: utilidad operativa", F_TITLE
    md["A2"] = "Una función de riesgo por celda (INPUT); el resto son fórmulas normales de Excel."
    md["A2"].font = F_NOTE
    header(md, 4, ["Supuesto simulado", "Tipo", "Valor"])
    corr = "Supuestos!$H$4:$I$5"
    p = param_cells
    formulas = [
        f"=RiskNormal({p[0][0]},{p[0][1]},RiskCorrmat({corr},1))",
        f"=RiskTriang({p[1][0]},{p[1][1]},{p[1][2]},RiskCorrmat({corr},2))",
        f"=RiskPert({p[2][0]},{p[2][1]},{p[2][2]})",
        "=RiskLogNorm(Supuestos!$B$23,Supuestos!$B$22)",
        f"=RiskUniform({p[4][0]},{p[4][1]})",
        f"=RiskPert({p[5][0]},{p[5][1]},{p[5][2]})",
        f"=RiskBernoulli({p[6][0]})",
        f"=RiskBinomial({p[7][0]},{p[7][1]})",
    ]
    for i, ((label, _, _, fmt), f) in enumerate(zip(SUPUESTOS, formulas)):
        r = 5 + i
        put(md, f"A{r}", label, font=F_BOLD)
        tag(md, f"B{r}", "INPUT")
        put(md, f"C{r}", f, fmt, F_BOLD)
    # C5 volume, C6 price, C7 FOB, C8 FX, C9 freight, C10 fixed, C11 client loss, C12 defaults

    header(md, 14, ["Cálculo", "Tipo", "Valor"])
    calc = [
        ("Volumen efectivo (t)", "=C5*(1-Supuestos!$B$26*C11)", UNITS, None),
        ("Ventas (S/)", "=C15*C6", SOLES, "OUTPUT"),
        ("Costo importado (S/)", "=C15*C7*C8*(1+C9)", SOLES, None),
        ("Margen bruto (S/)", "=C16-C17", SOLES, None),
        ("Gastos fijos (S/)", "=C10", SOLES, None),
        ("Pérdida por mora (S/)", "=C12*Supuestos!$B$27", SOLES, None),
        ("Utilidad operativa (S/)", "=C18-C19-C20", SOLES, "OUTPUT"),
        ("Margen operativo (%)", "=IF(C16=0,0,C21/C16)", PCT, "OUTPUT"),
        ("Pérdida (1 = sí)", "=IF(C21<0,1,0)", "0", "OUTPUT"),
    ]
    for i, (label, f, fmt, kind) in enumerate(calc):
        r = 15 + i
        put(md, f"A{r}", label, font=F_BOLD if kind else F_BODY)
        if kind:
            tag(md, f"B{r}", kind)
            put(md, f"C{r}", f, fmt, Font(name=ARIAL, size=10, bold=True, color="FFFFFF"), FILL_OUT)
        else:
            put(md, f"C{r}", f, fmt)
    md.sheet_properties.tabColor = "008C00"

    xlrisk_sheet(
        wb,
        outputs=[("Modelo!$C$21", "Utilidad operativa"), ("Modelo!$C$22", "Margen operativo"),
                 ("Modelo!$C$23", "Pérdida"), ("Modelo!$C$16", "Ventas")],
        inputs=[(f"Modelo!$C${5 + i}", f[1:]) for i, f in enumerate(formulas)],
    )
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(path)


def expected_base_case():
    """Base case with every input at its expected value (what Excel shows with Show Samples off)."""
    vol, price = 1200, (11500 + 12000 + 12800) / 3
    fob = (1900 + 4 * 2000 + 2300) / 6
    fx, freight = 3.75, 0.10
    fixed = (3000000 + 4 * 3200000 + 3600000) / 6
    client, defaults = 0.15, 40 * 0.05
    vol_ef = vol * (1 - 0.15 * client)
    sales = vol_ef * price
    cost = vol_ef * fob * fx * (1 + freight)
    profit = sales - cost - fixed - defaults * 25000
    return sales, profit


if __name__ == "__main__":
    build_basic(HERE / "XLRisk_Prueba_Basica.xlsx")
    build_multi(HERE / "XLRisk_Prueba_Multivariable.xlsx")
    s, p = expected_base_case()
    print(f"ok  base case: sales {s:,.0f}  operating profit {p:,.0f}")
