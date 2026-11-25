"""Generates the Spanish user manual of XLRisk in two formats from one source:

    docs/XLRisk_Manual_ES.xlsx   (Excel workbook, with a runnable example model)
    docs/manual-es.md            (the same text for GitHub)

Usage:  python docs/build_manual.py      (needs openpyxl)
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

VERSION = "1.1.0"
REPO = "https://github.com/Fjgl96/XLRisk"
RELEASE = f"{REPO}/releases/latest"
HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Content (single source for the workbook and the Markdown file)
# ---------------------------------------------------------------------------

INTRO = [
    "XLRisk es un complemento gratuito y de código abierto para hacer simulación Monte Carlo en Excel.",
    "Convierte un modelo determinístico (un solo escenario) en uno probabilístico: los supuestos inciertos "
    "pasan a ser distribuciones, Excel recalcula el modelo miles de veces y el resultado deja de ser un "
    "número para convertirse en un rango con su probabilidad.",
    "Ejemplo: en vez de \"la utilidad será S/ 30,000\", el modelo dice \"la utilidad esperada es S/ 28,850 "
    "y hay 35% de probabilidad de pérdida\".",
    f"Este manual corresponde a la versión {VERSION} del fork {REPO}, compatible con los libros hechos "
    "con el XLRisk original.",
]

INSTALL = [
    ("Descargar", f"Baja XLRisk.xlam desde {RELEASE} (sección Assets)."),
    ("Desbloquear", "Clic derecho en el archivo > Propiedades > marca Desbloquear > Aceptar. Sin esto Excel lo "
                    "abre en Vista protegida y no carga."),
    ("Copiar a la carpeta de complementos", "Cópialo a %APPDATA%\\Microsoft\\AddIns. Es una ubicación de "
                    "confianza de Excel, así que no vuelve a bloquearse."),
    ("Activar en Excel", "Desarrollador > Complementos de Excel > Examinar > elige XLRisk.xlam > Aceptar. Si "
                    "tenías una versión anterior, desmárcala primero. Aparece la pestaña XLRisk en la cinta."),
    ("Verificar", f"En cualquier celda escribe =XLRiskVersion(). Debe devolver {VERSION}."),
]

WORKFLOW = [
    ("1. Setup", "Crea la hoja XLRisk en el libro, con la configuración y la lista de inputs y outputs. "
                 "También abre el formulario de opciones."),
    ("2. Definir inputs", "Reemplaza cada supuesto incierto por una función de riesgo. Ejemplo: el precio 50 "
                 "pasa a =RiskTriang(45, 50, 56)."),
    ("3. Definir outputs", "Selecciona la celda resultado (utilidad, VAN, caja) y pulsa Add Output. Le pones un "
                 "nombre. Puedes tener varios outputs."),
    ("4. Iteraciones y semilla", "Elige las iteraciones en la cinta (1,000 a 10,000). Con una semilla distinta "
                 "de 0 en la hoja XLRisk, la simulación se puede reproducir exacta."),
    ("5. Run", "Corre la simulación. Stop la detiene. Al terminar crea una hoja de resultados nueva."),
]

RIBBON = [
    ("Setup", "Crea la hoja XLRisk y abre las opciones de simulación."),
    ("Show Samples", "Si está activo, las funciones de riesgo muestran valores aleatorios y cambian con F9. "
                     "Si está apagado, muestran su valor esperado (el caso base)."),
    ("Iterations", "Número de iteraciones de la simulación."),
    ("Add Output", "Registra la celda o rango seleccionado como output."),
    ("Run / Stop", "Corre o detiene la simulación."),
    ("Help / About", "Ayuda en línea y versión."),
]

# (function, syntax, parameters, value outside the simulation, when to use it, example)
FUNCTIONS = [
    ("RiskTriang", "RiskTriang(min, moda, max)",
     "Mínimo, valor más probable y máximo. Exige min < moda < max estricto.",
     "(min + moda + max) / 3",
     "Juicio experto en tres puntos: pesimista, más probable, optimista.",
     "=RiskTriang(45, 50, 56)"),
    ("RiskPert", "RiskPert(min, moda, max)",
     "Igual que RiskTriang. Exige min < moda < max estricto.",
     "(min + 4 x moda + max) / 6",
     "Como la triangular, pero con menos peso en los extremos. Plazos y costos de proyectos.",
     "=RiskPert(80, 100, 150)"),
    ("RiskNormal", "RiskNormal(media, desv)",
     "Media y desviación estándar de la variable.",
     "media",
     "Variables simétricas alrededor de un valor. Puede dar negativos.",
     "=RiskNormal(10000, 1500)"),
    ("RiskLogNorm", "RiskLogNorm(mu, sigma)",
     "OJO: mu y sigma son la media y la desviación de LN(X), no de X. Es la RiskLognorm2 de @RISK.",
     "EXP(mu + sigma^2 / 2)",
     "Variables positivas con cola a la derecha: precios, tipo de cambio, ventas. Ver Buenas prácticas para "
     "convertir media y desviación.",
     "=RiskLogNorm(LN(3.75) - 0.5*LN(1+(0.1/3.75)^2), SQRT(LN(1+(0.1/3.75)^2)))"),
    ("RiskUniform", "RiskUniform(min, max)",
     "Mínimo y máximo. Todos los valores igual de probables.",
     "(min + max) / 2",
     "Solo conoces el rango, sin un valor más probable.",
     "=RiskUniform(0.08, 0.12)"),
    ("RiskDiscrete", "RiskDiscrete(valores, probabilidades)",
     "Dos rangos del mismo tamaño. Valores en orden ascendente; probabilidades que suman 1.",
     "SUMAPRODUCTO(valores, probabilidades)",
     "Escenarios con su probabilidad: bajo, base, alto.",
     "=RiskDiscrete(A2:A4, B2:B4)"),
    ("RiskDUniform", "RiskDUniform(valores)",
     "Un rango de valores, todos con la misma probabilidad.",
     "promedio de los valores",
     "Elegir al azar entre opciones igual de probables.",
     "=RiskDUniform(A2:A6)"),
    ("RiskBernoulli", "RiskBernoulli(p)",
     "Probabilidad de que ocurra el evento (entre 0 y 1).",
     "p",
     "Evento que ocurre (1) o no (0): perder un cliente, aprobar un crédito.",
     "=RiskBernoulli(0.15)"),
    ("RiskBinomial", "RiskBinomial(n, p)",
     "Número de ensayos y probabilidad de éxito de cada uno.",
     "n x p",
     "Cuántos de n eventos independientes ocurren: clientes en mora de una cartera.",
     "=RiskBinomial(200, 0.04)"),
    ("RiskExpon", "RiskExpon(media)",
     "Media de la variable.",
     "media",
     "Tiempo entre eventos: llegadas, fallas.",
     "=RiskExpon(12)"),
    ("RiskGamma", "RiskGamma(alfa, beta)",
     "alfa = forma, beta = escala.",
     "alfa x beta",
     "Montos positivos asimétricos: siniestros, tiempos de servicio.",
     "=RiskGamma(2, 500)"),
    ("RiskErlang", "RiskErlang(k, lambda)",
     "k = forma (entero), lambda = escala. Es una Gamma con forma entera.",
     "k x lambda",
     "Tiempo hasta que ocurren k eventos.",
     "=RiskErlang(3, 10)"),
    ("RiskWeibull", "RiskWeibull(alfa, beta)",
     "alfa = forma, beta = escala.",
     "beta x GAMMA(1 + 1/alfa)",
     "Vida útil y tiempo hasta falla de activos.",
     "=RiskWeibull(1.5, 1000)"),
    ("RiskBeta", "RiskBeta(alfa, beta, [A], [B])",
     "alfa y beta = forma; A y B = mínimo y máximo (por defecto 0 y 1).",
     "A + alfa / (alfa + beta) x (B - A)",
     "Proporciones y tasas acotadas: tasa de conversión, porcentaje de recupero.",
     "=RiskBeta(2, 8)"),
    ("RiskCumul", "RiskCumul(min, max, valoresX, probAcum)",
     "Curva acumulada por tramos. valoresX crecientes y dentro de (min, max); probAcum crecientes.",
     "media de la curva definida",
     "Cuando el experto da percentiles: \"10% de que sea menos de X\".",
     "=RiskCumul(0, 100, B2:B4, C2:C4)"),
    ("RiskCorrmat", "RiskCorrmat(matriz, posición)",
     "Se pone como ÚLTIMO argumento de otra función. matriz = rango de correlaciones; posición = fila de "
     "esa variable en la matriz.",
     "no aplica",
     "Hacer que dos o más supuestos se muevan juntos (precio y cantidad, ventas de dos productos).",
     "=RiskNormal(100, 15, RiskCorrmat($H$2:$I$3, 1))"),
    ("RiskIsValidCorrmat", "RiskIsValidCorrmat(matriz)",
     "Rango de correlaciones.",
     "VERDADERO o FALSO",
     "Verificar que la matriz sea válida (simétrica, diagonal 1, semidefinida positiva).",
     "=RiskIsValidCorrmat(H2:I3)"),
    ("RiskCorrectCorrmat", "RiskCorrectCorrmat(matriz)",
     "Rango de correlaciones. Se ingresa como fórmula matricial del mismo tamaño.",
     "matriz corregida",
     "Corregir una matriz que no es válida, con el menor cambio posible.",
     "=RiskCorrectCorrmat(H2:I3)"),
    ("RiskSCorrel", "RiskSCorrel(rango1, rango2)",
     "Dos rangos del mismo tamaño.",
     "correlación de Spearman",
     "Medir la correlación de rangos entre dos series.",
     "=RiskSCorrel(A2:A101, B2:B101)"),
    ("XLRiskVersion", "XLRiskVersion()",
     "Sin argumentos.",
     f"\"{VERSION}\"",
     "Confirmar qué versión del complemento está cargada.",
     "=XLRiskVersion()"),
]

SETTINGS = [
    ("Seed", "0", "Semilla. Con 0 cada corrida usa números distintos. Con cualquier otro número la simulación "
                  "se reproduce exacta."),
    ("Update Screen", "FALSO", "Actualizar la pantalla en cada iteración. Apagado es mucho más rápido."),
    ("Iterations", "1000", "Iteraciones. También se cambia desde la cinta."),
    ("Latin Hypercube", "VERDADERO", "Muestreo estratificado. Converge más rápido que el muestreo simple; "
                  "conviene dejarlo activo."),
    ("Calculate data tables", "VERDADERO", "Recalcular las tablas de datos de Excel en cada iteración."),
    ("Macro before / after", "(vacío)", "Nombre de una macro a correr antes o después de cada iteración, "
                  "o al terminar la simulación."),
    ("LegacyRnd (opcional)", "no existe", "Si agregas una celda con nombre LegacyRnd = VERDADERO, se usa el "
                  "generador Rnd() de VBA del XLRisk original, para reproducir resultados antiguos."),
]

RESULTS = [
    ("Risk Outputs", "Los valores de cada output en cada iteración, junto con los inputs que los generaron."),
    ("Simulation Statistics", "Media, mediana, moda, desviación estándar, varianza, curtosis, asimetría, mínimo, "
                              "máximo, rango y conteo de iteraciones con error."),
    ("Risk Output Percentiles", "Tabla de percentiles de cada output (P5, P50, P95, etc.)."),
    ("Calculate Percent Rank", "Escribes un valor y te dice qué porcentaje de iteraciones quedó por debajo. "
                               "Sirve para probabilidad de pérdida o de no llegar a la meta."),
    ("Correlations", "Correlación de cada input con cada output."),
    ("Output Cumulative Distributions", "Curva acumulada de cada output."),
    ("Output Distributions", "Histograma de cada output (Excel 2016 o superior)."),
    ("Risk Tornado Diagram", "Qué inputs mueven más cada output, ordenados por correlación. Es el gráfico que "
                             "se lleva a comité."),
]

PRACTICES = [
    ("RiskLogNorm usa los parámetros del logaritmo",
     "Para una variable con media m y desviación s: sigma = RAIZ(LN(1 + (s/m)^2)) y "
     "mu = LN(m) - sigma^2/2. Luego =RiskLogNorm(mu, sigma). Si pones m y s directo, simula EXP(m), que "
     "suele ser un número absurdo."),
    ("Triangular y PERT exigen min < moda < max",
     "Si el valor más probable es igual al mínimo o al máximo devuelven #VALOR!. Usa un mínimo apenas menor."),
    ("RiskDiscrete exige valores ascendentes",
     "Ordena los escenarios de menor a mayor y verifica que las probabilidades sumen exactamente 1."),
    ("Una función de riesgo por celda",
     "XLRisk detecta los inputs buscando las funciones en las fórmulas. Pon cada distribución en su propia "
     "celda y referencia esa celda en el resto del modelo."),
    ("Cuántas iteraciones",
     "1,000 para explorar; 5,000 a 10,000 para presentar. Si la media cambia mucho entre dos corridas con "
     "semillas distintas, faltan iteraciones."),
    ("Fija la semilla para presentar",
     "Con una semilla fija, quien revise el modelo obtiene exactamente los mismos números."),
    ("Matrices de correlación",
     "Deben ser simétricas, con 1 en la diagonal y coherentes entre sí. Verifica con RiskIsValidCorrmat y "
     "corrige con RiskCorrectCorrmat."),
    ("Show Samples apagado para revisar el modelo",
     "Así cada función muestra su valor esperado y el modelo se ve como el caso base."),
]

TESTS = [
    "El complemento trae 20 pruebas automáticas que verifican el generador aleatorio, el muestreo y una "
    "simulación completa contra resultados exactos.",
    "Para correrlas: Alt+F8, escribe XLRisk.xlam!XLRiskSelfTests en Nombre de la macro y pulsa Ejecutar.",
    "Tarda alrededor de un minuto. Abre un libro con cada prueba en PASS (verde) o FAIL (rojo). No guarda nada.",
    "Conviene correrlas después de instalar una versión nueva.",
]

COMPARE = [
    ("Distribuciones", "Más de 90", "15 (ver hoja Funciones)"),
    ("Latin Hypercube y correlaciones (Iman-Conover)", "Sí", "Sí"),
    ("Semilla reproducible", "Sí", f"Sí, desde la versión {VERSION}"),
    ("Histogramas, percentiles, tornado", "Sí", "Sí"),
    ("RiskOutput dentro de la fórmula", "Sí", "No: se usa el botón Add Output (roadmap Fase 1)"),
    ("RiskMean, RiskPercentile, RiskTarget en celda", "Sí", "No: se leen de la hoja de resultados (Fase 1)"),
    ("RiskLognorm(media, desv) de la variable", "Sí", "No: RiskLogNorm usa los parámetros de LN(X)"),
    ("Escenarios en una corrida (RiskSimtable)", "Sí", "No (Fase 2)"),
    ("Ajuste de distribuciones a datos", "Sí", "No (Fase 3)"),
    ("Precio", "US$ 2,900 a 3,500 por licencia", "Gratis, código abierto (licencia MIT)"),
]

EXAMPLE_STEPS = [
    "1. Asegúrate de que XLRisk esté cargado (pestaña XLRisk en la cinta).",
    "2. XLRisk > Setup. Se crea la hoja XLRisk.",
    "3. Selecciona la celda Utilidad (C16) y pulsa Add Output. Nombre: Utilidad.",
    "4. Opcional: selecciona C17 (Pérdida) y pulsa Add Output. Su media es la probabilidad de pérdida.",
    "5. Iteraciones: 5000. En la hoja XLRisk pon una semilla, por ejemplo 2026.",
    "6. Run. Compara la media de la hoja de resultados con el Chequeo exacto (columna F).",
]

# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

ARIAL = "Arial"
NAVY = "1F3864"
F_TITLE = Font(name=ARIAL, size=16, bold=True, color=NAVY)
F_SUB = Font(name=ARIAL, size=11, bold=True, color=NAVY)
F_HDR = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=ARIAL, size=10)
F_BOLD = Font(name=ARIAL, size=10, bold=True)
F_CODE = Font(name="Consolas", size=10, color="000000")
F_INPUT = Font(name=ARIAL, size=10, color="0000FF")
F_LINK = Font(name=ARIAL, size=10, color="0000FF", underline="single")
F_WARN = Font(name=ARIAL, size=10, bold=True, color="C00000")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_OUT = PatternFill("solid", fgColor="008C00")
THIN = Side(style="thin", color="808080")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

SHEETS = [
    ("Inicio", "Qué es XLRisk y cómo usar este manual"),
    ("Instalación", "Descargar, desbloquear y activar el complemento"),
    ("Paso a paso", "El flujo de una simulación y los botones de la cinta"),
    ("Funciones", "Todas las funciones, con sintaxis y ejemplos"),
    ("Configuración", "La hoja XLRisk"),
    ("Resultados", "Cómo leer la hoja de resultados"),
    ("Ejemplo", "Modelo precio x cantidad listo para simular"),
    ("Buenas prácticas", "Errores comunes y cómo evitarlos"),
    ("Pruebas", "Verificar la instalación con las autopruebas"),
    ("XLRisk vs @RISK", "Qué tiene y qué le falta frente a @RISK"),
]


def title(ws, text, subtitle=None):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    if subtitle:
        ws["A2"] = subtitle
        ws["A2"].font = Font(name=ARIAL, size=10, italic=True, color="404040")
    ws["A3"] = "Volver al índice"
    ws["A3"].hyperlink = "#'Inicio'!A1"
    ws["A3"].font = F_LINK
    ws.sheet_view.showGridLines = False


def table(ws, row, headers, rows, widths=None, code_cols=()):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font, c.fill, c.border, c.alignment = F_HDR, FILL_HDR, BOX, WRAP
    for i, r in enumerate(rows, 1):
        for j, v in enumerate(r, 1):
            c = ws.cell(row=row + i, column=j)
            c.value = v
            if isinstance(v, str) and v.startswith("="):
                c.data_type = "s"          # show the formula as text
            c.font = F_CODE if j in code_cols else (F_BOLD if j == 1 else F_BODY)
            c.border, c.alignment = BOX, WRAP
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + j)].width = w
    return row + len(rows) + 2


def paragraphs(ws, row, items, width=110):
    ws.column_dimensions["A"].width = width
    for t in items:
        c = ws.cell(row=row, column=1, value=t)
        c.font, c.alignment = F_BODY, WRAP
        row += 1
    return row + 1


def build_workbook(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Inicio"

    # Inicio
    title(ws, f"XLRisk {VERSION}: manual en español", "Simulación Monte Carlo en Excel, gratis y open source")
    ws["A3"] = None
    r = paragraphs(ws, 5, INTRO)
    ws.cell(row=r, column=1, value="Índice").font = F_SUB
    r += 1
    for name, desc in SHEETS[1:]:
        c = ws.cell(row=r, column=1, value=f"{name}: {desc}")
        c.hyperlink = f"#'{name}'!A1"
        c.font = F_LINK
        r += 1
    r += 1
    ws.cell(row=r, column=1, value=f"Repositorio, descargas y código fuente: {REPO}").font = F_BODY
    ws.cell(row=r + 1, column=1, value="Licencia MIT. Basado en XLRisk de pyscripter "
                                       "(github.com/pyscripter/XLRisk).").font = F_BODY

    # Instalación
    ws = wb.create_sheet("Instalación")
    title(ws, "Instalación", SHEETS[1][1])
    r = table(ws, 5, ["Paso", "Qué hacer"], INSTALL, widths=[30, 100])
    ws.cell(row=r, column=1, value="Si algo falla").font = F_SUB
    table(ws, r + 1, ["Síntoma", "Solución"], [
        ("\"Este tipo de archivo no es compatible con la Vista protegida\"",
         "Falta el paso Desbloquear, o el archivo está fuera de una ubicación de confianza."),
        ("=XLRiskVersion() da #¿NOMBRE?",
         "Está cargada una versión anterior o ninguna. Revisa Complementos de Excel."),
        ("No aparece la pestaña XLRisk",
         "El complemento no está marcado en Complementos de Excel, o las macros están deshabilitadas "
         "sin notificación."),
    ], widths=[30, 100])

    # Paso a paso
    ws = wb.create_sheet("Paso a paso")
    title(ws, "Paso a paso", SHEETS[2][1])
    r = table(ws, 5, ["Paso", "Qué hacer"], WORKFLOW, widths=[26, 100])
    ws.cell(row=r, column=1, value="Botones de la pestaña XLRisk").font = F_SUB
    table(ws, r + 1, ["Botón", "Qué hace"], RIBBON, widths=[26, 100])

    # Funciones
    ws = wb.create_sheet("Funciones")
    title(ws, "Funciones", SHEETS[3][1])
    ws["A4"] = ("Todas las funciones de distribución aceptan RiskCorrmat como último argumento opcional. "
                "Fuera de la simulación devuelven el valor de la columna D.")
    ws["A4"].font = F_BODY
    rows = [f[1:] for f in FUNCTIONS]
    table(ws, 6, ["Sintaxis", "Parámetros", "Valor fuera de la simulación", "Cuándo usarla", "Ejemplo"],
          rows, widths=[34, 48, 26, 46, 44], code_cols=(1, 5))
    ws.freeze_panes = "A7"
    # highlight the RiskLogNorm warning
    for row in ws.iter_rows(min_row=7, max_row=6 + len(rows)):
        if row[0].value.startswith("RiskLogNorm"):
            row[1].font = F_WARN

    # Configuración
    ws = wb.create_sheet("Configuración")
    title(ws, "Configuración: la hoja XLRisk", SHEETS[4][1])
    ws["A4"] = ("Setup crea la hoja XLRisk. Los valores se editan ahí o desde el formulario de Setup. "
                "La hoja también lista los inputs detectados y los outputs registrados.")
    ws["A4"].font = F_BODY
    table(ws, 6, ["Ajuste", "Valor inicial", "Qué controla"], SETTINGS, widths=[26, 14, 100])

    # Resultados
    ws = wb.create_sheet("Resultados")
    title(ws, "Cómo leer los resultados", SHEETS[5][1])
    ws["A4"] = ("Cada Run crea una hoja de resultados nueva, con un bloque de navegación arriba a la "
                "izquierda que lleva a cada sección.")
    ws["A4"].font = F_BODY
    table(ws, 6, ["Sección", "Qué muestra"], RESULTS, widths=[32, 100])

    # Ejemplo
    build_example(wb.create_sheet("Ejemplo"))

    # Buenas prácticas
    ws = wb.create_sheet("Buenas prácticas")
    title(ws, "Buenas prácticas y errores comunes", SHEETS[7][1])
    table(ws, 5, ["Tema", "Qué hacer"], PRACTICES, widths=[40, 100])

    # Pruebas
    ws = wb.create_sheet("Pruebas")
    title(ws, "Autopruebas", SHEETS[8][1])
    paragraphs(ws, 5, TESTS)

    # Comparación
    ws = wb.create_sheet("XLRisk vs @RISK")
    title(ws, "XLRisk frente a @RISK", SHEETS[9][1])
    r = table(ws, 5, ["Capacidad", "@RISK", f"XLRisk {VERSION}"], COMPARE, widths=[46, 30, 56])
    ws.cell(row=r, column=1, value="@RISK es marca comercial de Lumivero. XLRisk es un proyecto independiente.").font = F_BODY

    for sheet in wb.worksheets:
        sheet.sheet_properties.tabColor = NAVY if sheet.title != "Ejemplo" else "008C00"
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def build_example(ws):
    title(ws, "Ejemplo: utilidad = (precio - costo variable) x cantidad - costo fijo", SHEETS[6][1])
    ws.column_dimensions["A"].width = 28
    for col, w in zip("BCDEF", (14, 14, 4, 30, 16)):
        ws.column_dimensions[col].width = w

    def hdr(row, labels, col=1):
        for j, h in enumerate(labels):
            c = ws.cell(row=row, column=col + j, value=h)
            c.font, c.fill, c.border, c.alignment = F_HDR, FILL_HDR, BOX, WRAP

    def inp(cell, value, fmt):
        c = ws[cell]
        c.value, c.font, c.fill, c.border, c.number_format = value, F_INPUT, FILL_INPUT, BOX, fmt

    hdr(5, ["Precio: escenario", "Precio (S/)", "Probabilidad"])
    for i, (n, v, p) in enumerate([("Bajo", 45, 0.25), ("Base", 50, 0.50), ("Alto", 56, 0.25)]):
        r = 6 + i
        ws.cell(row=r, column=1, value=n).font = F_BODY
        inp(f"B{r}", v, '"S/ "#,##0.00')
        inp(f"C{r}", p, "0%")
    hdr(9, ["Cantidad: escenario", "Unidades", "Probabilidad"])
    for i, (n, v, p) in enumerate([("Baja", 8000, 0.30), ("Base", 10000, 0.50), ("Alta", 12000, 0.20)]):
        r = 10 + i
        ws.cell(row=r, column=1, value=n).font = F_BODY
        inp(f"B{r}", v, "#,##0")
        inp(f"C{r}", p, "0%")

    ws["A14"], ws["A15"] = "Costo variable unitario (S/)", "Costo fijo (S/)"
    inp("B14", 32, '"S/ "#,##0.00')
    inp("B15", 150000, '"S/ "#,##0')
    for c in ("A14", "A15"):
        ws[c].font = F_BODY

    # Simulation model
    ws["A16"], ws["A17"] = "Utilidad (S/)", "Pérdida (1 = sí)"
    ws["A18"], ws["A19"] = "Precio simulado", "Cantidad simulada"
    for c in ("A16", "A17", "A18", "A19"):
        ws[c].font = F_BOLD
    ws["C18"] = "=RiskDiscrete(B6:B8,C6:C8)"
    ws["C19"] = "=RiskDiscrete(B10:B12,C10:C12)"
    ws["C16"] = "=(C18-B14)*C19-B15"
    ws["C17"] = "=IF(C16<0,1,0)"
    for c, fmt in (("C16", '"S/ "#,##0;("S/ "#,##0);-'), ("C17", "0"), ("C18", '"S/ "#,##0.00'), ("C19", "#,##0")):
        ws[c].number_format, ws[c].font, ws[c].border = fmt, F_BOLD, BOX
    ws["C16"].fill, ws["C16"].font = FILL_OUT, Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
    ws["B16"], ws["B17"] = "OUTPUT", "OUTPUT"
    ws["B18"], ws["B19"] = "INPUT", "INPUT"
    for c in ("B16", "B17", "B18", "B19"):
        ws[c].font = Font(name=ARIAL, size=9, bold=True, color="404040")
        ws[c].alignment = Alignment(horizontal="right")

    # Exact check (independent of XLRisk)
    hdr(5, ["Chequeo exacto", "Valor"], col=5)
    checks = [
        ("Precio esperado (S/)", "=SUMPRODUCT(B6:B8,C6:C8)", '"S/ "#,##0.00'),
        ("Cantidad esperada", "=SUMPRODUCT(B10:B12,C10:C12)", "#,##0"),
        ("Utilidad esperada (S/)", "=(F6-B14)*F7-B15", '"S/ "#,##0'),
        ("Probabilidad de pérdida",
         "=SUMPRODUCT(((B6:B8-B14)*B10-B15<0)*C6:C8)*C10+SUMPRODUCT(((B6:B8-B14)*B11-B15<0)*C6:C8)*C11"
         "+SUMPRODUCT(((B6:B8-B14)*B12-B15<0)*C6:C8)*C12", "0.0%"),
    ]
    for i, (lab, f, fmt) in enumerate(checks):
        r = 6 + i
        ws.cell(row=r, column=5, value=lab).font = F_BODY
        c = ws.cell(row=r, column=6, value=f)
        c.number_format, c.font, c.border = fmt, F_BOLD, BOX
    ws["E11"] = "Suma de probabilidades: precio"
    ws["F11"] = "=SUM(C6:C8)"
    ws["E12"] = "Suma de probabilidades: cantidad"
    ws["F12"] = "=SUM(C10:C12)"
    for c in ("E11", "E12"):
        ws[c].font = F_BODY
    for c in ("F11", "F12"):
        ws[c].number_format, ws[c].font = "0%", F_BODY

    ws["A22"] = "Cómo correr el ejemplo"
    ws["A22"].font = F_SUB
    for i, t in enumerate(EXAMPLE_STEPS):
        c = ws.cell(row=23 + i, column=1, value=t)
        c.font = F_BODY
    ws["A30"] = ("Celdas con fondo amarillo y texto azul: supuestos que puedes cambiar. Con 5,000 iteraciones la "
                 "media simulada queda a unos S/ 2,000 del valor exacto y la probabilidad de pérdida a unos 2 puntos.")
    ws["A30"].font = Font(name=ARIAL, size=9, italic=True, color="404040")


# ---------------------------------------------------------------------------
# Markdown
# ---------------------------------------------------------------------------

def md_table(headers, rows):
    out = ["| " + " | ".join(headers) + " |", "| " + " | ".join("---" for _ in headers) + " |"]
    for r in rows:
        cells = []
        for v in r:
            s = str(v).replace("|", "\\|")
            cells.append(f"`{s}`" if s.startswith("=") or s.startswith("Risk") or s.startswith("XLRisk") else s)
        out.append("| " + " | ".join(cells) + " |")
    return "\n".join(out)


def build_markdown(path):
    lines = [f"# XLRisk {VERSION}: manual en español", ""]
    lines += [p + "\n" for p in INTRO]
    lines += ["También está como libro de Excel, con el ejemplo listo para simular: "
              "[XLRisk_Manual_ES.xlsx](XLRisk_Manual_ES.xlsx).", ""]
    lines += ["## Instalación", "", md_table(["Paso", "Qué hacer"], INSTALL), ""]
    lines += ["## Paso a paso", "", md_table(["Paso", "Qué hacer"], WORKFLOW), ""]
    lines += ["### Botones de la pestaña XLRisk", "", md_table(["Botón", "Qué hace"], RIBBON), ""]
    lines += ["## Funciones", "",
              "Todas las funciones de distribución aceptan `RiskCorrmat` como último argumento opcional. "
              "Fuera de la simulación devuelven el valor de la columna \"Fuera de la simulación\".", "",
              md_table(["Sintaxis", "Parámetros", "Fuera de la simulación", "Cuándo usarla", "Ejemplo"],
                       [f[1:] for f in FUNCTIONS]), ""]
    lines += ["## Configuración: la hoja XLRisk", "",
              md_table(["Ajuste", "Valor inicial", "Qué controla"], SETTINGS), ""]
    lines += ["## Cómo leer los resultados", "", md_table(["Sección", "Qué muestra"], RESULTS), ""]
    lines += ["## Ejemplo: precio x cantidad", "",
              "Precio 45 / 50 / 56 con probabilidades 25% / 50% / 25%; cantidad 8,000 / 10,000 / 12,000 con "
              "30% / 50% / 20%; costo variable S/ 32; costo fijo S/ 150,000. Utilidad esperada exacta "
              "S/ 28,850 y probabilidad de pérdida exacta 35%. El libro de Excel trae el modelo armado.", ""]
    lines += [s for s in EXAMPLE_STEPS] + [""]
    lines += ["## Buenas prácticas y errores comunes", "", md_table(["Tema", "Qué hacer"], PRACTICES), ""]
    lines += ["## Autopruebas", ""] + [t + "\n" for t in TESTS]
    lines += ["## XLRisk frente a @RISK", "", md_table(["Capacidad", "@RISK", f"XLRisk {VERSION}"], COMPARE), "",
              "@RISK es marca comercial de Lumivero. XLRisk es un proyecto independiente.", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    build_workbook(HERE / "XLRisk_Manual_ES.xlsx")
    build_markdown(HERE / "manual-es.md")
    print("ok")
