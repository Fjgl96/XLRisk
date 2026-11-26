"""Draws the cell illustrations of the Spanish manual (docs/img/*.png).

Each picture looks like a piece of an Excel sheet: name box and formula bar,
column letters, row numbers, the active cell framed in green, and numbered
red callouts explained under the grid, in the style of the classic books that
taught Excel.

Usage:  python docs/illustrations.py      (needs Pillow)
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent
OUT = HERE / "img"
FONT_DIRS = ["/usr/share/fonts/truetype/crosextra", "/usr/share/fonts/truetype/liberation",
             "C:/Windows/Fonts"]
SCALE = 2  # draw at 2x for sharp text

INK = (32, 32, 32)
GRID = (208, 208, 208)
HEADER_BG = (217, 217, 217)
HEADER_INK = (64, 64, 64)
ACTIVE = (16, 124, 65)       # Excel green
INPUT_BG = (255, 255, 0)
INPUT_INK = (0, 0, 255)
OUTPUT_BG = (0, 140, 0)
CALLOUT = (192, 0, 0)
NAVY = (31, 56, 100)
MUTED = (96, 96, 96)


def font(size, bold=False, italic=False):
    names = []
    if bold:
        names += ["Carlito-Bold.ttf", "LiberationSans-Bold.ttf", "calibrib.ttf"]
    elif italic:
        names += ["Carlito-Italic.ttf", "LiberationSans-Italic.ttf", "calibrii.ttf"]
    names += ["Carlito-Regular.ttf", "LiberationSans-Regular.ttf", "calibri.ttf"]
    for d in FONT_DIRS:
        for n in names:
            p = Path(d) / n
            if p.exists():
                return ImageFont.truetype(str(p), size * SCALE)
    return ImageFont.load_default()


F_CELL = font(13)
F_BOLD = font(13, bold=True)
F_ITAL = font(12, italic=True)
F_HDR = font(12)
F_BAR = font(13)
F_FX = font(13, italic=True)
F_NOTE = font(13)
F_NUM = font(11, bold=True)

ROW_H = 24
HDR_W = 34
BAR_H = 30
PAD = 14


def col_letter(i):
    return "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[i]


def parse(ref):
    col = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".index(ref[0])
    return col, int(ref[1:])


def is_number(text):
    t = text.replace(",", "").replace("%", "").replace("S/", "").replace("US$", "").strip()
    try:
        float(t)
        return True
    except ValueError:
        return t in ("TRUE", "FALSE", "VERDADERO", "FALSO", "#N/A")


def render(spec):
    widths = spec["cols"]
    nrows = spec["rows"]
    notes = spec.get("notes", [])
    s = SCALE
    grid_w = HDR_W + sum(widths)
    width = PAD * 2 + grid_w
    note_h = 26 * len(notes) + (12 if notes else 0)
    height = PAD * 2 + BAR_H + 8 + ROW_H * (nrows + 1) + note_h
    img = Image.new("RGB", (width * s, height * s), "white")
    d = ImageDraw.Draw(img)

    def box(x0, y0, x1, y1, **kw):
        d.rectangle([x0 * s, y0 * s, x1 * s, y1 * s], **kw)

    def text(x, y, t, f, fill=INK, anchor="lm"):
        d.text((x * s, y * s), t, font=f, fill=fill, anchor=anchor)

    # Name box and formula bar
    x0, y0 = PAD, PAD
    box(x0, y0, x0 + 70, y0 + BAR_H, outline=GRID, width=s)
    text(x0 + 8, y0 + BAR_H / 2, spec["active"], F_BAR)
    text(x0 + 84, y0 + BAR_H / 2, "fx", F_FX, fill=MUTED)
    box(x0 + 104, y0, x0 + grid_w, y0 + BAR_H, outline=GRID, width=s)
    text(x0 + 112, y0 + BAR_H / 2, spec["formula"], F_BAR)

    # Column letters and row numbers
    gy = y0 + BAR_H + 8
    box(x0, gy, x0 + grid_w, gy + ROW_H, fill=HEADER_BG)
    xs = [x0 + HDR_W]
    for w in widths:
        xs.append(xs[-1] + w)
    for i in range(len(widths)):
        text((xs[i] + xs[i + 1]) / 2, gy + ROW_H / 2, col_letter(i), F_HDR, fill=HEADER_INK, anchor="mm")
    for r in range(1, nrows + 1):
        y = gy + ROW_H * r
        box(x0, y, x0 + HDR_W, y + ROW_H, fill=HEADER_BG)
        text(x0 + HDR_W / 2, y + ROW_H / 2, str(r), F_HDR, fill=HEADER_INK, anchor="mm")

    def cell_box(ref):
        c, r = parse(ref)
        return xs[c], gy + ROW_H * r, xs[c + 1], gy + ROW_H * (r + 1)

    # Grid lines
    for r in range(1, nrows + 2):
        y = gy + ROW_H * r
        d.line([(x0 + HDR_W) * s, y * s, (x0 + grid_w) * s, y * s], fill=GRID, width=s)
    for x in xs:
        d.line([x * s, (gy + ROW_H) * s, x * s, (gy + ROW_H * (nrows + 1)) * s], fill=GRID, width=s)

    # Cells
    for ref, (value, style) in spec["cells"].items():
        bx0, by0, bx1, by1 = cell_box(ref)
        f, fill, bg = F_CELL, INK, None
        if style == "bold":
            f = F_BOLD
        elif style == "input":
            bg, fill = INPUT_BG, INPUT_INK
        elif style == "output":
            bg, fill, f = OUTPUT_BG, (255, 255, 255), F_BOLD
        elif style == "note":
            f, fill = F_ITAL, MUTED
        elif style == "title":
            f, fill = F_BOLD, NAVY
        if bg:
            box(bx0 + 0.5, by0 + 0.5, bx1 - 0.5, by1 - 0.5, fill=bg)
        if is_number(value) and style not in ("note", "title"):
            text(bx1 - 6, (by0 + by1) / 2, value, f, fill=fill, anchor="rm")
        else:
            text(bx0 + 6, (by0 + by1) / 2, value, f, fill=fill)

    # Active cell
    ax0, ay0, ax1, ay1 = cell_box(spec["active"])
    d.rectangle([ax0 * s - s, ay0 * s - s, ax1 * s + s, ay1 * s + s], outline=ACTIVE, width=3 * s)
    d.rectangle([(ax1 - 3) * s, (ay1 - 3) * s, (ax1 + 3) * s, (ay1 + 3) * s], fill=ACTIVE)

    # Callouts: a numbered red circle at the top right corner of each cell
    def badge(cx, cy, n):
        rad = 9
        d.ellipse([(cx - rad) * s, (cy - rad) * s, (cx + rad) * s, (cy + rad) * s], fill=CALLOUT)
        text(cx, cy + 0.5, str(n), F_NUM, fill=(255, 255, 255), anchor="mm")

    ny = gy + ROW_H * (nrows + 1) + 18
    for i, (ref, note) in enumerate(notes, 1):
        if ref:
            bx0, by0, bx1, by1 = cell_box(ref)
            value, style = spec["cells"].get(ref, ("", ""))
            right_aligned = is_number(value) and style not in ("note", "title")
            # the badge goes on the empty side of the cell, so it never covers its value
            cx = bx0 + 12 if right_aligned else bx1 - 12
            badge(cx, (by0 + by1) / 2, i)
        badge(x0 + 9, ny + 26 * (i - 1), i)
        text(x0 + 26, ny + 26 * (i - 1), note, F_NOTE)

    OUT.mkdir(exist_ok=True)
    img.save(OUT / spec["file"], optimize=True)
    return OUT / spec["file"]


# ---------------------------------------------------------------------------
# The illustrations
# ---------------------------------------------------------------------------

SPECS = [
    dict(
        file="01_supuesto.png", key="supuesto", original=True,
        title="Convertir un supuesto en una distribución",
        cols=[200, 100, 260], rows=5, active="B2", formula="=RiskTriang(45, 50, 56)",
        cells={
            "A1": ("Supuesto", "bold"), "B1": ("Valor", "bold"),
            "A2": ("Precio de venta (S/)", ""), "B2": ("50.33", "input"),
            "A3": ("Cantidad (unidades)", ""), "B3": ("10,000", "input"),
            "A4": ("Costo variable (S/)", ""), "B4": ("32.00", "input"),
            "C2": ("antes decía 50", "note"),
        },
        notes=[
            ("B2", "En lugar del número fijo se escribe la función: mínimo 45, más probable 50, máximo 56."),
            (None, "Fuera de la simulación la celda muestra el valor esperado: (45 + 50 + 56) / 3 = 50.33."),
            (None, "Una función por celda. El resto del modelo sigue usando B2 como siempre."),
        ]),
    dict(
        file="02_escenarios.png", key="escenarios", original=True,
        title="Escenarios con su probabilidad: RiskDiscrete",
        cols=[150, 90, 90, 200], rows=6, active="B6", formula="=RiskDiscrete(B2:B4, C2:C4)",
        cells={
            "A1": ("Escenario", "bold"), "B1": ("Precio", "bold"), "C1": ("Prob.", "bold"),
            "A2": ("Bajo", ""), "B2": ("45", "input"), "C2": ("25%", "input"),
            "A3": ("Base", ""), "B3": ("50", "input"), "C3": ("50%", "input"),
            "A4": ("Alto", ""), "B4": ("56", "input"), "C4": ("25%", "input"),
            "C5": ("100%", "bold"),
            "A6": ("Precio simulado", "bold"), "B6": ("50.25", ""),
        },
        notes=[
            ("B2", "Valores en orden ascendente: de menor a mayor."),
            ("C5", "Las probabilidades tienen que sumar exactamente 100%."),
            ("B6", "Fuera de la simulación muestra el promedio ponderado: 45x25% + 50x50% + 56x25% = 50.25."),
        ]),
    dict(
        file="03_correlacion.png", key="correlacion", original=True,
        title="Hacer que dos supuestos se muevan juntos: RiskCorrmat",
        cols=[150, 90, 24, 90, 90, 90], rows=4, active="B3",
        formula="=RiskTriang(11500, 12000, 12800, RiskCorrmat($E$2:$F$3, 2))",
        cells={
            "A1": ("Supuesto", "bold"), "B1": ("Valor", "bold"),
            "A2": ("Volumen (t)", ""), "B2": ("1,200", ""),
            "A3": ("Precio (S/ por t)", ""), "B3": ("12,100", ""),
            "E1": ("Volumen", "bold"), "F1": ("Precio", "bold"),
            "D2": ("Volumen", "bold"), "E2": ("1.00", ""), "F2": ("-0.50", "input"),
            "D3": ("Precio", "bold"), "E3": ("-0.50", ""), "F3": ("1.00", ""),
        },
        notes=[
            ("F2", "Matriz de correlación: 1 en la diagonal y simétrica. -0.50: si el precio sube, el volumen tiende a bajar."),
            ("B2", "El volumen ocupa la posición 1: =RiskNormal(1200, 120, RiskCorrmat($E$2:$F$3, 1))"),
            ("B3", "El precio ocupa la posición 2: RiskCorrmat va siempre como último argumento."),
        ]),
    dict(
        file="04_output_boton.png", key="output_boton", original=True,
        title="Marcar el resultado con Add Output",
        cols=[180, 110, 250], rows=6, active="B6", formula="=(B2-B4)*B3-B5",
        cells={
            "A1": ("Modelo", "bold"),
            "A2": ("Precio (S/)", ""), "B2": ("50.25", ""), "C2": ("=RiskDiscrete(...)", "note"),
            "A3": ("Cantidad", ""), "B3": ("9,800", ""), "C3": ("=RiskDiscrete(...)", "note"),
            "A4": ("Costo variable (S/)", ""), "B4": ("32.00", "input"),
            "A5": ("Costo fijo (S/)", ""), "B5": ("150,000", "input"),
            "A6": ("Utilidad (S/)", "bold"), "B6": ("28,850", "output"),
        },
        notes=[
            ("B6", "Selecciona la celda resultado y pulsa XLRisk > Add Output."),
            (None, "Escribe un nombre, por ejemplo Utilidad. Queda registrado en la hoja XLRisk."),
            (None, "La fórmula no cambia: XLRisk solo anota este valor en cada iteración."),
        ]),
    dict(
        file="05_riskoutput.png", key="riskoutput", original=False,
        title="Marcar el resultado desde la fórmula: RiskOutput (versión 1.2)",
        cols=[180, 110, 250], rows=6, active="B6", formula='=RiskOutput("Utilidad") + (B2-B4)*B3-B5',
        cells={
            "A1": ("Modelo", "bold"),
            "A2": ("Precio (S/)", ""), "B2": ("50.25", ""),
            "A3": ("Cantidad", ""), "B3": ("9,800", ""),
            "A4": ("Costo variable (S/)", ""), "B4": ("32.00", "input"),
            "A5": ("Costo fijo (S/)", ""), "B5": ("150,000", "input"),
            "A6": ("Utilidad (S/)", "bold"), "B6": ("28,850", "output"),
        },
        notes=[
            ("B6", "RiskOutput(\"Utilidad\") suma 0, así que el valor no cambia. Es la misma sintaxis de @RISK."),
            (None, "Al correr, XLRisk detecta la celda solo: no hace falta pulsar Add Output."),
        ]),
    dict(
        file="06_estadisticas.png", key="estadisticas", original=False,
        title="Traer los resultados a las celdas (versión 1.2)",
        cols=[150, 90, 24, 170, 90, 190], rows=6, active="E3", formula="=RiskTarget(B6, 0)",
        cells={
            "A1": ("Modelo", "bold"), "A6": ("Utilidad (S/)", "bold"), "B6": ("28,850", "output"),
            "D1": ("Indicador", "bold"), "E1": ("Valor", "bold"), "F1": ("Fórmula", "bold"),
            "D2": ("Utilidad media", ""), "E2": ("28,925", ""), "F2": ("=RiskMean(B6)", "note"),
            "D3": ("Prob. de pérdida", ""), "E3": ("35.2%", ""), "F3": ("=RiskTarget(B6, 0)", "note"),
            "D4": ("Peor caso", ""), "E4": ("-46,000", ""), "F4": ("=RiskMin(B6)", "note"),
            "D5": ("Mejor caso", ""), "E5": ("138,000", ""), "F5": ("=RiskMax(B6)", "note"),
            "D6": ("Percentil 5", ""), "E6": ("-46,000", ""), "F6": ("=RiskPercentile(B6, 0.05)", "note"),
        },
        notes=[
            ("E3", "RiskTarget(B6, 0): probabilidad de que la utilidad sea 0 o menos."),
            (None, "Se actualizan solas al terminar cada simulación y siguen funcionando al reabrir el libro."),
            (None, "Antes de la primera simulación, o mientras corre, muestran #N/A."),
        ]),
    dict(
        file="07_hoja_xlrisk.png", key="hoja_xlrisk", original=True,
        title="La hoja XLRisk",
        cols=[210, 70, 16, 110, 200, 16, 110, 90], rows=6, active="B2", formula="2026",
        cells={
            "A1": ("Simulation Settings", "bold"),
            "A2": ("Seed", ""), "B2": ("2026", "input"),
            "A3": ("Update Screen", ""), "B3": ("FALSE", "input"),
            "A4": ("Iterations", ""), "B4": ("5000", "input"),
            "A5": ("Latin Hypercube", ""), "B5": ("TRUE", "input"),
            "A6": ("Calculate data tables", ""), "B6": ("TRUE", "input"),
            "D1": ("Simulation Inputs", "bold"), "D3": ("Range", "bold"), "E3": ("Formula", "bold"),
            "D4": ("Modelo!$B$2", ""), "E4": ("RiskDiscrete(...)", ""),
            "D5": ("Modelo!$B$3", ""), "E5": ("RiskDiscrete(...)", ""),
            "G1": ("Simulation Outputs", "bold"), "G3": ("Range", "bold"), "H3": ("Name", "bold"),
            "G4": ("Modelo!$B$6", ""), "H4": ("Utilidad", ""),
        },
        notes=[
            ("B2", "Semilla: con un número distinto de 0 la simulación se puede repetir exacta."),
            ("B4", "Iteraciones: 1,000 para explorar, 5,000 a 10,000 para presentar."),
            ("D4", "Inputs: XLRisk los encuentra solo, buscando las funciones Risk en las fórmulas."),
            ("G4", "Outputs: los registrados con Add Output."),
        ]),
    dict(
        file="08_resultados.png", key="resultados", original=True,
        title="Leer la hoja de resultados",
        cols=[200, 110, 24, 200, 110], rows=7, active="E3", formula="=PERCENTRANK.INC(...)",
        cells={
            "A1": ("Simulation Statistics", "title"), "B1": ("Utilidad", "bold"),
            "A2": ("Mean", ""), "B2": ("28,925", ""),
            "A3": ("Median", ""), "B3": ("30,000", ""),
            "A4": ("Std. Deviation", ""), "B4": ("46,273", ""),
            "A5": ("Minimum", ""), "B5": ("-46,000", ""),
            "A6": ("Maximum", ""), "B6": ("138,000", ""),
            "D1": ("Percent Rank", "title"),
            "D2": ("Enter Value:", "bold"), "E2": ("0", "input"),
            "D3": ("% rank:", "bold"), "E3": ("35.2%", ""),
        },
        notes=[
            ("B2", "Media: el resultado esperado del modelo."),
            ("B5", "Mínimo y máximo: el rango completo de lo que puede pasar."),
            ("E3", "En Enter Value escribe 0: % rank da la probabilidad de pérdida (aproximada)."),
        ]),
]


def build_all():
    return [render(sp) for sp in SPECS]


if __name__ == "__main__":
    for p in build_all():
        print(p.name)
