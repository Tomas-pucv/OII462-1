"""Genera la presentación de la Tarea 7 (≤ 5 min) a partir de figuras_t7/ (figuras y resultados.json).

Uso (desde la raíz del repositorio):
    python scripts/make_diagramas_t7.py
    python scripts/make_presentation_t7.py
    soffice --headless --convert-to pdf Tarea7_WTAP_presentacion.pptx
"""
import json
import os

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

GRUPO = "[Nombre(s) del grupo]"           # reemplazar por los nombres del grupo y volver a generar

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG = os.environ.get("T7_FIG", os.path.join(ROOT, "figuras_t7"))
OUT = os.environ.get("T7_OUT_PPTX", os.path.join(ROOT, "Tarea7_WTAP_presentacion.pptx"))
R = json.load(open(os.path.join(FIG, "resultados.json"), encoding="utf-8"))
TMP = os.path.join(FIG, "_tmp")
os.makedirs(TMP, exist_ok=True)

FONT = "Arial"
NAVY = RGBColor(0x1B, 0x2A, 0x41)
INK = RGBColor(0x2E, 0x35, 0x40)
MUTED = RGBColor(0x6B, 0x75, 0x85)
LIGHT = RGBColor(0xF3, 0xF5, 0xF8)
LINE = RGBColor(0xD5, 0xDA, 0xE1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLUE = RGBColor(0x1F, 0x77, 0xB4)
ORANGE = RGBColor(0xE8, 0x71, 0x0A)
GREEN = RGBColor(0x23, 0x89, 0x3A)
PURPLE = RGBColor(0x7B, 0x3F, 0xB0)
RED = RGBColor(0xC0, 0x39, 0x2B)
CHG = RGBColor(0xD3, 0x54, 0x00)
MODELS = ["T6 · Base", "EGAT · 1 capa", "EGAT · 2 capas", "EGAT · 3 capas"]

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
SW, SH = 13.333, 7.5
BLANK = prs.slide_layouts[6]
SLIDES = []          # (slide, notes)
TOTAL = 7


# ---------------------------------------------------------------- helpers (base: presentación del proyecto)
def box(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None, line_w=1.0):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(line_w)
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s


def text(slide, x, y, w, h, runs, size=14, color=INK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.1, shape=None):
    """runs: str, o lista de párrafos; cada párrafo es str o lista de (texto, {opciones})."""
    if shape is None:
        shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    paragraphs = runs if isinstance(runs, list) else [runs]
    for k, par in enumerate(paragraphs):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        pieces = par if isinstance(par, list) else [(par, {})]
        for txt, opt in pieces:
            r = p.add_run()
            r.text = txt
            f = r.font
            f.name = opt.get("font", FONT)
            f.size = Pt(opt.get("size", size))
            f.bold = opt.get("bold", bold)
            f.italic = opt.get("italic", False)
            f.color.rgb = opt.get("color", color)
        if isinstance(par, list) and par and "space_after" in par[0][1]:
            p.space_after = Pt(par[0][1]["space_after"])
    return shape


def image(slide, path, x, y, w=None, h=None, align="center"):
    iw, ih = Image.open(path).size
    ratio = iw / ih
    if w is not None and h is not None:
        if w / h > ratio:
            nw, nh = h * ratio, h
        else:
            nw, nh = w, w / ratio
        ox = x + {"center": (w - nw) / 2, "left": 0, "right": w - nw}[align]
        oy = y + (h - nh) / 2
    elif w is not None:
        nw, nh, ox, oy = w, w / ratio, x, y
    else:
        nw, nh, ox, oy = h * ratio, h, x, y
    return slide.shapes.add_picture(path, Inches(ox), Inches(oy), Inches(nw), Inches(nh))


def chip(slide, x, y, w, h, label, fill, size=12, color=WHITE, bold=True):
    s = box(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
    text(slide, 0, 0, 0, 0, label, size=size, bold=bold, color=color, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=s)
    return s


def new_slide(label, col, title, subtitle=None, notes=""):
    s = prs.slides.add_slide(BLANK)
    n = len(prs.slides)
    box(s, 0, 0, SW, SH, fill=WHITE)
    box(s, 0.5, 0.42, 0.09, 0.62, fill=col)
    text(s, 0.72, 0.3, 10.2, 0.6, title, size=25, bold=True, color=NAVY)
    if subtitle:
        text(s, 0.74, 0.86, 12.0, 0.4, subtitle, size=13, color=MUTED)
    chip(s, SW - 2.75, 0.4, 2.25, 0.36, label, col, size=10.5)
    text(s, 0.5, 7.07, 9, 0.3, "OII462 · Tarea 7 · El WTAP como grafo: encoder EGAT y profundidad", size=10,
         color=MUTED)
    text(s, SW - 1.4, 7.07, 0.9, 0.3, f"{n} / {TOTAL}", size=10, color=MUTED, align=PP_ALIGN.RIGHT)
    SLIDES.append((s, notes))
    return s


def card(slide, x, y, w, h, title, body, accent=BLUE, title_size=14, body_size=12, fill=LIGHT):
    box(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    box(slide, x, y + 0.15, 0.07, h - 0.3, fill=accent)
    text(slide, x + 0.22, y + 0.1, w - 0.35, 0.4, title, size=title_size, bold=True, color=accent)
    text(slide, x + 0.22, y + 0.5, w - 0.35, h - 0.58, body, size=body_size, color=INK, spacing=1.08)


def kpi(slide, x, y, w, h, value, label, color, sub=None, vsize=24):
    box(slide, x, y, w, h, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    box(slide, x, y, w, 0.08, fill=color)
    text(slide, x, y + 0.12, w, 0.55, value, size=vsize, bold=True, color=color, align=PP_ALIGN.CENTER)
    text(slide, x + 0.05, y + 0.64, w - 0.1, 0.35, label, size=11, bold=True, color=INK, align=PP_ALIGN.CENTER)
    if sub:
        text(slide, x + 0.05, y + 0.92, w - 0.1, 0.3, sub, size=9.5, color=MUTED, align=PP_ALIGN.CENTER)


def table(slide, x, y, w, rows, widths, row_h=0.36, size=10.5, header_fill=NAVY, fills=None, aligns=None, bold_cells=()):
    t = slide.shapes.add_table(len(rows), len(widths), Inches(x), Inches(y), Inches(w), Inches(row_h * len(rows))).table
    for c_, wd in enumerate(widths):
        t.columns[c_].width = Inches(wd)
    for r_, row in enumerate(rows):
        t.rows[r_].height = Inches(row_h)
        for c_, val in enumerate(row):
            cell = t.cell(r_, c_)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.fill.solid()
            if r_ == 0:
                cell.fill.fore_color.rgb = header_fill
            elif fills and fills[r_ - 1] is not None:
                cell.fill.fore_color.rgb = fills[r_ - 1]
            else:
                cell.fill.fore_color.rgb = LIGHT if r_ % 2 else WHITE
            al = aligns[c_] if aligns else PP_ALIGN.LEFT
            text(slide, 0, 0, 0, 0, val, size=size, bold=(r_ == 0 or (r_, c_) in bold_cells),
                 color=WHITE if r_ == 0 else INK, align=al, anchor=MSO_ANCHOR.MIDDLE, shape=cell, spacing=1.0)
    return t


def formula_png(latex, name, fontsize=24, color="#1B2A41"):
    path = os.path.join(TMP, f"{name}.png")
    fig = plt.figure(figsize=(8, 1))
    fig.text(0, 0.5, latex, fontsize=fontsize, color=color, va="center")
    fig.savefig(path, dpi=250, bbox_inches="tight", pad_inches=0.05, transparent=True)
    plt.close(fig)
    return path


def crop(path, name, box_frac):
    """Recorta una figura (fracciones x0, y0, x1, y1) y la guarda en _tmp."""
    im = Image.open(path)
    w, h = im.size
    x0, y0, x1, y1 = box_frac
    out = os.path.join(TMP, name)
    im.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h))).save(out)
    return out


def num(v, d=2):
    if v is None:
        return "—"
    return f"{v:.{d}f}".replace(".", ",")


def pct(v, d=2):
    return num(v, d) + "%"


def pm(m, s, d=2):
    return f"{num(m, d)} ± {num(s, d)}"


def F(name):
    return os.path.join(FIG, name)


EV = R["eval"]
SIZES = list(EV)
CONV = R["convergence"]
HIST = R["histories"]
NSEED = R["n_seeds"]
REF2 = R["ref_modulo2"]
REF6 = R["ref_t6"]
ABL = R["ablacion"]
GAT = R["gat_check"]
CAMPO = R["campo_receptivo"]
T6, E1, E2, E3 = MODELS
EGATS = MODELS[1:]
SHORT = {T6: "T6 · Base", E1: "EGAT-1", E2: "EGAT-2", E3: "EGAT-3"}
t10 = "10×8"


def gm(label, name):
    return EV[label][name]["mean"]


def rngn(vals, d=1):
    lo, hi = num(min(vals), d), num(max(vals), d)
    return lo if lo == hi else f"{lo}–{hi}"


def rng_(v, marca=False):
    ok = [x for x in v if x is not None]
    r = "–".join(str(x) for x in sorted({min(ok), max(ok)})) if ok else "—"
    return r + (f" ({len(ok)}/{len(v)})" if marca and ok and len(ok) < len(v) else "")


def avg_sizes(name):
    return float(np.mean([gm(l, name) for l in SIZES]))


SEC = {n: CONV[n]["sec_per_epoch"] for n in MODELS}
T6_AVG = {k: float(np.mean([REF6[l][k][0] for l in REF6])) for k in ("Base", "A", "B")}
WINS = {n: [l for l in SIZES if min(MODELS, key=lambda m: gm(l, m)) == n] for n in MODELS}
E1_WORST = sum(max(MODELS, key=lambda m: gm(l, m)) == E1 for l in SIZES)
BEATS_MMR_100 = gm("100×80", E3) < gm("100×80", "MMR")
SEC_CAPA = (SEC[E3] - SEC[E1]) / 2


def lista(xs):
    xs = list(xs)
    return xs[0] if len(xs) == 1 else ", ".join(xs[:-1]) + " y " + xs[-1]

# ================================================================ 1. portada
s = prs.slides.add_slide(BLANK)
box(s, 0, 0, SW, SH, fill=WHITE)
box(s, 0, 0, 5.0, SH, fill=NAVY)
text(s, 0.55, 0.7, 4.2, 0.4, "OII462 · TAREA 7", size=13, bold=True, color=RGBColor(0x8F, 0xB8, 0xDE))
text(s, 0.55, 1.15, 4.3, 2.4, "El WTAP como grafo: encoder EGAT y profundidad", size=28, bold=True, color=WHITE,
     spacing=1.0)
text(s, 0.55, 3.55, 4.1, 1.5, "Representamos el problema como un grafo bipartito, cambiamos el encoder por capas EGAT y "
     "medimos cuántas capas (saltos) necesita", size=14, color=RGBColor(0xD6, 0xE2, 0xEE), spacing=1.15)
box(s, 0.55, 5.25, 0.9, 0.05, fill=ORANGE)
text(s, 0.55, 5.45, 4.1, 0.4, GRUPO, size=16, bold=True, color=WHITE)
text(s, 0.55, 5.85, 4.1, 0.4, "Optimización combinatoria · 2026", size=12, color=RGBColor(0xB5, 0xC4, 0xD4))
text(s, 5.6, 0.65, 7.2, 0.5, "En una diapositiva", size=22, bold=True, color=NAVY)
items = [("1", "El problema como grafo", "bipartito armas ↔ objetivos; cada arista lleva la probabilidad de impacto $p_{ij}$"
          .replace("$p_{ij}$", "p_ij"), BLUE),
         ("2", "Encoder EGAT", "atención aditiva que usa la arista; el decoder de la Tarea 6 queda igual", GREEN),
         ("3", "Ablación de profundidad", "1, 2 y 3 capas = 1, 2 y 3 saltos; el grafo tiene diámetro 2", PURPLE),
         ("4", "Convergencia", f"EGAT-3 llega a gap < 3 % en {rng_(CONV[E3]['epoch_gap_below_3'])} épocas "
                            f"(T6: {rng_(CONV[T6]['epoch_gap_below_3'])}); misma configuración, {NSEED} semillas", ORANGE),
         ("5", "Resultados", f"prueba 10×8: T6 {pct(gm(t10, T6))} · EGAT-1 {pct(gm(t10, E1))} · EGAT-2 {pct(gm(t10, E2))}"
                             f" · EGAT-3 {pct(gm(t10, E3))}", RED)]
for k, (n, t, d, col) in enumerate(items):
    y = 1.4 + k * 1.05
    c_ = box(s, 5.6, y, 0.7, 0.7, fill=col, shape=MSO_SHAPE.OVAL)
    text(s, 0, 0, 0, 0, n, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=c_)
    text(s, 6.5, y - 0.02, 6.3, 0.45, t, size=17, bold=True, color=NAVY)
    text(s, 6.5, y + 0.37, 6.5, 0.45, d, size=12, color=MUTED)
SLIDES.append((s, "(≈10 s) Presentamos el WTAP como grafo, el encoder EGAT que reemplaza al de la Tarea 6 y cuánto "
                  "importa su profundidad."))

# ================================================================ 2. el grafo
s = new_slide("GRAFO", BLUE, "El WTAP como grafo bipartito G = (V, E)",
              "Dos tipos de nodos; toda la información que los relaciona está en las aristas",
              notes="(≈45 s) El WTAP no trae un grafo, pero es naturalmente bipartito. Los nodos son las armas y los "
                    "objetivos. Las armas no tienen datos propios: todas parten del mismo vector aprendido. Los objetivos "
                    "llevan su valor. Hay una arista entre cada arma y cada objetivo que puede atacar, con la probabilidad "
                    "de impacto como atributo; en nuestras instancias el grafo es completo. Es no dirigido, pero la "
                    "información viaja en los dos sentidos. Lo importante para la ablación: desde un arma, los objetivos "
                    "están a un salto y las demás armas a dos. El estado, es decir qué armas quedan libres y cuánto "
                    "sobrevive cada objetivo, entra en el decoder.")
image(s, F("grafo_wtap.png"), 0.4, 1.35, w=8.5, h=3.45)
rows = [["Elemento", "Representa", "Atributos"],
        ["Nodo arma (W)", "un arma por asignar", "ninguno propio: vector aprendido común"],
        ["Nodo objetivo (T)", "un objetivo enemigo", "valor V_j / max V"],
        ["Arista (W·T)", "el arma puede atacarlo", "probabilidad de impacto p_ij"]]
table(s, 0.45, 4.95, 8.4, rows, [2.2, 2.6, 3.6], row_h=0.44, size=11.5,
      fills=[RGBColor(0xDC, 0xEA, 0xF7), RGBColor(0xFD, 0xEB, 0xD9), LIGHT])
card(s, 9.15, 1.4, 3.8, 2.6, "Por qué bipartito",
     [[("No depende de cuántas armas u objetivos haya ni de su orden: el mismo modelo sirve para cualquier tamaño "
        "(como el MKP del material).", {})]], accent=BLUE, title_size=13.5, body_size=11.5)
card(s, 9.15, 4.2, 3.8, 2.7, "Saltos (hops)",
     [[("Desde un arma: los objetivos están a ", {}), ("1 salto", {"bold": True}), (" y las demás armas a ", {}),
       ("2 saltos", {"bold": True}), (f". El diámetro es {R['grafo']['diametro']}: una 3.ª capa ya no alcanza nodos "
                                       "nuevos.", {})]], accent=PURPLE, title_size=13.5, body_size=11.5)

# ================================================================ 3. arquitectura
s = new_slide("ARQUITECTURA", GREEN, "Encoder EGAT con el decoder de la Tarea 6",
              "Cambiamos solo el encoder; así la comparación aísla su efecto",
              notes="(≈50 s) Elegimos EGAT y no GAT: las armas no tienen datos propios, así que una GAT las deja "
                    f"idénticas; lo comprobamos: la distancia entre armas es {num(GAT['gat']['armas'], 0)} con GAT y "
                    f"{num(GAT['egat']['armas'], 1)} con EGAT. En cada capa, cada arma escucha a todos los objetivos y "
                    "cada objetivo a todas las armas. El puntaje de atención suma tres miradas: la del nodo, la del vecino "
                    "y la de la arista, en vez del producto punto de la Tarea 6, y el mensaje también lleva la arista. "
                    "Con L capas, la información viaja L saltos. El decoder es el de la Base de la Tarea 6, que fue "
                    "nuestro mejor modelo, sin ningún cambio.")
image(s, F("arq_egat.png"), 0.35, 1.35, w=8.75, h=5.55)
box(s, 9.3, 1.4, 3.6, 5.5, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
text(s, 9.45, 1.48, 3.4, 0.4, "Cambios respecto de la Tarea 6", size=14, bold=True, color=CHG)
chg = [("1", "Grafo", "bipartito explícito, con aristas p_ij y auto-bucles"),
       ("2", "Puntaje", "atención aditiva: nodo + vecino + arista (antes producto punto)"),
       ("3", "Transformaciones", "una matriz por tipo de nodo (antes consulta, clave y valor)"),
       ("4", "Arista", "un vector por arista, en el puntaje y en el mensaje"),
       ("5", "Profundidad", "1, 2 o 3 capas (antes 2)")]
for k, (n, t, d) in enumerate(chg):
    y = 1.98 + k * 0.82
    c_ = box(s, 9.48, y + 0.04, 0.34, 0.34, fill=CHG, shape=MSO_SHAPE.OVAL)
    text(s, 0, 0, 0, 0, n, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=c_)
    text(s, 9.9, y - 0.02, 2.95, 0.8, [[(t + ": ", {"bold": True, "color": NAVY}), (d, {})]], size=10.5, spacing=1.0)
text(s, 9.45, 6.12, 3.4, 0.75, [[("¿Por qué EGAT? ", {"bold": True, "color": GREEN}),
                                  ("con GAT las armas quedan idénticas: no tienen datos propios.", {})]], size=10.5,
     spacing=1.0)

# ================================================================ 4. ablación de profundidad
s = new_slide("ABLACIÓN", PURPLE, "Ablación de profundidad: ¿cuántos saltos?",
              f"EGAT con 1, 2 y 3 capas; puntos = semillas; línea azul = Base de la Tarea 6",
              notes="(≈45 s) Cada capa EGAT lleva la información un salto más lejos. Con una capa, un arma solo conoce los "
                    "objetivos y sus propias probabilidades, no a las armas con las que compite, y el gap sube a "
                    f"{num(gm(t10, E1), 1)} por ciento en 10 por 8. Con dos capas ya ve a las armas rivales y cubre todo "
                    "el grafo, porque el diámetro es 2: queda igual que la Base de la Tarea 6, que también tiene dos "
                    "capas. Con tres capas no aparecen nodos nuevos, pero sí información nueva: cada arma sabe qué otras "
                    f"opciones tienen sus rivales. Es la mejor: {num(gm(t10, E3), 1)} por ciento en 10 por 8. La "
                    "limitación es que tiene más parámetros que la Base.")
image(s, F("profundidad.png"), 0.45, 1.3, w=12.4, h=3.6)
rows = [["Capas", "Un arma conoce…"],
        ["1", "los objetivos que puede atacar"],
        ["2", "+ las armas que compiten por ellos"],
        ["3", "nada nuevo (diámetro 2): más combinación"]]
table(s, 0.45, 5.05, 5.6, rows, [0.9, 4.7], row_h=0.43, size=11.5,
      fills=[RGBColor(0xEE, 0xF8, 0xEA), RGBColor(0xDD, 0xF1, 0xD6), RGBColor(0xC9, 0xE8, 0xBF)])
text(s, 6.35, 5.0, 6.55, 1.95,
     [[("1 capa: ", {"bold": True, "color": PURPLE, "space_after": 4}),
       (f"no ve la competencia entre armas → {pct(gm(t10, E1))} en 10×8.", {})],
      [("2 capas: ", {"bold": True, "color": PURPLE, "space_after": 4}),
       (f"cubre el grafo (diámetro 2) → {pct(gm(t10, E2))}, como la T6 ({pct(gm(t10, T6))}), que también tiene 2.", {})],
      [("3 capas: ", {"bold": True, "color": PURPLE, "space_after": 4}),
       (f"ningún nodo nuevo, pero ve las opciones de las armas rivales → {pct(gm(t10, E3))}.", {})],
      [("Costo: ", {"bold": True, "color": NAVY}),
       (f"≈ {num(SEC_CAPA, 1)} s más por época por capa; EGAT-3 tiene más parámetros que la T6.", {})]],
     size=12)

# ================================================================ 5. convergencia
s = new_slide("RESULTADOS", RED, "Convergencia: misma configuración, 4 modelos",
              f"Media de {NSEED} semillas (banda = mínimo–máximo); 6 000 instancias 10×8, Adam + ReduceLROnPlateau",
              notes="(≈35 s) Los cuatro modelos parten igual. Más capas EGAT significa aprender más rápido: EGAT-3 "
                    f"baja de 3 por ciento de gap en {rng_(CONV[E3]['epoch_gap_below_3'])} épocas, contra "
                    f"{rng_(CONV[T6]['epoch_gap_below_3'])} de la Base; EGAT-1 casi nunca lo logra. EGAT-3 también "
                    "termina con la menor pérdida y es la más estable entre semillas. Cada época le cuesta "
                    f"{num(100 * (SEC[E3] / SEC[T6] - 1), 0)} por ciento más, pero como converge antes, en minutos "
                    "igual es la más rápida.")
image(s, F("convergencia.png"), 0.35, 1.35, w=8.3, h=5.6)
rows = [["", "Gap val. época 0", "Épocas a gap<3 %", "Mejor loss val.", "s / época"]]
for n in MODELS:
    cv = CONV[n]
    rows.append([SHORT[n], pct(cv["gap_init"], 0), rng_(cv["epoch_gap_below_3"], marca=True),
                 pm(cv["best_val_loss_mean"], cv["best_val_loss_std"], 3), num(cv["sec_per_epoch"], 1)])
table(s, 8.85, 1.45, 4.1, rows, [0.95, 0.68, 0.72, 1.1, 0.65], row_h=0.44, size=9,
      fills=[RGBColor(0xDC, 0xEA, 0xF7), RGBColor(0xEE, 0xF8, 0xEA), RGBColor(0xDD, 0xF1, 0xD6),
             RGBColor(0xC9, 0xE8, 0xBF)], aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 4)
text(s, 8.85, 3.85, 4.1, 3.05,
     [[("Velocidad: ", {"bold": True, "color": GREEN, "space_after": 5}),
       (f"gap < 3 % en {rng_(CONV[E3]['epoch_gap_below_3'])} épocas (EGAT-3), "
        f"{rng_(CONV[E2]['epoch_gap_below_3'])} (EGAT-2) y {rng_(CONV[T6]['epoch_gap_below_3'])} (T6); EGAT-1 casi "
        "nunca.", {})],
      [("Nivel final: ", {"bold": True, "color": GREEN, "space_after": 5}),
       (f"pérdida {num(CONV[E3]['best_val_loss_mean'], 3)} (EGAT-3) vs {num(CONV[T6]['best_val_loss_mean'], 3)} "
        "(T6).", {})],
      [("Costo: ", {"bold": True, "color": RED, "space_after": 5}),
       (f"{num(SEC[E3], 1)} s por época vs {num(SEC[T6], 1)} s, pero llega antes a gap < 3 %: "
        f"{num(CONV[E3]['minutes_to_gap3'], 1)} vs {num(CONV[T6]['minutes_to_gap3'], 1)} min.", {})],
      [("Estabilidad: ", {"bold": True, "color": NAVY}),
       (f"{len(MODELS) * NSEED} corridas sin divergencias; EGAT-3 es la más estable.", {})]],
     size=11.5)

# ================================================================ 6. evaluación por tamaño
s = new_slide("RESULTADOS", RED, "Evaluación por tamaño: gap vs el experto ILS",
              f"Entrenados solo con 10×8; mismas instancias de la Tarea 6 (30 por tamaño, 50 en 10×8) × {NSEED} semillas",
              notes="(≈45 s) Los evaluamos sin reentrenar en siete tamaños, con las tres semillas de cada modelo. En 10 "
                    f"por 8, EGAT con tres capas logra {num(gm(t10, E3), 1)} por ciento, contra "
                    f"{num(gm(t10, T6), 1)} de la Base de la Tarea 6, y encuentra la solución del experto en casi la "
                    f"mitad de las ejecuciones. Es la mejor en {len(WINS[E3])} de los siete tamaños, con "
                    f"{num(avg_sizes(E3), 1)} por ciento de promedio" +
                    (", y en 100 por 80 es el primero de nuestros modelos que supera a MMR. " if BEATS_MMR_100 else ". ") +
                    f"Pero en {lista(WINS[T6]).replace('×', ' por ')} la Base sigue siendo mejor, y ahí EGAT-3 varía "
                    f"más entre semillas. EGAT con una capa es la peor en {E1_WORST} de los siete tamaños.")
rows = [["Tamaño"] + [SHORT[n] for n in MODELS] + ["MMR"]]
bold = []
for r_, l in enumerate(SIZES, start=1):
    best = min(MODELS, key=lambda n: gm(l, n))
    row = [l]
    for c_, n in enumerate(MODELS, start=1):
        row.append(pm(EV[l][n]["mean"], EV[l][n]["std"]))
        if n == best:
            bold.append((r_, c_))
    rows.append(row + [pm(EV[l]["MMR"]["mean"], EV[l]["MMR"]["std"])])
table(s, 0.45, 1.45, 7.6, rows, [1.0, 1.32, 1.32, 1.32, 1.32, 1.32], row_h=0.42, size=10.5,
      aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 5, bold_cells=bold)
text(s, 0.45, 1.5 + 0.42 * len(rows), 7.6, 0.35, f"Gap (%) media ± desv. estándar sobre {NSEED} semillas × instancias; en "
     "negrita, el mejor de los cuatro modelos.", size=10, color=MUTED)
cx = crop(F("comparacion.png"), "comparacion_der.png", (0.445, 0.0, 1.0, 1.0))
image(s, cx, 8.2, 1.4, w=4.8, h=3.6)
text(s, 8.25, 5.1, 4.75, 1.85,
     [[("• ", {"space_after": 3}), (f"EGAT-3: la mejor en {len(WINS[E3])} de {len(SIZES)} tamaños; promedio "
                                   f"{pct(avg_sizes(E3))} vs {pct(avg_sizes(T6))} de la T6.", {})],
      [("• ", {"space_after": 3}), (f"100×80: por primera vez bajo MMR ({pct(gm('100×80', E3))} vs "
                                   f"{pct(gm('100×80', 'MMR'))}).", {})],
      [("• ", {"space_after": 3}), (f"{lista(WINS[T6])}: la T6 sigue siendo mejor.", {})],
      [("• ", {}), (f"EGAT-1: la peor en {E1_WORST} de {len(SIZES)} tamaños.", {})]],
     size=11)

# ================================================================ 7. conclusiones
s = new_slide("CONCLUSIONES", NAVY, "Conclusiones", None,
              notes="(≈20 s) En resumen: el WTAP es naturalmente un grafo bipartito, y la EGAT es la forma de usarlo, "
                    "porque las armas solo se distinguen por sus aristas. Lo que más pesa es la profundidad: con tres "
                    "capas obtenemos nuestro mejor modelo, que recomendamos como nueva base.")
concl = [("El WTAP como grafo",
          "Bipartito armas–objetivos con p_ij en las aristas: no depende del tamaño ni del orden, y sus aristas son las "
          "acciones del decoder.", BLUE),
         ("EGAT en vez de producto punto",
          "Sin datos propios, las armas solo se distinguen por sus aristas: una GAT las deja iguales. A igual profundidad "
          "(2 capas), la EGAT rinde como la T6 y aprende más rápido.", GREEN),
         ("Profundidad",
          f"1 salto no basta ({pct(gm(t10, E1))}); 2 saltos igualan a la T6 ({pct(gm(t10, E2))}); 3 capas dan nuestro "
          f"mejor modelo ({pct(gm(t10, E3))} en 10×8, {pct(avg_sizes(E3))} en {len(SIZES)} tamaños).", PURPLE),
         ("Recomendación",
          f"EGAT de 3 capas como nuevo modelo base. Pendiente: {lista(WINS[T6])}, una T6 de 3 capas y probar 4 capas.",
          RED)]
for k, (t, d, col) in enumerate(concl):
    card(s, 0.6 + (k % 2) * 6.15, 1.35 + (k // 2) * 2.75, 5.95, 2.55, t, d, accent=col, title_size=16, body_size=14.5)

for sl, notes in SLIDES:
    sl.notes_slide.notes_text_frame.text = notes
prs.save(OUT)
print("Presentación guardada en", OUT, "·", len(prs.slides), "diapositivas")
