"""Genera la presentación de la Tarea 6 (≤ 5 min) a partir de figuras_t6/ (figuras y resultados.json).

Uso (desde la raíz del repositorio):
    python scripts/make_diagramas_t6.py
    python scripts/make_presentation_t6.py
    soffice --headless --convert-to pdf Tarea6_WTAP_presentacion.pptx
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
FIG = os.environ.get("T6_FIG", os.path.join(ROOT, "figuras_t6"))
OUT = os.environ.get("T6_OUT_PPTX", os.path.join(ROOT, "Tarea6_WTAP_presentacion.pptx"))
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
PALE = {"Base": RGBColor(0xDC, 0xEA, 0xF7), "A · sesgo MMR": RGBColor(0xEC, 0xE4, 0xF5),
        "B · ISAB": RGBColor(0xDF, 0xF0, 0xDF), "MMR": RGBColor(0xFD, 0xEB, 0xD9), "Módulo 2": RGBColor(0xE8, 0xEE, 0xF6)}
COL = {"Base": BLUE, "A · sesgo MMR": PURPLE, "B · ISAB": GREEN}
MODELS = ["Base", "A · sesgo MMR", "B · ISAB"]

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
SW, SH = 13.333, 7.5
BLANK = prs.slide_layouts[6]
SLIDES = []          # (slide, notes)
TOTAL = 8


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
    text(s, 0.5, 7.07, 9, 0.3, "OII462 · Tarea 6 · Encoder estático, complejidad y variantes para el WTAP", size=10,
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
CRIT = R["criterio_mmr"]
TIMING = R["timing"]
SL = R["slopes_step"]
SPD = R["speedup_vs_mod2"]
REF2 = R["ref_modulo2"]
DIAG = R["diagnostico_modulo2"]
NSEED = R["n_seeds"]


def gm(label, name):
    return EV[label][name]["mean"]


def step_ms(name, n):
    return 1e3 * TIMING[name]["dec"][TIMING[name]["n"].index(n)]


big = SIZES[-1]
# ================================================================ 1. portada
s = prs.slides.add_slide(BLANK)
box(s, 0, 0, SW, SH, fill=WHITE)
box(s, 0, 0, 5.0, SH, fill=NAVY)
text(s, 0.55, 0.7, 4.2, 0.4, "OII462 · TAREA 6", size=13, bold=True, color=RGBColor(0x8F, 0xB8, 0xDE))
text(s, 0.55, 1.15, 4.3, 2.4, "Encoder estático, complejidad y variantes del Transformer para el WTAP", size=28, bold=True,
     color=WHITE, spacing=1.0)
text(s, 0.55, 3.75, 4.1, 1.3, "Rediseñamos el modelo del módulo 2 para codificar una vez por instancia, con un decoder "
     "lineal, y lo comparamos con dos variantes", size=14, color=RGBColor(0xD6, 0xE2, 0xEE), spacing=1.15)
box(s, 0.55, 5.25, 0.9, 0.05, fill=ORANGE)
text(s, 0.55, 5.45, 4.1, 0.4, GRUPO, size=16, bold=True, color=WHITE)
text(s, 0.55, 5.85, 4.1, 0.4, "Optimización combinatoria · 2026", size=12, color=RGBColor(0xB5, 0xC4, 0xD4))
text(s, 5.6, 0.65, 7.2, 0.5, "En una diapositiva", size=22, bold=True, color=NAVY)
items = [("1", "Base rediseñada", "encoder bipartito armas ↔ objetivos (1 vez por instancia) + decoder con el estado", BLUE),
         ("2", "Complejidad", f"decoder lineal en n = W·T; en 50×40 una decisión pasa de {num(step_ms('Módulo 2', 2000), 0)} ms "
                              f"a {num(step_ms('Base', 2000), 1)} ms", ORANGE),
         ("3", "Variante A", "sesgo inductivo con la ganancia de MMR (salida + atención)", PURPLE),
         ("4", "Variante B", "decoder con ISAB: los pares interactúan en tiempo lineal", GREEN),
         ("5", "Resultados", f"prueba 10×8: Base {pct(gm('10×8', 'Base'))} · A {pct(gm('10×8', 'A · sesgo MMR'))} · "
                             f"B {pct(gm('10×8', 'B · ISAB'))} vs MMR {pct(gm('10×8', 'MMR'))}", RED)]
for k, (n, t, d, col) in enumerate(items):
    y = 1.4 + k * 1.05
    c_ = box(s, 5.6, y, 0.7, 0.7, fill=col, shape=MSO_SHAPE.OVAL)
    text(s, 0, 0, 0, 0, n, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=c_)
    text(s, 6.5, y - 0.02, 6.3, 0.45, t, size=17, bold=True, color=NAVY)
    text(s, 6.5, y + 0.37, 6.4, 0.45, d, size=12, color=MUTED)
SLIDES.append((s, "(≈10 s) Presentamos el rediseño del Transformer del módulo 2 para el WTAP: una base con el encoder "
                  "estático, su complejidad, dos variantes y la comparación experimental."))

# ================================================================ 2. base rediseñada
s = new_slide("ARQUITECTURA", BLUE, "Arquitectura base rediseñada",
              "El encoder solo ve V y p y corre 1 vez por instancia; el decoder integra el estado en cada paso",
              notes=f"(≈60 s) En el módulo 2, {DIAG['n_dinamicas']} de las 16 features dependían del estado, así que el "
                    "'encoder' se recalculaba en cada decisión. Ahora el encoder solo recibe la instancia: las armas no "
                    "tienen información propia, así que parten de un token aprendible; los objetivos, de su valor; y la "
                    "matriz p entra como aristas de una atención bipartita armas-objetivos, en el puntaje y en el mensaje. "
                    "Corre una vez. El decoder recibe el estado (armas libres y la probabilidad de supervivencia de cada "
                    "objetivo), arma un contexto, hace un glimpse y puntúa cada par con un puntero aditivo que también ve "
                    "p_ij. Los círculos naranjos son los cinco cambios respecto del módulo 2.")
image(s, F("arq_base.png"), 0.35, 1.35, w=8.75, h=5.55)
box(s, 9.3, 1.4, 3.6, 5.5, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
text(s, 9.45, 1.48, 3.4, 0.4, "Cambios respecto del módulo 2", size=14, bold=True, color=CHG)
chg = [("1", "Tokens", "W armas (token aprendible) + T objetivos; p_ij como aristas. Antes: 1+W·T filas × 16 features"),
       ("2", "Encoder", "atención bipartita con aristas, 4 cabezas, 1 vez por instancia. Antes: en cada paso"),
       ("3", "Estado", f"solo en el decoder (q_j, valor en juego, máscara). Antes: {DIAG['n_dinamicas']} features "
                      "dinámicas y heurísticas en la entrada"),
       ("4", "Contexto", "promedios de armas libres y objetivos + glimpse MHA. Antes: fila global"),
       ("5", "Puntero", "aditivo por pares con la arista p_ij y C·tanh. Antes: q·h_r/√d")]
for k, (n, t, d) in enumerate(chg):
    y = 1.98 + k * 0.97
    c_ = box(s, 9.48, y + 0.04, 0.34, 0.34, fill=CHG, shape=MSO_SHAPE.OVAL)
    text(s, 0, 0, 0, 0, n, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=c_)
    text(s, 9.9, y - 0.02, 2.95, 0.95, [[(t + ": ", {"bold": True, "color": NAVY}), (d, {})]], size=10.5, spacing=1.0)

# ================================================================ 3. complejidad
s = new_slide("COMPLEJIDAD", ORANGE, "Complejidad: encoder 1 vez, decoder lineal en n = W·T",
              "Tiempos medidos con PyTorch en 1 hilo (pesos aleatorios); pendientes log-log en los tamaños más grandes",
              notes=f"(≈35 s) Con n igual a W por T, el encoder nuevo es lineal en n (la atención es bipartita) y se paga una "
                    f"vez. El decoder también es lineal: es el mínimo, porque debe entregar un logit por par. El módulo 2 era "
                    f"cuadrático por paso: pendiente {num(SL['Módulo 2'])} contra {num(SL['Base'])} de la base. En 50 por "
                    f"40, una decisión baja de {num(step_ms('Módulo 2', 2000), 0)} a {num(step_ms('Base', 2000), 1)} "
                    f"milisegundos y la instancia completa es {num(SPD['50×40'], 0)} veces más rápida. La única pieza "
                    "cuadrática posible era la interacción entre pares de la variante B, y por eso usa ISAB (curva roja = "
                    "la versión con atención completa).")
cx = crop(F("complejidad.png"), "complejidad_paso.png", (0.0, 0.0, 0.345, 1.0))
image(s, cx, 0.4, 1.35, w=6.3, h=5.5)
rows = [["", "Encoder (1 vez)", "Decoder (por paso)"],
        ["Módulo 2", "— (todo en cada paso)", "O(W²T) + O(n²·d)"],
        ["Base", "O(n·d + (W+T)·d²)", "O(n·d + (W+T)·d²)"],
        ["A · sesgo MMR", "igual + O(n)", "igual + O(n)"],
        ["B · ISAB", "igual", "O(n·m·d + n·d²)"],
        ["B atención completa", "igual", "O(n²·d)"]]
table(s, 6.95, 1.45, 6.0, rows, [1.85, 2.05, 2.1], row_h=0.42, size=10.5,
      fills=[PALE["Módulo 2"], PALE["Base"], PALE["A · sesgo MMR"], PALE["B · ISAB"], RGBColor(0xF9, 0xDE, 0xDC)])
kpi(s, 6.95, 4.25, 1.9, 1.25, num(SL["Módulo 2"]), "pendiente por paso", RGBColor(0x55, 0x55, 0x55), "módulo 2")
kpi(s, 9.0, 4.25, 1.9, 1.25, num(SL["Base"]), "pendiente por paso", BLUE, "base (≤ 1: lineal)")
kpi(s, 11.05, 4.25, 1.9, 1.25, f"{num(SPD['50×40'], 0)}×", "más rápida", ORANGE, "instancia 50×40")
text(s, 6.95, 5.75, 6.0, 1.1,
     [[("Decoder lineal: ", {"bold": True, "color": NAVY}),
       ("es el mínimo posible, porque debe entregar un logit por cada uno de los n pares. La base y A no necesitan atención "
        "eficiente; B sí, y por eso usa ISAB.", {})]], size=12)

# ================================================================ 4. variante A
cob = CRIT["cobertura"]
s = new_slide("VARIANTE A", PURPLE, "Variante A: sesgo inductivo con el criterio de MMR",
              "Las dos formas del material: sesgo de salida en el puntero y sesgo atencional en el encoder",
              notes=f"(≈40 s) El criterio es la ganancia marginal de MMR, V_j por q_j por p_ij. Es robusto: es la heurística "
                    f"clásica del WTAP, óptima cuando las armas son iguales, y es la construcción inicial del estado del arte. "
                    f"En nuestros datos, la acción del experto es la de mayor ganancia en {num(100 * cob['1'], 0)}% de las "
                    f"decisiones y está entre las 3 mejores en {num(100 * cob['3'], 0)}%. Lo sumamos a los logits con un peso "
                    "alfa aprendible, y en el encoder sumamos la ganancia estática a los puntajes de atención con un alfa por "
                    "capa y cabeza. Como el puntero parte en cero, la variante A arranca exactamente desde MMR y aprende "
                    "correcciones.")
image(s, F("arq_variante_A.png"), 0.35, 1.35, w=8.0, h=5.6)
f1 = formula_png(r"$u_{ij} \leftarrow u_{ij} + \alpha\,\ln\frac{g_{ij}(t)}{g_{\max}(t)},\quad g_{ij}=V_j\,q_j\,p_{ij}$",
                 "fA1", fontsize=22, color="#7B3FB0")
f2 = formula_png(r"$s^{(l,h)}_{ij} \leftarrow s^{(l,h)}_{ij} + \alpha^{(l,h)}\ln\frac{V_j\,p_{ij}}{\max\,Vp}$", "fA2",
                 fontsize=22, color="#7B3FB0")
box(s, 8.55, 1.4, 4.4, 5.5, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
text(s, 8.7, 1.48, 4.2, 0.35, "Sesgo de salida (decoder, α inicia en 1)", size=12.5, bold=True, color=PURPLE)
image(s, f1, 8.75, 1.85, w=4.0, h=0.6)
text(s, 8.7, 2.5, 4.2, 0.35, "Sesgo atencional (encoder, α inician en 0)", size=12.5, bold=True, color=PURPLE)
image(s, f2, 8.75, 2.85, w=4.0, h=0.62)
text(s, 8.7, 3.6, 4.2, 0.35, "¿Por qué MMR? Es un criterio robusto", size=12.5, bold=True, color=NAVY)
text(s, 8.7, 3.95, 4.15, 1.6,
     [[("• Heurística clásica (den Broeder et al., 1959), óptima con armas iguales; construcción inicial del estado del "
        "arte (Ahuja et al., 2007).", {"space_after": 3})],
      [(f"• Gap {pct(CRIT['gap_mmr_val'], 1)} vs {pct(CRIT['gap_rand_val'], 0)} del aleatorio (validación).", {})]],
     size=10.5, spacing=1.0)
cob_rows = [["Acción del experto en el ranking de MMR", "top-1", "top-2", "top-3"],
            ["% de decisiones", pct(100 * cob["1"], 1), pct(100 * cob["2"], 1), pct(100 * cob["3"], 1)]]
table(s, 8.7, 5.75, 4.1, cob_rows, [2.2, 0.63, 0.63, 0.64], row_h=0.38, size=9.5,
      aligns=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.CENTER], header_fill=PURPLE)

# ================================================================ 5. variante B
s = new_slide("VARIANTE B", GREEN, "Variante B: decoder con atención eficiente (ISAB)",
              "En cada paso los pares factibles interactúan a través de 16 puntos inductores, en tiempo lineal",
              notes="(≈35 s) En la base, los pares solo compiten en el softmax final: un par no sabe si otra arma libre "
                    "aprovecharía mejor su objetivo. La variante B reemplaza el contexto por un bloque ISAB: 16 puntos "
                    "aprendibles resumen a los pares factibles y cada par lee ese resumen; el contexto es el resumen inducido. "
                    "Una self-attention completa entre pares costaría n al cuadrado por paso, lo mismo que el módulo 2; ISAB "
                    "cuesta n por m. Elegimos ISAB y no LUNA porque el estado es un conjunto sin orden que cambia en cada paso, "
                    "sin una memoria que deba persistir; y no atención dispersa porque requeriría un criterio heurístico, que "
                    "es el rol de la variante A.")
image(s, F("arq_variante_B.png"), 0.35, 1.35, w=8.0, h=5.6)
card(s, 8.55, 1.4, 4.4, 2.05, "Qué cambia respecto de la base",
     [[("Contexto (promedios + glimpse) → ", {}), ("ISAB sobre los pares", {"bold": True}),
       (": 16 puntos inductores resumen a los pares factibles y cada par lee ese resumen; el contexto es el resumen "
        "inducido.", {})]], accent=GREEN, title_size=13, body_size=11)
card(s, 8.55, 3.6, 4.4, 3.3, "¿Por qué ISAB?",
     [[("• ", {}), ("Lineal: ", {"bold": True}), ("O(n·m·d) por paso vs O(n²·d) de la atención completa ", {}),
       (f"(pendiente {num(SL['B · ISAB'])} vs {num(SL['B con atención completa'])}).", {}), ],
      [("• ", {}), ("Conjuntos: ", {"bold": True}),
       ("el estado es un conjunto sin orden que cambia en cada paso; ISAB (Set Transformer) crea un resumen nuevo en cada "
        "capa, sin la memoria persistente de LUNA.", {})],
      [("• ", {}), ("Sin heurística: ", {"bold": True}),
       ("la atención dispersa necesitaría un criterio para conectar pares; ese rol lo cubre la variante A.", {})]],
     accent=GREEN, title_size=13, body_size=11)

# ================================================================ 6. convergencia
h0 = {n: CONV[n] for n in MODELS}
s = new_slide("RESULTADOS", RED, "Convergencia: misma configuración, 3 modelos",
              f"Media de {NSEED} semillas (banda = mínimo–máximo); 6 000 instancias 10×8, Adam + ReduceLROnPlateau, "
              "early stopping",
              notes="(≈35 s) CONV_NOTES")
image(s, F("convergencia.png"), 0.35, 1.35, w=8.3, h=5.6)
rows = [["", "Gap val. época 0", "Mejor época", "Loss val. (mejor)", "s/época"]]
for n in MODELS:
    rows.append([n, pct(h0[n]["gap_init"], 1), "–".join(str(x) for x in sorted({min(h0[n]["best_epoch"]), max(h0[n]["best_epoch"])})),
                 pm(h0[n]["best_val_loss_mean"], h0[n]["best_val_loss_std"], 3), num(h0[n]["sec_per_epoch"], 1)])
table(s, 8.85, 1.45, 4.1, rows, [1.25, 0.8, 0.65, 0.85, 0.55], row_h=0.5, size=9,
      fills=[PALE[n] for n in MODELS], aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 4)
text(s, 8.85, 3.6, 4.1, 3.3, "CONV_BULLETS", size=11.5)

# ================================================================ 7. evaluación por tamaño
s = new_slide("RESULTADOS", RED, "Evaluación por tamaño: gap vs el experto ILS",
              "Entrenados solo con 10×8; mismas instancias del módulo 2 (30 por tamaño, 50 de prueba en 10×8)",
              notes="(≈45 s) EVAL_NOTES")
rows = [["Tamaño"] + ["Base", "A · MMR", "B · ISAB", "MMR", "Módulo 2"]]
bold = []
for r_, l in enumerate(SIZES, start=1):
    best = min(MODELS, key=lambda n: gm(l, n))
    row = [l]
    for c_, n in enumerate(MODELS, start=1):
        row.append(pm(EV[l][n]["mean"], EV[l][n]["std"]))
        if n == best:
            bold.append((r_, c_))
    row += [pm(EV[l]["MMR"]["mean"], EV[l]["MMR"]["std"]), pm(*REF2[l]) if l in REF2 else "—"]
    rows.append(row)
table(s, 0.45, 1.45, 7.6, rows, [1.05, 1.32, 1.32, 1.32, 1.32, 1.27], row_h=0.42, size=10.5,
      aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 5, bold_cells=bold)
text(s, 0.45, 1.5 + 0.42 * len(rows), 7.6, 0.35, "Gap (%) media ± desv. estándar entre instancias; en negrita, el mejor de "
     "los tres modelos.", size=10, color=MUTED)
cx = crop(F("comparacion.png"), "comparacion_der.png", (0.43, 0.0, 1.0, 1.0))
image(s, cx, 8.2, 1.4, w=4.8, h=3.6)
text(s, 8.25, 5.1, 4.75, 1.85, "EVAL_BULLETS", size=11.5)

# ================================================================ 8. conclusiones
s = new_slide("CONCLUSIONES", NAVY, "Conclusiones", None,
              notes="(≈20 s) CONCL_NOTES")
concl = [("Responsabilidades separadas", "CONCL_1", BLUE),
         ("Complejidad", "CONCL_2", ORANGE),
         ("Sesgo inductivo (A)", "CONCL_3", PURPLE),
         ("ISAB (B)", "CONCL_4", GREEN)]
for k, (t, d, col) in enumerate(concl):
    card(s, 0.6 + (k % 2) * 6.15, 1.35 + (k // 2) * 2.75, 5.95, 2.55, t, d, accent=col, title_size=15, body_size=13)

for sl, notes in SLIDES:
    sl.notes_slide.notes_text_frame.text = notes
prs.save(OUT)
print("Presentación guardada en", OUT, "·", len(prs.slides), "diapositivas")
