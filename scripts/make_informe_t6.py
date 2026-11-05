"""Genera el informe de la Tarea 6 (puntos 1, 2, 3, 5 y 6) a partir de figuras_t6/ (figuras y resultados.json).

Uso (desde la raíz del repositorio):
    python scripts/make_diagramas_t6.py
    python scripts/make_informe_t6.py
    soffice --headless --convert-to pdf Tarea6_WTAP_informe.docx
"""
import json
import os

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

GRUPO = "[Nombre(s) del grupo]"           # reemplazar por los nombres del grupo y volver a generar

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG = os.environ.get("T6_FIG", os.path.join(ROOT, "figuras_t6"))
OUT = os.environ.get("T6_OUT", os.path.join(ROOT, "Tarea6_WTAP_informe.docx"))
R = json.load(open(os.path.join(FIG, "resultados.json"), encoding="utf-8"))
TMP = os.path.join(FIG, "_tmp_informe")
os.makedirs(TMP, exist_ok=True)

FONT = "Arial"
NAVY = RGBColor(0x1B, 0x2A, 0x41)
BLUE = RGBColor(0x1F, 0x77, 0xB4)
MUTED = RGBColor(0x6B, 0x75, 0x85)
INK = RGBColor(0x2E, 0x35, 0x40)
MODELS = ["Base", "A · sesgo MMR", "B · ISAB"]
FILL = {"Base": "DCEAF7", "A · sesgo MMR": "ECE4F5", "B · ISAB": "DFF0DF", "MMR": "FDEBD9", "Aleatorio": "EEEEEE",
        "Módulo 2": "E8EEF6"}

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)      # carta
sec.left_margin = sec.right_margin = Cm(2.0)
sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.8)
TEXT_W = 21.59 - 4.0

st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(10.5)
st.font.color.rgb = INK
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(5)
st.paragraph_format.line_spacing = 1.12
for name, size in [("Heading 1", 15), ("Heading 2", 12)]:
    h = doc.styles[name]
    h.font.name, h.font.size, h.font.bold = FONT, Pt(size), True
    h.font.color.rgb = NAVY
    rfonts = h.element.rPr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts"); h.element.rPr.append(rfonts)
    for att in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(att), FONT)
    for att in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rfonts.get(qn(att)) is not None:
            del rfonts.attrib[qn(att)]
    h.paragraph_format.space_before = Pt(12 if name == "Heading 1" else 8)
    h.paragraph_format.space_after = Pt(5)
    h.paragraph_format.keep_with_next = True


# ---------------------------------------------------------------- helpers (base: informe de la Tarea 5)
def _runs(runs):
    # str, o lista de str / (texto, opciones)
    if isinstance(runs, str):
        return [(runs, {})]
    return [(r, {}) if isinstance(r, str) else r for r in runs]


def para(runs, style=None, align=None, size=None, space_after=None, keep=False):
    """runs: str o lista de (texto, {bold, italic, color, font, size})."""
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if keep:
        p.paragraph_format.keep_with_next = True
    for txt, opt in _runs(runs):
        r = p.add_run(txt)
        r.bold = opt.get("bold", False)
        r.italic = opt.get("italic", False)
        if "color" in opt:
            r.font.color.rgb = opt["color"]
        if "font" in opt:
            r.font.name = opt["font"]
        if size or "size" in opt:
            r.font.size = Pt(opt.get("size", size))
        if opt.get("sub"):
            r.font.subscript = True
        if opt.get("sup"):
            r.font.superscript = True
    return p


def bullet(runs):
    p = para(runs, style="List Bullet", space_after=2)
    p.paragraph_format.left_indent = Cm(0.9)
    return p


def c(txt):
    return (txt, {"font": "Liberation Mono", "size": 9.5})


def b(txt):
    return (txt, {"bold": True})


def i(txt):
    return (txt, {"italic": True})


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def cell_text(cell, runs, size=9, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    for txt, opt in _runs(runs):
        r = p.add_run(txt)
        r.font.size = Pt(opt.get("size", size))
        r.bold = opt.get("bold", bold)
        r.italic = opt.get("italic", False)
        if "font" in opt:
            r.font.name = opt["font"]
        if color is not None:
            r.font.color.rgb = color
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def table(rows, widths, header_fill="1B2A41", fills=None, size=9, aligns=None, bold_rows=()):
    t = doc.add_table(rows=len(rows), cols=len(widths))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for c_, w in enumerate(widths):
        t.columns[c_].width = Cm(w)
    for r_, row in enumerate(rows):
        for c_, val in enumerate(row):
            cell = t.cell(r_, c_)
            cell.width = Cm(widths[c_])
            al = (aligns[c_] if aligns else WD_ALIGN_PARAGRAPH.LEFT)
            if r_ == 0:
                cell_text(cell, val, size=size, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), align=al)
                shade(cell, header_fill)
            else:
                cell_text(cell, val, size=size, align=al, bold=r_ in bold_rows)
                if fills and fills[r_ - 1]:
                    shade(cell, fills[r_ - 1])
    for row in t.rows:                                   # evita que una fila se parta entre páginas
        trPr = row._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit"); cant.set(qn("w:val"), "true"); trPr.append(cant)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figure(path, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(path, width=Cm(width_cm))
    para([(caption, {"italic": True, "color": MUTED})], align=WD_ALIGN_PARAGRAPH.CENTER, size=9, space_after=8)


def formula(latex, name, height_cm=1.0, fontsize=17):
    path = os.path.join(TMP, f"{name}.png")
    fig = plt.figure(figsize=(8, 0.6))
    fig.text(0, 0.5, latex, fontsize=fontsize, color="#1B2A41", va="center")
    fig.savefig(path, dpi=300, bbox_inches="tight", pad_inches=0.04, transparent=True)
    plt.close(fig)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    p.add_run().add_picture(path, height=Cm(height_cm))


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def num(v, d=2):
    return f"{v:.{d}f}".replace(".", ",")


def pct(v, d=2):
    return num(v, d) + " %"


def pm(m, s, d=2):
    return f"{num(m, d)} ± {num(s, d)}"


def miles(n):
    return f"{n:,}".replace(",", ".")


def F(name):
    return os.path.join(FIG, name)


EV = R["eval"]
SIZES = list(EV)
CONV = R["convergence"]
HIST = R["histories"]
CRIT = R["criterio_mmr"]
TIMING = R["timing"]
SL = R["slopes_step"]
SLE = R["slopes_encoder"]
SLT = R["slopes_total"]
SPD = R["speedup_vs_mod2"]
REF2 = R["ref_modulo2"]
REFL = R.get("ref_ils_largo", {})
DIAG = R["diagnostico_modulo2"]
PAR = R["params"]
NSEED = R["n_seeds"]


def gm(label, name):
    return EV[label][name]["mean"]


def best_model(label):
    return min(MODELS, key=lambda n: gm(label, n))


# ================================================================ portada / encabezado
para([("OII462 · Tarea 6", {"bold": True, "color": BLUE, "size": 10})], space_after=0)
para([("Encoder estático, complejidad y variantes del Transformer para el WTAP", {"bold": True, "color": NAVY, "size": 19})],
     space_after=2)
para([("Informe: rediseño de la arquitectura, complejidad temporal, variantes y comparación experimental "
       "(puntos 1, 2, 3, 5 y 6)", {"color": MUTED, "size": 11})], space_after=2)
para([(GRUPO, {"bold": True}), ("  ·  Optimización combinatoria  ·  2026", {"color": MUTED})], size=10, space_after=10)

t10 = "10×8"
para([b("Resumen. "),
      (f"El Transformer con puntero del módulo 2 recibía en cada paso una secuencia de 1 + W·T filas con 16 features, de las "
       f"cuales {DIAG['n_dinamicas']} dependen del estado; por eso su \"encoder\" se recalculaba en cada una de las W "
       "decisiones, con un costo cuadrático en el número de pares por paso. Lo rediseñamos con responsabilidades separadas: "
       "un encoder bipartito armas ↔ objetivos que solo ve V", {}), ("j", {"sub": True}),
      (" y p", {}), ("ij", {"sub": True}),
      (" (token aprendible para las armas y la matriz como aristas) y que corre una vez por instancia, y un decoder que "
       "integra el estado y es lineal en n = W·T por paso. Sobre esa base proponemos dos variantes: ", {}),
      b("A"), (", con sesgo inductivo (de salida y atencional) basado en la ganancia marginal de MMR, y ", {}), b("B"),
      (", que re-contextualiza los pares en cada paso con un bloque ISAB de costo lineal. Una decisión de 50×40 pasa de "
       f"{num(1e3 * TIMING['Módulo 2']['dec'][TIMING['Módulo 2']['n'].index(2000)], 1)} ms a "
       f"{num(1e3 * TIMING['Base']['dec'][TIMING['Base']['n'].index(2000)], 2)} ms. ", {}),
      ("RESUMEN_RESULTADOS", {})])

# ================================================================ 1. rediseño
doc.add_heading("1. Rediseño de la arquitectura (punto 1)", level=1)
doc.add_heading("1.1 Diagnóstico del modelo del módulo 2", level=2)
para(["En el módulo 2 la entrada era una matriz (1 + W·T) × 16: una fila global y una fila por par (arma i → objetivo j). "
      "El bloque de self-attention de 2 capas y el puntero se ejecutaban juntos en cada paso, porque la matriz se recalculaba "
      "con state2vec tras cada asignación. Recorriendo una construcción completa verificamos qué columnas cambian entre pasos:"])
table([["Tipo", "Features del módulo 2"],
       [f"Estáticas ({len(DIAG['estaticas'])})", ", ".join(DIAG["estaticas"])],
       [f"Dinámicas ({DIAG['n_dinamicas']})", ", ".join(DIAG["dinamicas"])]],
      [3.2, TEXT_W - 3.6], fills=["DCEAF7", "F9DEDC"])
para([("Ese \"encoder\" no cumplía la definición de la tarea: dependía del estado y no podía ejecutarse una sola vez. Su costo "
       "por paso era la self-attention sobre 1 + W·T filas, O((WT)", {}), ("2", {"sup": True}),
      (") por capa, más state2vec, que es O(W", {}), ("2", {"sup": True}),
      ("T) porque mmr_completa simula MMR hasta el final.", {})])

doc.add_heading("1.2 Arquitectura base rediseñada", level=2)
figure(F("arq_base.png"), TEXT_W,
       "Figura 1. Arquitectura base rediseñada. A la izquierda, el encoder estático (1 vez por instancia); a la derecha, el "
       "decoder, que integra el estado en cada uno de los W pasos. Los círculos naranjos marcan los cambios respecto del "
       "módulo 2.")
para([b("Encoder (estático). "),
      ("Las armas no tienen información propia: solo su fila de la matriz p. Siguiendo la sugerencia del enunciado, cada arma "
       "parte de un ", {}), b("token aprendible"),
      (" compartido y cada objetivo de un embedding lineal de V", {}), ("j", {"sub": True}), ("/max V. La matriz entra como ", {}),
      b("aristas de un grafo bipartito"),
      (": en cada una de las 2 capas, las armas atienden a los objetivos y los objetivos a las armas (en paralelo, MHA de 4 "
       "cabezas), y la arista p̃", {}), ("ij", {"sub": True}),
      (" se inyecta en el puntaje (un peso por cabeza) y en el mensaje (un vector aprendido), como en la EGAT del material y en "
       "MatNet (Kwon et al., 2021). Tras cada atención hay residual + LayerNorm y un feed-forward 64→128→64. La salida son "
       "los embeddings H", {}), ("w", {"sup": True}), (" (W×d) y H", {}), ("t", {"sup": True}),
      (" (T×d) y las partes del puntero que no dependen del estado, que se precalculan una vez.", {})])
formula(r"$s^{(h)}_{ij} = \frac{(W_Q h^w_i)^{(h)}\cdot (W_K h^t_j)^{(h)}}{\sqrt{d_h}} + \beta^{(h)}\tilde p_{ij},"
        r"\qquad m_i = W_O[\sum_j \mathrm{softmax}_j(s^{(h)}_{ij})\,((W_V h^t_j)^{(h)} + w_e^{(h)}\tilde p_{ij})]_h$",
        "encoder", height_cm=1.15)
para([b("Decoder (dinámico). "),
      ("Recibe el estado de la solución parcial: la máscara de armas libres y, por objetivo, q", {}), ("j", {"sub": True}),
      (", v", {}), ("j", {"sub": True}), ("q", {}), ("j", {"sub": True}), (" y ln(v", {}), ("j", {"sub": True}), ("q", {}),
      ("j", {"sub": True}),
      ("). Suma un MLP de ese estado al embedding de cada objetivo, arma un contexto con el promedio de las armas libres, el "
       "de los objetivos y el progreso, hace un glimpse (MHA de 4 cabezas, Kool et al., 2019) y puntúa cada par con un ", {}),
      b("puntero aditivo"), (" que también ve la arista:", {})])
formula(r"$u_{ij} = C\,\tanh((W_r c)^{\top}\,\mathrm{ReLU}(W_A h^w_i + W_B h_j(t) + W_E\tilde p_{ij})/\sqrt{d}),"
        r"\quad u_{ij}=-10^9 \;\mathrm{si}\; i \;\mathrm{ya\;fue\;usada}$", "pointer", height_cm=0.85)
para(["El modelo ya no recibe features heurísticas (ganancia, arrepentimiento, anticipación con MMR): debe aprenderlas de "
      "V, p y q. Así la información dinámica entra solo al decoder y el encoder cumple la definición de la tarea."])

doc.add_heading("1.3 Cambios respecto del módulo 2", level=2)
table([["#", "Aspecto", "Módulo 2", "Base rediseñada"],
       ["1", "Tokens de entrada", "1 + W·T filas (fila global + pares) con 16 features",
        "W armas (token aprendible) + T objetivos (V_j) + matriz p_ij como aristas"],
       ["2", "Encoder", "self-attention entre las 1 + W·T filas, 1 cabeza, se ejecuta en cada paso",
        "atención cruzada armas ↔ objetivos con aristas, 4 cabezas, 1 vez por instancia"],
       ["3", "Información dinámica", "dentro de las features de entrada (12 de 16)",
        "solo en el decoder: máscara, q_j, valor en juego y progreso; sin features heurísticas"],
       ["4", "Contexto del decoder", "[h_global ; promedio de filas disponibles]",
        "[promedio de armas libres ; promedio de objetivos con estado ; progreso] + glimpse MHA"],
       ["5", "Puntero", "q·h_r/√d sobre las filas", "aditivo sobre pares, con la arista p_ij y recorte C·tanh"],
       ["", "Entrenamiento", "60 000 estados independientes",
        "teacher forcing: 1 encoder + W pasos del decoder por instancia"]],
      [0.6, 3.0, 6.4, 7.6], fills=[None] * 6, size=8.8)

# ================================================================ 2. complejidad
doc.add_heading("2. Complejidad temporal del encoder y del decoder (punto 2)", level=1)
para([("Sea n = W·T el tamaño de la matriz p, que es también el número de acciones; d = 64 la dimensión, L = 2 las capas y "
       "m = 16 los puntos inductores de la variante B. Un decoder que entrega un logit por acción necesita al menos Ω(n) por "
       "paso, así que ese es el mínimo alcanzable.", {})])
table([["Modelo", "Encoder (1 vez por instancia)", "Cómputo por paso del decoder", "Por instancia (W pasos)"],
       ["Módulo 2", "— (todo se recalcula en cada paso)", "state2vec O(W²T) + self-attention O(L·n²·d)", "O(W·n²·d)"],
       ["Base", "O(L·(n·d + (W+T)·d²))", "O(n·d + (W+T)·d²)  → lineal", "O(W·n·d)"],
       ["A · sesgo MMR", "igual a la base + O(L·n)", "base + O(n) (ganancias)  → lineal", "O(W·n·d)"],
       ["B · ISAB", "igual a la base", "O(n·m·d + n·d²)  → lineal", "O(W·n·m·d)"],
       ["B con self-attention completa", "igual a la base", "O(n²·d)  → cuadrático", "O(W·n²·d)"]],
      [3.6, 4.4, 5.6, 3.9], fills=["E8EEF6", FILL["Base"], FILL["A · sesgo MMR"], FILL["B · ISAB"], "F9DEDC"], size=8.8)
para(["El encoder rediseñado es incluso lineal en n, porque la atención es bipartita (cada arma mira T objetivos y cada "
      "objetivo W armas), aunque la tarea permitiría uno cuadrático. El decoder de la base y de la variante A es lineal, así "
      "que no necesita un mecanismo de atención eficiente. La única pieza cuadrática que aparecería en el decoder es la "
      "interacción entre pares de la variante B; por eso la implementamos con ISAB (sección 3.2)."])
figure(F("complejidad.png"), TEXT_W,
       "Figura 2. Tiempos medidos con PyTorch en 1 hilo (mediana de varias repeticiones, pesos aleatorios). Izquierda: una "
       "decisión; centro: el encoder; derecha: total estimado por instancia (encoder + W pasos). Pendientes log-log ajustadas "
       "con los 4 tamaños más grandes de cada curva.")
k10 = TIMING["Base"]["n"].index(80)
k2000b, k2000m = TIMING["Base"]["n"].index(2000), TIMING["Módulo 2"]["n"].index(2000)
table([["Modelo", "Pendiente por paso", "Pendiente del encoder", "Pendiente por instancia", "Paso en 10×8", "Paso en 50×40"],
       *[[name, num(SL[name]), num(SLE[name]) if name in SLE else "—", num(SLT[name]),
          f"{num(1e3 * TIMING[name]['dec'][TIMING[name]['n'].index(80)], 2)} ms",
          f"{num(1e3 * TIMING[name]['dec'][TIMING[name]['n'].index(2000)], 2)} ms"]
         for name in ["Módulo 2", "Base", "A · sesgo MMR", "B · ISAB", "B con atención completa"]]],
      [4.6, 2.5, 2.6, 2.7, 2.6, 2.6], size=8.8,
      aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 5,
      fills=["E8EEF6", FILL["Base"], FILL["A · sesgo MMR"], FILL["B · ISAB"], "F9DEDC"])
para([b("Lectura. "),
      (f"Las pendientes confirman el análisis: el paso del módulo 2 y el de B con self-attention completa crecen como n² "
       f"(pendientes {num(SL['Módulo 2'])} y {num(SL['B con atención completa'])}), mientras que el de los tres modelos "
       f"nuevos crece a lo más linealmente (pendientes {num(SL['Base'])}, {num(SL['A · sesgo MMR'])} y {num(SL['B · ISAB'])}; "
       "en tamaños pequeños domina el costo fijo de lanzar las operaciones, por eso la pendiente queda bajo 1). El encoder "
       f"cuesta {num(1e3 * TIMING['Base']['enc'][k10], 2)} ms en 10×8 y se paga una sola vez. Por instancia, la base es "
       f"{num(SPD['10×8'], 1)} veces más rápida que el módulo 2 en 10×8 y {num(SPD['50×40'], 0)} veces en 50×40.", {})])

# ================================================================ 3. variantes
doc.add_heading("3. Variantes propuestas (punto 3)", level=1)
doc.add_heading("3.1 Variante A — sesgo inductivo con el criterio de MMR", level=2)
figure(F("arq_variante_A.png"), TEXT_W * 0.93,
       "Figura 3. Variante A: la base más el sesgo inductivo (en morado), atencional en el encoder y de salida en el puntero.")
para(["Aplicamos las dos formas de sesgo inductivo vistas en el material, ambas con la ganancia marginal de MMR, "
      "g_ij = V_j·q_j·p_ij, y con pesos aprendibles para que el modelo decida cuánto confiar en el sesgo:"])
bullet([b("Sesgo de salida (decoder, dinámico): "),
        ("u_ij ← u_ij + α·ln(g_ij(t)/g_max(t)), con α aprendible que parte en 1. La política queda π(i, j) ∝ e^u_ij·(g_ij/g_max)^α: "
         "con u = 0 es un \"MMR suave\" y con α → ∞ es MMR exacto, así que la red aprende correcciones sobre MMR. Como "
         "inicializamos W_r = 0 en los tres modelos, la variante A parte exactamente desde la política de MMR.", {})])
bullet([b("Sesgo atencional (encoder, estático): "),
        ("en cada capa, dirección y cabeza se suma α^(l,dir,h)·ln(V_j p_ij / max Vp) al puntaje armas ↔ objetivos (la ganancia "
         "de MMR con la asignación vacía), con α^(l,dir,h) aprendibles que parten en 0, como en el material.", {})])
para([b("¿Por qué MMR es un criterio robusto? "),
      ("(i) Es la heurística constructiva clásica del WTAP (den Broeder et al., 1959) y es óptima cuando todas las armas son "
       "iguales: si p_ij = p_j el problema es una asignación separable y convexa, donde elegir siempre el mayor retorno "
       "marginal es óptimo. (ii) Es la construcción inicial de métodos del estado del arte (Ahuja et al., 2007; ver también "
       "la revisión de Kline et al., 2019) y de nuestro propio experto ILS. (iii) En nuestros datos produce soluciones muy "
       f"superiores al promedio: gap {pct(CRIT['gap_mmr_val'])} frente a {pct(CRIT['gap_rand_val'], 1)} del constructivo "
       "aleatorio en 100 instancias de validación, y la acción del experto está casi siempre entre las de mayor ganancia:", {})])
cob = CRIT["cobertura"]
table([["Posición de la acción del experto en el ranking de MMR", "top-1", "top-2", "top-3", "top-5", "top-10"],
       ["% de las 3 000 decisiones de validación"] + [pct(100 * cob[k], 1) for k in ("1", "2", "3", "5", "10")]],
      [7.5, 2.0, 2.0, 2.0, 2.0, 2.0], size=9, aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 5)
para(["Es decir, MMR acierta la decisión exacta del experto en ≈80 % de los casos y la tiene entre sus 3 mejores en ≈99 %: "
      "el sesgo empuja al modelo hacia una región muy buena de acciones y le deja aprender cuál elegir dentro de ella."])

doc.add_heading("3.2 Variante B — decoder con atención eficiente (ISAB)", level=2)
figure(F("arq_variante_B.png"), TEXT_W * 0.93,
       "Figura 4. Variante B: el contexto de la base (promedios + glimpse) se reemplaza por un bloque ISAB sobre los pares "
       "factibles (en verde).")
para(["En la base, los pares solo compiten en el softmax final: un par no \"sabe\" si otra arma libre aprovecharía mejor su "
      "objetivo, información que en el módulo 2 aportaban el arrepentimiento y la self-attention entre filas. La variante B "
      "re-contextualiza el estado en cada paso: los pares factibles Z_ij(t) interactúan mediante un Induced Set Attention Block "
      "(Lee et al., 2019) con m = 16 puntos inductores aprendibles. Primero los puntos resumen a los pares, I' = "
      "LN(I + MHA(I, Z, Z)) + FF; luego cada par lee ese resumen, Z' = LN(Z + MHA(Z, I', I')) + FF. El contexto pasa a ser "
      "el resumen inducido, c = W_c[promedio de I' ; progreso], y el puntero usa los pares contextualizados."])
para([b("Justificación del mecanismo. "),
      ("Una self-attention completa entre los n pares costaría O(n²·d) en cada paso: es justamente el costo del módulo 2 y "
       "rompe la exigencia de un decoder lineal (Figura 2, curva roja). ISAB la reemplaza por dos atenciones cruzadas de costo "
       "O(n·m·d). Lo preferimos a LUNA porque el estado del WTAP es un conjunto sin orden que cambia en cada paso (desaparecen "
       "los pares del arma asignada): no hay una secuencia ni una memoria que deba persistir entre capas, que es lo que aporta "
       "el depósito de LUNA, mientras que ISAB es el bloque del Set Transformer para conjuntos y construye un resumen nuevo en "
       "cada capa. La atención dispersa requeriría un criterio heurístico para decidir qué pares conectar, que es el papel del "
       "sesgo de la variante A; así cada variante aísla un efecto.", {})])

# ================================================================ 4. configuración
doc.add_heading("4. Configuración experimental (punto 4)", level=1)
cfg = R["train_cfg"]
table([["Hiperparámetro", "Valor (igual para los tres modelos)"],
       ["d / cabezas / capas del encoder / FF", "64 / 4 / 2 / 64 → 128 → 64"],
       ["Normalización · recorte de logits", "LayerNorm post-residual · C = 10"],
       ["Inicialización del puntero", "W_r = 0 (los logits parten en 0)"],
       ["Datos", f"{miles(R['n_train'])} instancias 10×8 de entrenamiento ({miles(R['n_train_decisions'])} decisiones), "
                 f"300 de validación ({miles(R['n_val_decisions'])}); mismas semillas que en las tareas anteriores"],
       ["Optimizador", f"Adam, lr = {cfg['lr']:g}; ReduceLROnPlateau (factor {num(cfg['lr_factor'], 1)}, paciencia "
                       f"{cfg['lr_patience']}, mínimo {cfg['lr_min']:g}), como sl_train del material"],
       ["Batch · recorte de gradiente", f"{cfg['batch_size']} instancias ({cfg['batch_size'] * 10} decisiones) · norma "
                                        f"{num(cfg['grad_clip'], 1)}"],
       ["Épocas · early stopping", f"máx. {cfg['max_epochs']} · paciencia {cfg['patience']} sobre la pérdida de validación "
                                   "(restaura la mejor época)"],
       ["Pérdida", "entropía cruzada sobre los W·T logits enmascarados (teacher forcing)"],
       ["Semillas", f"{NSEED} (inicialización y orden de los batches); los datos son los mismos"],
       ["Específicos de cada variante", "A: α inicia en 1 y α^(l,dir,h) en 0 · B: m = 16 puntos inductores"],
       ["Parámetros entrenables", " · ".join(f"{n}: {miles(PAR[n])}" for n in MODELS)]],
      [5.2, TEXT_W - 5.6], size=9)

# ================================================================ 5. convergencia
doc.add_heading("5. Convergencia (punto 5)", level=1)
figure(F("convergencia.png"), TEXT_W,
       f"Figura 5. Curvas por época (línea = media de {NSEED} semillas; banda = mínimo–máximo). La época 0 es el modelo sin "
       "entrenar; el gap de validación se mide con el rollout greedy en 100 instancias.")


def lst(v):
    v = [x for x in v if x is not None]
    return "–".join(str(x) for x in (sorted(set(v)) if len(set(v)) <= 1 else [min(v), max(v)])) if v else "no alcanza"


rows = [["Modelo", "Gap val. época 0", "Épocas hasta gap < 3 %", "Épocas hasta pérdida < 0,35", "Mejor época",
         "Mejor pérdida val. (media ± desv.)", "Oscilación del gap (pp)", "s / época", "Minutos hasta gap < 3 %"]]
for n in MODELS:
    cv = CONV[n]
    rows.append([n, pct(cv["gap_init"], 1), lst(cv["epoch_gap_below_3"]), lst(cv["epoch_loss_below_035"]),
                 lst(cv["best_epoch"]), pm(cv["best_val_loss_mean"], cv["best_val_loss_std"], 3),
                 num(cv["oscillation_gap"], 2), num(cv["sec_per_epoch"], 1), num(cv["minutes_to_gap3"], 1)])
table(rows, [2.5, 1.6, 1.8, 1.9, 1.6, 2.6, 1.8, 1.4, 1.9], size=8.3, fills=[FILL[n] for n in MODELS],
      aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 8)
para([("Épocas: rango entre las semillas. Oscilación del gap: promedio de |Δ gap de validación| entre épocas consecutivas en "
       "las últimas 10 épocas, en puntos porcentuales. Minutos: tiempo de reloj medio hasta el primer gap de validación "
       "< 3 % (CPU, incluye la validación de cada época).", {"italic": True, "color": MUTED, "size": 9})])
para([b("Análisis. "), ("ANALISIS_CONVERGENCIA", {})])
if os.path.exists(F("alpha.png")):
    figure(F("alpha.png"), TEXT_W * 0.9,
           "Figura 6. Variante A: evolución del peso α del sesgo de salida y pesos del sesgo atencional aprendidos (semilla 0).")

# ================================================================ 6. evaluación
doc.add_heading("6. Evaluación con instancias de distintos tamaños (punto 6)", level=1)
para([("Los modelos se entrenaron solo con 10×8 y se evalúan sin cambios en ", {}), (f"{len(SIZES)} tamaños", {"bold": True}),
      (", con las mismas instancias del proyecto del módulo 2 (30 por tamaño; las 50 de prueba en 10×8), de modo que la columna "
       "\"Módulo 2\" es directamente comparable; 50×40 y 100×80 son nuevos. La referencia es el experto ILS de las tareas "
       "anteriores (óptimo en todas las instancias verificadas de 5×4 y 10×8): gap = 100·(costo/costo ILS − 1). Se evalúan "
       f"los {NSEED} modelos entrenados de cada arquitectura (una por semilla) con el rollout greedy, que verificamos idéntico "
       f"al GreedyAgent de eda; la tabla reporta la media y la desviación estándar sobre las {NSEED} semillas × las "
       "instancias de cada tamaño, y más abajo se muestra la variación entre semillas.", {})])
hdr = ["Tamaño", "W·T"] + MODELS + ["MMR", "Aleatorio", "Módulo 2"]
rows = [hdr]
for l in SIZES:
    Wl, Tl = map(int, l.split("×"))
    best = best_model(l)
    row = [l, str(Wl * Tl)]
    for n in MODELS:
        row.append([(pm(EV[l][n]["mean"], EV[l][n]["std"]), {"bold": n == best})])
    row += [pm(EV[l]["MMR"]["mean"], EV[l]["MMR"]["std"]), pm(EV[l]["Aleatorio"]["mean"], EV[l]["Aleatorio"]["std"], 0),
            pm(*REF2[l]) if l in REF2 else "—"]
    rows.append(row)
table(rows, [1.6, 1.1, 2.4, 2.4, 2.4, 2.3, 2.6, 2.3], size=8.5,
      aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 7)
para([(f"Gap vs el experto ILS (%), media ± desviación estándar sobre {NSEED} semillas × instancias (MMR y el constructivo "
       "aleatorio son deterministas: solo instancias); en negrita, el mejor de los tres modelos en cada tamaño. La columna "
       "Módulo 2 es un único modelo, evaluado en las mismas instancias.", {"italic": True, "color": MUTED, "size": 9})])
rows = [["Tamaño"] + [f"{n}: media ± desv. entre semillas (rango)" for n in MODELS]]
for l in SIZES:
    row = [l]
    for n in MODELS:
        sm = EV[l][n]["seed_means"]
        row.append(f"{pm(np.mean(sm), np.std(sm))}  ({num(min(sm))}–{num(max(sm))})")
    rows.append(row)
para([b("Variación entre semillas. "),
      ("Gap medio de cada uno de los modelos entrenados, resumido por arquitectura:", {})], keep=True)
table(rows, [1.8, 5.2, 5.2, 5.2], size=8.5, aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 3)
figure(F("tamanos.png"), TEXT_W,
       f"Figura 7. Gap medio por tamaño: barras = media de {NSEED} semillas × instancias, barras de error = ±1 desviación "
       "estándar, puntos negros = gap medio de cada semilla.")
figure(F("comparacion.png"), TEXT_W,
       f"Figura 8. Izquierda: distribución del gap en las 50 instancias de prueba 10×8 × {NSEED} semillas. Derecha: gap medio "
       "por tamaño (banda = rango entre semillas).")
d10 = EV["10×8"]
table([[f"Método (prueba 10×8; {NSEED} × 50 ejecuciones)", "Gap medio", "Desv.", "Mediana", "% óptimo", "Gana a MMR"],
       *[[n, pct(d10[n]["mean"]), num(d10[n]["std"]), pct(d10[n]["median"]), pct(d10[n]["pct_optimal"], 0),
          pct(d10[n]["wins_vs_mmr"], 0)] for n in MODELS + ["MMR"]],
       ["Módulo 2 (proyecto, 1 modelo)", pct(REF2["10×8"][0]), num(REF2["10×8"][1]), "1,85 %", "30 %", "72 %"]],
      [5.4, 2.1, 1.8, 2.1, 2.1, 2.4], size=9, aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 5,
      fills=[FILL[n] for n in MODELS] + [FILL["MMR"], FILL["Módulo 2"]])
if REFL:
    rows = [["Tamaño", "ILS largo mejora al ILS de referencia en"] + [f"{n} vs ILS largo" for n in MODELS] + ["MMR vs ILS largo"]]
    for l, dd in REFL.items():
        rows.append([l, pct(dd["mejora_ils_largo"])] + [pct(dd["gap_vs_ils_largo"][n]) for n in MODELS + ["MMR"]])
    para([b("¿Qué tan buena es la referencia en instancias grandes? "),
          ("Repetimos el ILS con 2 000 iteraciones y perturbación proporcional a W en los tres tamaños más grandes:", {})],
         keep=True)
    table(rows, [1.7, 3.6, 2.6, 2.7, 2.6, 2.6], size=8.5, aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 5)
para([b("Análisis. "), ("ANALISIS_EVALUACION", {})])

# ================================================================ 7. conclusiones
doc.add_heading("7. Conclusiones", level=1)
for txt in ["CONCLUSION_1", "CONCLUSION_2", "CONCLUSION_3", "CONCLUSION_4"]:
    bullet([(txt, {})])

doc.add_heading("Referencias", level=1)
refs = ["Ahuja, R. K., Kumar, A., Jha, K. C. y Orlin, J. B. (2007). Exact and heuristic algorithms for the weapon-target "
        "assignment problem. Operations Research, 55(6), 1136–1146.",
        "den Broeder, G. G., Ellison, R. E. y Emerling, L. (1959). On optimum target assignments. Operations Research, 7(3), "
        "322–326.",
        "Kline, A., Ahner, D. y Hill, R. (2019). The weapon-target assignment problem. Computers & Operations Research, 105, "
        "226–236.",
        "Kool, W., van Hoof, H. y Welling, M. (2019). Attention, learn to solve routing problems! ICLR.",
        "Kwon, Y.-D., Choo, J., Yoon, I., Park, M., Park, D. y Gwon, Y. (2021). Matrix encoding networks for neural "
        "combinatorial optimization. NeurIPS.",
        "Lee, J., Lee, Y., Kim, J., Kosiorek, A., Choi, S. y Teh, Y. W. (2019). Set Transformer: a framework for "
        "attention-based permutation-invariant neural networks. ICML.",
        "Ma, X., Kong, X., Wang, S., Zhou, C., May, J., Ma, H. y Zettlemoyer, L. (2021). Luna: Linear unified nested "
        "attention. NeurIPS.",
        "Material del curso (TSP-Framework): notebooks de sesgo inductivo, atención lineal (LUNA, ISAB), esparsificación y GAT/EGAT."]
for r_ in refs:
    p = para(r_, size=9, space_after=3)
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.6)

# pie de página con número
for s_ in doc.sections:
    fp = s_.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ = fp.add_run("OII462 · Tarea 6 · WTAP  —  ")
    r_.font.size = Pt(8); r_.font.color.rgb = MUTED
    for kind, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        run = fp.add_run()
        run.font.size = Pt(8); run.font.color.rgb = MUTED
        if kind:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), kind); run._r.append(el)
        else:
            el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = txt; run._r.append(el)

doc.save(OUT)
print("Informe guardado en", OUT)
