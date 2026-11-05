"""Genera el informe de la Tarea 6 (puntos 1, 2, 3, 5 y 6) a partir de figuras_t6/ (figuras y resultados.json).

Es una versión en lenguaje simple: el detalle técnico (fórmulas, hiperparámetros y todas las mediciones) está en el
notebook Tarea6_WTAP_variantes.ipynb. Todas las cifras del texto se leen de resultados.json.

Uso (desde la raíz del repositorio):
    python scripts/make_diagramas_t6.py
    python scripts/make_informe_t6.py
    soffice --headless --convert-to pdf Tarea6_WTAP_informe.docx
"""
import json
import os

import numpy as np
from PIL import Image
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
FILL = {"Base": "DCEAF7", "A · sesgo MMR": "ECE4F5", "B · ISAB": "DFF0DF", "MMR": "FDEBD9", "Módulo 2": "E8EEF6"}

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
    """runs: str o lista de (texto, {bold, italic, color, size, sub, sup})."""
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
        if size or "size" in opt:
            r.font.size = Pt(opt.get("size", size))
        if opt.get("sub"):
            r.font.subscript = True
        if opt.get("sup"):
            r.font.superscript = True
    return p


def bullet(runs):
    p = para(runs, style="List Bullet", space_after=3)
    p.paragraph_format.left_indent = Cm(0.9)
    return p


def b(txt):
    return (txt, {"bold": True})


def sub(txt):
    return (txt, {"sub": True})


def note(txt):
    """Nota en letra pequeña bajo una tabla."""
    return para([(txt, {"italic": True, "color": MUTED, "size": 9})])


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
        if color is not None:
            r.font.color.rgb = color
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def table(rows, widths, header_fill="1B2A41", fills=None, size=9, aligns=None):
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
                cell_text(cell, val, size=size, align=al)
                if fills and fills[r_ - 1]:
                    shade(cell, fills[r_ - 1])
    for k, row in enumerate(t.rows):                     # evita que una fila (y la tabla) se parta entre páginas
        trPr = row._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit"); cant.set(qn("w:val"), "true"); trPr.append(cant)
        if k < len(t.rows) - 1:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_with_next = True
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figure(path, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(path, width=Cm(width_cm))
    para([(caption, {"italic": True, "color": MUTED})], align=WD_ALIGN_PARAGRAPH.CENTER, size=9, space_after=8)


def crop(path, name, box_frac):
    """Recorta una figura (fracciones x0, y0, x1, y1) y la guarda en la carpeta temporal."""
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
    return num(v, d) + " %"


def pm(m, s, d=2):
    return f"{num(m, d)} ± {num(s, d)}"


def miles(n):
    return f"{n:,}".replace(",", ".")


def rngn(vals, d=1):
    vals = [v for v in vals if v is not None]
    lo, hi = num(min(vals), d), num(max(vals), d)
    return lo if lo == hi else f"{lo}–{hi}"


def F(name):
    return os.path.join(FIG, name)


# ---------------------------------------------------------------- datos que cita el texto
EV = R["eval"]
SIZES = list(EV)
CONV = R["convergence"]
HIST = R["histories"]
CRIT = R["criterio_mmr"]
TIMING = R["timing"]
SPD = R["speedup_vs_mod2"]
REF2 = R["ref_modulo2"]
REFL = R.get("ref_ils_largo", {})
DIAG = R["diagnostico_modulo2"]
NSEED = R["n_seeds"]
TT = R["eval_time_ms"]
A_, B_ = "A · sesgo MMR", "B · ISAB"
t10 = "10×8"


def gm(label, name):
    return EV[label][name]["mean"]


def best_model(label):
    return min(MODELS, key=lambda n: gm(label, n))


def step_ms(name, n):
    return 1e3 * TIMING[name]["dec"][TIMING[name]["n"].index(n)]


def glast(name, k=10):
    """Gap de validación medio de las últimas k épocas, promedio de las semillas."""
    return float(np.mean([np.mean(h["val_gap"][-k:]) for h in HIST[name]]))


NRUN = len(MODELS) * NSEED
SMALL = [l for l in SIZES if l in REF2]                       # tamaños compartidos con el módulo 2 (hasta 30×20)
SPREAD = max(max(gm(l, n) for n in MODELS) - min(gm(l, n) for n in MODELS) for l in SIZES)
SEC = {n: CONV[n]["sec_per_epoch"] for n in MODELS}
ALPHA_BEST = [h["alpha"][h["best_epoch"] - 1] for h in HIST[A_]]
EP_GAP3 = [e for n in MODELS for e in CONV[n]["epoch_gap_below_3"] if e is not None]
DEC_PER_INST = R["n_train_decisions"] // R["n_train"]         # decisiones por instancia de entrenamiento (= W)
d10 = EV[t10]
k10 = TIMING["Base"]["n"].index(80)


def ep_rng(v):
    v = [x for x in v if x is not None]
    return "no lo alcanza" if not v else (str(min(v)) if min(v) == max(v) else f"{min(v)}–{max(v)}")


# ================================================================ portada y resumen
para([("OII462 · Tarea 6", {"bold": True, "color": BLUE, "size": 10})], space_after=0)
para([("Encoder estático, complejidad y variantes del Transformer para el WTAP", {"bold": True, "color": NAVY, "size": 19})],
     space_after=2)
para([("Informe de los puntos 1, 2, 3, 5 y 6", {"color": MUTED, "size": 11})], space_after=2)
para([(GRUPO, {"bold": True}), ("  ·  Optimización combinatoria  ·  2026", {"color": MUTED})], size=10, space_after=10)

para([b("Resumen. "),
      (f"En el módulo 2, nuestro modelo volvía a procesar toda la instancia en cada decisión, porque "
       f"{DIAG['n_dinamicas']} de sus {DIAG['n_features']} datos de entrada cambiaban después de cada asignación. En esta "
       "tarea lo rediseñamos en dos partes con responsabilidades separadas: un encoder que lee una sola vez los datos fijos "
       "de la instancia (el valor de cada objetivo y las probabilidades de impacto) y un decoder que, en cada paso, mira el "
       "estado actual y elige el siguiente par arma → objetivo. Así, el costo de cada decisión crece en proporción al tamaño "
       "de la instancia y no a su cuadrado. Probamos además dos variantes: A, que le da al modelo una pista de la heurística "
       "clásica MMR, y B, que deja que los pares intercambien información en cada paso sin perder esa eficiencia. En las "
       f"instancias de prueba de 10×8, los tres modelos quedan a {rngn([gm(t10, n) for n in MODELS])} % del experto, contra "
       f"{pct(gm(t10, 'MMR'), 1)} de MMR y {pct(REF2[t10][0], 1)} del modelo del módulo 2, y son parecidos entre sí.", {})])

para([b("Términos usados en este informe")], keep=True, space_after=3)
table([["Término", "Significado"],
       ["Instancia W×T", "Un problema con W armas y T objetivos. Las acciones posibles son los n = W·T pares arma → objetivo."],
       ["Experto ILS", "El algoritmo de búsqueda local de las tareas anteriores. El modelo lo imita y lo usamos como referencia "
                       "(en las instancias pequeñas verificadas encuentra el óptimo)."],
       ["Gap", "Cuánto peor es una solución que la del experto, en porcentaje: 0 % significa igual de buena."],
       ["MMR", "Heurística clásica que, en cada paso, asigna el par que más reduce el valor esperado que sobrevive."],
       ["Encoder y decoder", "Las dos partes del modelo: el encoder lee la instancia y el decoder toma las decisiones una a "
                             "una."],
       ["Época y semilla", "Una época es una pasada por todos los datos de entrenamiento. Cada semilla es una repetición del "
                           "entrenamiento con otro punto de partida aleatorio."]],
      [3.4, TEXT_W - 3.8], size=9, fills=["F3F5F8", None] * 3)

# ================================================================ 1. rediseño
doc.add_heading("1. Rediseño de la arquitectura (punto 1)", level=1)
para([(f"El modelo del módulo 2 recibía una tabla con una fila por cada par arma → objetivo y {DIAG['n_features']} datos "
       f"por fila. Al recorrer una construcción completa comprobamos que {DIAG['n_dinamicas']} de esos datos cambian "
       "después de cada asignación (por ejemplo, si el arma ya fue usada o cuánto valor queda en el objetivo). Por eso todo "
       "el modelo debía recalcularse en cada decisión: no tenía un encoder propiamente tal, y el costo de cada decisión "
       "crecía con el cuadrado del número de pares.", {})])
bullet([b("Encoder (una vez por instancia). "),
        ("Recibe solo lo que no cambia: el valor V", {}), sub("j"),
        (" de cada objetivo y las probabilidades de impacto p", {}), sub("ij"),
        (". Como las armas no tienen datos propios, todas parten del mismo vector aprendido (un token aprendible) y se "
         "diferencian por sus probabilidades, que el modelo usa como conexiones entre armas y objetivos: cada arma se "
         "informa de los objetivos que puede alcanzar, y cada objetivo, de las armas que lo pueden atacar.", {})])
bullet([b("Decoder (en cada decisión). "),
        ("Recibe el estado: qué armas siguen libres y cuánto valor sobrevive en cada objetivo. Combina esa información con "
         "lo que entregó el encoder, da un puntaje a cada par arma libre → objetivo y elige el de mayor puntaje.", {})])
para([("Además, el modelo ya no recibe pistas calculadas con MMR, como el del módulo 2: debe aprenderlas por su cuenta a "
       "partir de los valores, las probabilidades y el estado.", {})])
figure(F("arq_base.png"), TEXT_W * 0.95,
       "Figura 1. Arquitectura base rediseñada. Izquierda: el encoder, que se calcula una vez por instancia solo con los "
       "datos fijos. Derecha: el decoder, que se repite en cada decisión y usa el estado actual. Los círculos naranjos "
       "numerados marcan los cambios respecto del módulo 2 (ver la tabla).")
para([("La tabla resume los cambios, con la numeración de la Figura 1:", {})], keep=True)
table([["#", "Qué cambió", "Módulo 2", "Ahora"],
       ["1", "Qué recibe el modelo", f"Una fila por par arma → objetivo, con {DIAG['n_features']} datos",
        "Armas y objetivos por separado; las probabilidades, como conexiones entre ellos"],
       ["2", "Encoder", "Se recalculaba en cada decisión", "Se calcula una sola vez por instancia"],
       ["3", "Estado de la solución", "Mezclado con los datos de entrada, junto con pistas de MMR",
        "Solo en el decoder: armas libres y valor que sobrevive en cada objetivo"],
       ["4", "Resumen del estado", "Una fila global y el promedio de las filas disponibles",
        "Promedios de las armas libres y de los objetivos, refinados con atención"],
       ["5", "Puntaje de cada par", "Comparaba cada fila con el resumen",
        "Combina la información del arma, del objetivo y de su probabilidad de impacto"],
       ["", "Entrenamiento", f"{miles(R['n_train_decisions'])} decisiones sueltas",
        "Instancias completas: una lectura del encoder y W decisiones"]],
      [0.6, 3.4, 5.9, 7.5], size=9)

# ================================================================ 2. complejidad
doc.add_heading("2. Complejidad temporal (punto 2)", level=1)
para([("Llamamos n = W·T al número de pares arma → objetivo. Un costo es ", {}), b("lineal"),
      (" si crece en proporción a n (una instancia el doble de grande toma el doble de tiempo) y ", {}), b("cuadrático"),
      (" si crece con n² (el doble de grande toma cuatro veces más). El decoder se ejecuta W veces por instancia, así que su "
       "costo es el que más importa; el enunciado pide que sea a lo más lineal.", {})])
table([["Modelo", "Encoder (una vez por instancia)", "Decoder (en cada decisión)"],
       ["Módulo 2", "No existe por separado: todo se recalcula en cada paso",
        "Cuadrático, O(n²): cada par se compara con todos los demás"],
       ["Base y variante A", "Lineal, O(n): cada arma mira los T objetivos y cada objetivo las W armas",
        "Lineal, O(n): se revisa cada par una vez"],
       ["Variante B (con ISAB)", "Lineal, O(n)", "Lineal, O(n·m): los pares se comunican a través de m = 16 resúmenes"],
       ["Variante B sin ISAB (solo para comparar)", "Lineal, O(n)", "Cuadrático, O(n²)"]],
      [4.4, 6.2, 6.8], size=9, fills=[FILL["Módulo 2"], FILL["Base"], FILL[B_], "F9DEDC"])
para([("Un decoder lineal es lo mejor posible, porque en cada decisión hay que dar un puntaje a cada uno de los n pares. El "
       "encoder también es lineal, aunque la tarea permitía que fuera cuadrático. La única parte que podría volverse "
       "cuadrática es la comunicación entre pares de la variante B, y por eso usa ISAB (sección 3.2).", {})])
figure(crop(F("complejidad.png"), "complejidad_paso.png", (0.0, 0.0, 0.328, 1.0)), 9.5,
       "Figura 2. Tiempo de una decisión según el número de pares n, medido en un hilo de CPU (ambos ejes en escala "
       "logarítmica). Las curvas más empinadas, del módulo 2 y de B sin ISAB («B con atención completa» en la leyenda), "
       "crecen como n²; las de los modelos nuevos, "
       "como n.")
para([(f"Las mediciones confirman el análisis: en una instancia de 50×40, una decisión pasa de "
       f"{num(step_ms('Módulo 2', 2000), 1)} ms en el módulo 2 a {num(step_ms('Base', 2000), 2)} ms en la base, y resolver "
       f"la instancia completa es {num(SPD['50×40'], 0)} veces más rápido. El encoder cuesta "
       f"{num(1e3 * TIMING['Base']['enc'][k10], 2)} ms en 10×8 y se paga una sola vez.", {})])

# ================================================================ 3. variantes
doc.add_heading("3. Variantes (punto 3)", level=1)
doc.add_heading("3.1 Variante A: una pista de MMR (sesgo inductivo)", level=2)
para([("Un sesgo inductivo es una pista que le damos al modelo para que no parta de cero. Usamos la de MMR: cuánto valor "
       "esperado elimina asignar el arma i al objetivo j en el estado actual, g", {}), sub("ij"), (" = V", {}), sub("j"),
      (" · q", {}), sub("j"), (" · p", {}), sub("ij"), (", donde q", {}), sub("j"),
      (" es la probabilidad de que el objetivo j siga en pie. La pista se suma en dos lugares, como en el material del "
       "curso:", {})], keep=True)
bullet([b("En el puntaje de cada par (decoder): "),
        ("los pares con mayor ganancia de MMR reciben más puntaje. Al comienzo el modelo elige exactamente lo mismo que "
         "MMR, y durante el entrenamiento aprende cuándo conviene apartarse.", {})])
bullet([b("En la atención del encoder: "),
        ("cada arma presta más atención a los objetivos donde su ganancia inicial es mayor.", {})])
para([("En ambos casos la pista se multiplica por pesos α que el modelo aprende, así que él decide cuánto confiar en ella.",
       {})])
para([b("¿Por qué MMR es un criterio robusto?")], keep=True, space_after=2)
cob = CRIT["cobertura"]
bullet(["Es la heurística clásica del WTAP (den Broeder et al., 1959) y es óptima cuando todas las armas son iguales."])
bullet(["Es el punto de partida de métodos del estado del arte (Ahuja et al., 2007) y de nuestro propio experto."])
bullet([(f"En nuestros datos, el experto elige la misma acción que MMR en el {pct(100 * cob['1'], 0)} de las decisiones, "
         f"y esa acción está entre las 3 mejores de MMR en el {pct(100 * cob['3'], 0)} de los casos.", {})])

figure(F("arq_variante_A.png"), TEXT_W * 0.78,
       "Figura 3. Variante A: la base más la pista de MMR (en morado), en el encoder y en el puntaje final.")
doc.add_heading("3.2 Variante B: los pares se comunican (ISAB)", level=2)
para([("En la base, cada par se evalúa por separado: un par no \"sabe\" si otra arma libre aprovecharía mejor su objetivo. "
       "En la variante B, en cada decisión los pares intercambian información. Hacerlo comparando cada par con todos los "
       "demás sería cuadrático, así que usamos un bloque ISAB (Lee et al., 2019): 16 puntos aprendidos resumen el conjunto "
       "de pares y luego cada par lee esos resúmenes. Así la comunicación cuesta O(n·m) en vez de O(n²).", {})])
para([b("¿Por qué ISAB? "),
      ("Lo elegimos frente a LUNA (Ma et al., 2021) porque el estado es un conjunto sin orden que cambia en cada paso, sin "
       "una memoria que deba conservarse, que es lo que aporta LUNA. Tampoco usamos atención dispersa, porque necesitaría "
       "una regla para decidir qué pares conectar, y ese papel ya lo cumple la pista de MMR de la variante A. Así, cada "
       "variante prueba una sola idea.", {})])

figure(F("arq_variante_B.png"), TEXT_W * 0.78,
       "Figura 4. Variante B: el resumen del estado de la base se reemplaza por un bloque ISAB sobre los pares disponibles "
       "(en verde).")
# ================================================================ 4. entrenamiento
doc.add_heading("4. Entrenamiento (punto 4)", level=1)
cfg = R["train_cfg"]
para([(f"Los tres modelos se entrenaron exactamente igual: con los mismos datos ({miles(R['n_train'])} instancias de "
       f"10×8 resueltas por el experto y {miles(R['n_val_decisions'] // DEC_PER_INST)} de validación), el mismo optimizador "
       f"y la misma tasa de aprendizaje, hasta {cfg['max_epochs']} épocas con parada temprana (el entrenamiento se detiene "
       f"si deja de mejorar) y {NSEED} semillas por modelo. La lista completa de hiperparámetros está en el notebook.",
       {})])

# ================================================================ 5. convergencia
doc.add_heading("5. Convergencia (punto 5)", level=1)
para([("La Figura 5 muestra las curvas de entrenamiento y la tabla resume la velocidad de cada modelo:", {})], keep=True)
rows = [["Modelo", "Gap antes de entrenar", "Épocas hasta gap\u00a0<\u00a03\u00a0%", "Mejor pérdida de validación", "Segundos por época"]]
for n in MODELS:
    cv = CONV[n]
    rows.append([n, pct(cv["gap_init"], 0 if cv["gap_init"] > 100 else 1), ep_rng(cv["epoch_gap_below_3"]),
                 pm(cv["best_val_loss_mean"], cv["best_val_loss_std"], 3), num(cv["sec_per_epoch"], 1)])
table(rows, [3.4, 3.2, 3.4, 4.2, 3.0], size=9, fills=[FILL[n] for n in MODELS],
      aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 4)
note(f"Épocas: rango entre las {NSEED} semillas. Pérdida: promedio ± desviación entre semillas (menor es mejor).")
cb, ca, cB = CONV["Base"], CONV[A_], CONV[B_]
bullet([b("Velocidad, en épocas. "),
        (f"A parte con ventaja: como arranca imitando a MMR, su gap antes de entrenar es {pct(ca['gap_init'], 1)}, contra "
         f"{pct(cb['gap_init'], 0)} de los otros dos (sin entrenar, mandan todas las armas al mismo objetivo). Pero la "
         "ventaja dura poco: después de la primera época los tres ya superan a MMR y llegan a un gap menor a 3 % en "
         f"{min(EP_GAP3)} a {max(EP_GAP3)} épocas.", {})])
bullet([b("Velocidad, en tiempo real. "),
        (f"B es la que más rápido baja la pérdida y la que termina más abajo, pero cada época le toma "
         f"{num(SEC[B_] / SEC['Base'], 1)} veces más ({num(SEC[B_], 1)} s contra {rngn([SEC['Base'], SEC[A_]])} s): medida "
         "en minutos, es la más lenta.", {})])
bullet([b("Estabilidad. "),
        (f"Las {NRUN} corridas ({len(MODELS)} modelos × {NSEED} semillas) convergieron sin saltos ni divergencias, y las "
         "semillas de un mismo modelo terminan muy cerca. Al final, los tres llegan a un gap de validación parecido "
         f"({rngn([glast(n) for n in MODELS], 2)} % en las últimas 10 épocas): la menor pérdida de B no se tradujo en "
         "mejores soluciones.", {})])
bullet([b("La pista de MMR se usa. "),
        (f"El peso α con que la variante A suma la pista de MMR subió de 1 a {rngn(ALPHA_BEST)} durante el entrenamiento: "
         "el modelo aprendió a confiar más en ella.", {})])

figure(F("convergencia.png"), TEXT_W * 0.88,
       f"Figura 5. Curvas de entrenamiento por época (línea = promedio de {NSEED} semillas; banda = mínimo y máximo). Abajo "
       "a la derecha, el gap de las soluciones que construye cada modelo en 100 instancias de validación; la línea naranja "
       "es MMR.")
# ================================================================ 6. evaluación
doc.add_heading("6. Evaluación con instancias de distintos tamaños (punto 6)", level=1)
para([(f"Los modelos se entrenaron solo con instancias de 10×8 y los probamos, sin cambios, en {len(SIZES)} tamaños, de "
       f"{SIZES[0]} a {SIZES[-1]}. Usamos las mismas instancias que el proyecto del módulo 2 (30 por tamaño; las 50 de "
       "prueba en 10×8), así que la columna «Módulo 2» es directamente comparable. Cada valor es el gap respecto del "
       f"experto: promedio ± desviación estándar sobre las {NSEED} semillas y todas las instancias.", {})], keep=True)
rows = [["Tamaño"] + MODELS + ["MMR", "Módulo 2"]]
for l in SIZES:
    best = best_model(l)
    rows.append([l] + [[(pm(gm(l, n), EV[l][n]["std"]), {"bold": n == best})] for n in MODELS] +
                [pm(gm(l, "MMR"), EV[l]["MMR"]["std"]), pm(*REF2[l]) if l in REF2 else "—"])
table(rows, [2.2, 3.0, 3.0, 3.0, 3.0, 3.0], size=9, aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 5)
rand = [gm(l, "Aleatorio") for l in SIZES]
note(f"Gap (%) respecto del experto ILS; menor es mejor. En negrita, el mejor de los tres modelos en cada tamaño. Como "
     f"referencia, asignar las armas al azar da gaps de {num(min(rand), 0)} a {num(max(rand), 0)} %. La Figura 6, al "
     "final de la sección, muestra los mismos datos en barras.")
models_small = [gm(l, n) for l in SMALL for n in MODELS]
mmr_small = [gm(l, "MMR") for l in SMALL]
t_models10 = [TT[t10][n] for n in MODELS]
speed10 = [TT[t10]["Experto ILS"] / TT[t10][n] for n in MODELS]
bullet([b("En el tamaño de entrenamiento (10×8), "),
        (f"los tres modelos quedan entre {pct(min(gm(t10, n) for n in MODELS))} y {pct(max(gm(t10, n) for n in MODELS))}, "
         f"muy por debajo de MMR ({pct(gm(t10, 'MMR'))}) y mejor que el modelo del módulo 2 ({pct(REF2[t10][0])}).", {})])
bullet([b("En tamaños que no vieron al entrenar, "),
        (f"siguen funcionando bien: hasta {SMALL[-1]} se mantienen entre {pct(min(models_small), 1)} y "
         f"{pct(max(models_small), 1)} (MMR: {rngn(mmr_small)} %) y mejoran al módulo 2 en todos los tamaños "
         "comparables.", {})])
bullet([b("En las instancias más grandes, "),
        (f"MMR se vuelve muy bueno: en 50×40 los modelos todavía le ganan, pero en 100×80 MMR "
         f"({pct(gm('100×80', 'MMR'))}) queda levemente mejor en promedio. Ahí el propio experto ya no es óptimo: una "
         f"búsqueda 20 veces más larga lo mejora en {rngn([d['mejora_ils_largo'] for d in REFL.values()])} %.", {})])
bullet([b("Entre las tres arquitecturas "),
        (f"las diferencias son pequeñas (a lo más {num(SPREAD, 1)} puntos) y del mismo tamaño que las diferencias entre "
         "semillas, así que ninguna gana claramente. Lo que distingue a B es que encuentra la misma solución que el "
         f"experto con más frecuencia: en el {pct(d10[B_]['pct_optimal'], 0)} de las ejecuciones de 10×8, contra "
         f"{rngn([d10['Base']['pct_optimal'], d10[A_]['pct_optimal']], 0)} % de las otras dos.", {})])
bullet([b("En tiempo, "),
        ((f"en 10×8 cada modelo resuelve una instancia en menos de 1 ms" if max(t_models10) < 1 else
          f"en 10×8 cada modelo resuelve una instancia en {rngn(t_models10)} ms") +
         f", entre {num(min(speed10), 0)} y {num(max(speed10), 0)} veces más rápido que el experto. En 100×80, en cambio, "
         "son más lentos que él, porque el modelo toma W decisiones y cada una revisa todos los pares.", {})])

figure(F("tamanos.png"), TEXT_W,
       "Figura 6. Gap por tamaño: la barra es el promedio, la línea vertical ±1 desviación estándar y los puntos negros, el "
       "promedio de cada semilla.")
# ================================================================ 7. conclusiones
doc.add_heading("7. Conclusiones", level=1)
bullet([b("El rediseño funciona. "),
        ("Separar un encoder que lee la instancia una vez de un decoder que decide paso a paso dio un modelo más preciso que "
         f"el del módulo 2 en todos los tamaños comparables y {num(SPD['50×40'], 0)} veces más rápido en 50×40.", {})])
bullet([b("El costo es lineal. "),
        ("Cada decisión crece en proporción al número de pares, que es lo mínimo posible; la variante B mantiene esa "
         "propiedad gracias a ISAB.", {})])
bullet([b("La pista de MMR (A) ayuda al comienzo, "),
        ("porque el modelo parte desde una buena política, pero no mejora el resultado final: la base aprende lo mismo en "
         "una época.", {})])
bullet([b("La comunicación entre pares (B) "),
        ("logra más soluciones iguales a las del experto, pero no un gap promedio menor, y entrenarla cuesta "
         f"{num(SEC[B_] / SEC['Base'], 0)} veces más. Con estos resultados, la base es la opción con mejor relación entre "
         "calidad y costo.", {})])

doc.add_heading("Referencias", level=1)
refs = ["Ahuja, R. K., Kumar, A., Jha, K. C. y Orlin, J. B. (2007). Exact and heuristic algorithms for the weapon-target "
        "assignment problem. Operations Research, 55(6), 1136–1146.",
        "den Broeder, G. G., Ellison, R. E. y Emerling, L. (1959). On optimum target assignments. Operations Research, 7(3), "
        "322–326.",
        "Lee, J., Lee, Y., Kim, J., Kosiorek, A., Choi, S. y Teh, Y. W. (2019). Set Transformer: a framework for "
        "attention-based permutation-invariant neural networks. ICML.",
        "Ma, X., Kong, X., Wang, S., Zhou, C., May, J., Ma, H. y Zettlemoyer, L. (2021). Luna: Linear unified nested "
        "attention. NeurIPS.",
        "Material del curso (TSP-Framework): notebooks de sesgo inductivo, atención lineal (LUNA, ISAB) y esparsificación."]
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
