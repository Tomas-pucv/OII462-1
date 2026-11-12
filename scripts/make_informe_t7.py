"""Genera el informe de la Tarea 7 (puntos 1, 2, 4, 5 y 6) a partir de figuras_t7/ (figuras y resultados.json).

Versión en lenguaje simple: el detalle técnico (código, fórmulas y todas las mediciones) está en el notebook
Tarea7_WTAP_egat.ipynb. Todas las cifras del texto se leen de resultados.json.

Uso (desde la raíz del repositorio):
    python scripts/make_diagramas_t7.py
    python scripts/make_informe_t7.py
    soffice --headless --convert-to pdf Tarea7_WTAP_informe.docx
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
FIG = os.environ.get("T7_FIG", os.path.join(ROOT, "figuras_t7"))
OUT = os.environ.get("T7_OUT", os.path.join(ROOT, "Tarea7_WTAP_informe.docx"))
R = json.load(open(os.path.join(FIG, "resultados.json"), encoding="utf-8"))
TMP = os.path.join(FIG, "_tmp_informe")
os.makedirs(TMP, exist_ok=True)

FONT = "Arial"
NAVY = RGBColor(0x1B, 0x2A, 0x41)
BLUE = RGBColor(0x1F, 0x77, 0xB4)
MUTED = RGBColor(0x6B, 0x75, 0x85)
INK = RGBColor(0x2E, 0x35, 0x40)
MODELS = ["T6 · Base", "EGAT · 1 capa", "EGAT · 2 capas", "EGAT · 3 capas"]
FILL = {"T6 · Base": "DCEAF7", "EGAT · 1 capa": "EEF8EA", "EGAT · 2 capas": "DDF1D6", "EGAT · 3 capas": "C9E8BF",
        "MMR": "FDEBD9", "Módulo 2": "E8EEF6"}

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
NSEED = R["n_seeds"]
TT = R["eval_time_ms"]
REF2 = R["ref_modulo2"]
REF6 = R["ref_t6"]
REFL = R.get("ref_ils_largo", {})
ABL = R["ablacion"]
GAT = R["gat_check"]
CAMPO = R["campo_receptivo"]
PARE = R["params_encoder"]
T6, E1, E2, E3 = MODELS
EGATS = MODELS[1:]
t10 = "10×8"


def gm(label, name):
    return EV[label][name]["mean"]


def best_model(label, names=MODELS):
    return min(names, key=lambda n: gm(label, n))


def avg_sizes(name):
    """Gap medio de un modelo en todos los tamaños (promedio de las medias por tamaño)."""
    return float(np.mean([gm(l, name) for l in SIZES]))


def ep_rng(v):
    v = [x for x in v if x is not None]
    return "no lo alcanza" if not v else (str(min(v)) if min(v) == max(v) else f"{min(v)}–{max(v)}")


def short(name):
    return {T6: "T6 · Base", E1: "EGAT-1", E2: "EGAT-2", E3: "EGAT-3"}.get(name, name)


# Tarea 6: ¿por qué la Base? (mismas instancias; cifras guardadas en figuras_t6)
T6_AVG = {k: float(np.mean([REF6[l][k][0] for l in REF6])) for k in ("Base", "A", "B")}
T6_BASE_WINS = sum(min(("Base", "A", "B"), key=lambda k: REF6[l][k][0]) == "Base" for l in REF6)
NRUN = len(MODELS) * NSEED
SEC = {n: CONV[n]["sec_per_epoch"] for n in MODELS}

# ================================================================ portada y resumen
para([("OII462 · Tarea 7", {"bold": True, "color": BLUE, "size": 10})], space_after=0)
para([("El WTAP como grafo: encoder EGAT y profundidad del encoder", {"bold": True, "color": NAVY, "size": 19})],
     space_after=2)
para([("Informe de los puntos 1, 2, 4, 5 y 6", {"color": MUTED, "size": 11})], space_after=2)
para([(GRUPO, {"bold": True}), ("  ·  Optimización combinatoria  ·  2026", {"color": MUTED})], size=10, space_after=10)

para([b("Resumen. "),
      ("Representamos el WTAP como un grafo bipartito: las armas y los objetivos son nodos, y cada par arma–objetivo es "
       "una arista con su probabilidad de impacto. Sobre ese grafo reemplazamos el encoder de nuestro mejor modelo de la "
       "Tarea 6 (la Base) por capas EGAT, que usan la arista tanto para decidir a qué vecino prestar atención como en el "
       "mensaje que reciben, y dejamos el decoder igual. Entrenamos con la misma configuración la EGAT con 1, 2 y 3 capas "
       f"({NSEED} semillas cada una) y la comparamos con la Base re-entrenada. ", {}),
      ("RESUMEN_RESULTADOS", {})])

para([b("Términos usados en este informe")], keep=True, space_after=3)
table([["Término", "Significado"],
       ["Grafo bipartito", "Un grafo con dos tipos de nodos en el que las aristas solo unen nodos de tipos distintos; "
                           "aquí, armas y objetivos."],
       ["Arista y atributo", "La conexión entre un arma y un objetivo. Su atributo es la probabilidad de impacto p_ij."],
       ["Salto (hop)", "Pasar de un nodo a un vecino por una arista. Cada capa del encoder lleva la información un salto "
                       "más lejos."],
       ["EGAT", "Red atencional de grafos con aristas: cada nodo combina la información de sus vecinos, ponderada por una "
                "atención que también mira la arista que los une."],
       ["Gap", "Cuánto peor es una solución que la del experto, en porcentaje: 0 % significa igual de buena."],
       ["Experto ILS y MMR", "El algoritmo de búsqueda local que el modelo imita y que usamos como referencia, y la "
                             "heurística clásica que en cada paso asigna el par que más reduce el valor que sobrevive."],
       ["Época y semilla", "Una época es una pasada por todos los datos de entrenamiento. Cada semilla es una repetición "
                           "del entrenamiento con otro punto de partida aleatorio."]],
      [3.4, TEXT_W - 3.8], size=9, fills=["F3F5F8", None] * 3 + ["F3F5F8"])

# ================================================================ 1. el grafo
doc.add_heading("1. El WTAP como grafo (punto 1)", level=1)
para([("En el WTAP hay W armas y T objetivos: el objetivo j vale V", {}), sub("j"),
      (" y el arma i lo destruye con probabilidad p", {}), sub("ij"),
      (". El problema no trae un grafo explícito, pero su estructura es naturalmente bipartita: hay dos tipos de "
       "entidades, y toda la información que las relaciona está en las probabilidades. Lo representamos como un grafo "
       "G = (V, E):", {})], keep=True)
table([["Elemento", "Qué representa", "Atributos"],
       ["Nodo arma (W nodos)", "Un arma que hay que asignar",
        "Ninguno propio: todas parten del mismo vector aprendido. En el decoder se agrega si está libre o ya asignada."],
       ["Nodo objetivo (T nodos)", "Un objetivo enemigo",
        "Su valor normalizado, V_j / max V. En el decoder se agrega la probabilidad q_j de que siga en pie."],
       ["Arista (i, j) (W·T aristas)", "El arma i puede atacar al objetivo j (p_ij > 0)",
        "La probabilidad de impacto p_ij, llevada a una escala fija entre −1 y 1."],
       ["Auto-bucle (uno por nodo)", "El nodo se considera a sí mismo al combinar información", "Un atributo aprendido."]],
      [3.6, 5.0, 8.6], size=9, fills=["DCEAF7", "FDEBD9", "F3F5F8", None])
para([("Como en nuestras instancias todas las armas pueden atacar a todos los objetivos (p", {}), sub("ij"),
      (" está entre 0,6 y 0,9), el grafo es ", {}), b("bipartito completo"),
      (": no hay aristas entre armas ni entre objetivos, y las aristas coinciden con las acciones del decoder (elegir un "
       "par arma → objetivo). El grafo es ", {}), b("no dirigido"),
      (", porque la relación es la misma vista desde el arma o desde el objetivo, pero la información viaja en los dos "
       "sentidos: de los objetivos a las armas y de las armas a los objetivos, con parámetros distintos para cada "
       "sentido.", {})])
figure(F("grafo_wtap.png"), TEXT_W,
       "Figura 1. Izquierda: una instancia de 3 armas y 2 objetivos como grafo; el número en cada arista es p_ij y el "
       "valor de cada objetivo aparece a su lado. Derecha: cuántos saltos (capas) necesita la información de un arma para "
       "llegar a cada nodo.")
para([b("¿Por qué esta representación? "),
      ("Es la misma idea del ejemplo del MKP del material (Gasse et al., 2019): no depende de cuántas armas u objetivos "
       "haya ni de su orden, así que el mismo modelo sirve para instancias de cualquier tamaño. Además, desde un arma se "
       "llega a todos los objetivos en 1 salto y a todas las demás armas en 2 saltos (a través de los objetivos que "
       f"comparten): el diámetro del grafo es {R['grafo']['diametro']}. Esto será clave para la ablación de profundidad.",
       {})])

# ================================================================ 2. encoder EGAT
doc.add_heading("2. Encoder EGAT (punto 2)", level=1)
para([b("¿GAT o EGAT? EGAT. "),
      ("Las armas no tienen datos propios: lo único que las distingue son sus probabilidades, que están en las aristas. "
       "Una GAT común solo mira los nodos, así que todas las armas quedarían iguales. Lo comprobamos con un encoder de 3 "
       "capas sin entrenar: con GAT, la distancia entre los vectores de dos armas cualesquiera es exactamente "
       f"{num(GAT['gat']['armas'], 0)}; con EGAT llega a {num(GAT['egat']['armas'])}. Por eso el material indica que en "
       "problemas como el WTAP la EGAT es obligatoria. Para no perder la invariancia al orden de las armas, todas parten "
       "del mismo vector aprendido, y no de uno distinto por índice.", {})])
para([("En cada capa, cada arma recibe información de todos los objetivos y cada objetivo de todas las armas, además "
       "de sí mismo. Una capa EGAT hace cinco cosas:", {})], keep=True)
bullet(["Transforma los vectores de los nodos y convierte la probabilidad de cada arista en un vector."])
bullet(["Calcula un puntaje para cada vecino sumando tres miradas: la del nodo que recibe, la del vecino y la de la arista "
        "que los une (atención aditiva, en vez del producto punto de la Tarea 6)."])
bullet(["Normaliza los puntajes de los vecinos para que sumen 1 (softmax)."])
bullet(["Arma un mensaje con cada vecino y su arista, y los suma ponderados por esos pesos, con 4 cabezas en paralelo."])
bullet(["Conserva la información anterior del nodo (conexión residual) y la procesa con una pequeña red, igual que en la "
        "Tarea 6."])
para(["Así, con L capas cada nodo reúne información que está a L saltos de distancia."])
figure(F("arq_egat.png"), TEXT_W * 0.95,
       "Figura 2. Arquitectura completa. El encoder EGAT (izquierda) reemplaza al de la Tarea 6; el decoder (derecha) es "
       "el de la Base de la Tarea 6, sin cambios. Los círculos numerados marcan los cambios (ver la tabla).")
para([("La tabla resume los cambios respecto de la Tarea 6, con la numeración de la Figura 2:", {})], keep=True)
table([["#", "Qué cambió", "Tarea 6 (Base)", "Tarea 7 (EGAT)"],
       ["1", "Grafo", "Implícito: las armas atienden a los objetivos y viceversa",
        "Explícito: grafo bipartito con aristas p_ij y auto-bucles"],
       ["2", "Puntaje de atención", "Producto punto entre los vectores, más un término por p_ij",
        "Atención aditiva: suma de las miradas del nodo, del vecino y de la arista"],
       ["3", "Transformaciones", "Tres matrices: consulta, clave y valor", "Una matriz por tipo de nodo, como en GAT"],
       ["4", "Arista", "Un número por cabeza en el puntaje y un vector en el mensaje",
        "Un vector por arista, en el puntaje y en el mensaje"],
       ["5", "Profundidad", "2 capas", "1, 2 o 3 capas (ablación)"],
       ["", "Decoder", "—", "Sin cambios"]],
      [0.6, 3.2, 6.4, 7.2], size=9)
para([b("¿Por qué el decoder de la Base? "),
      (f"Entre los modelos de la Tarea 6, la Base tuvo el menor gap promedio en los {len(REF6)} tamaños "
       f"({pct(T6_AVG['Base'])}, contra {pct(T6_AVG['A'])} de la variante A y {pct(T6_AVG['B'])} de la B), fue la mejor en "
       f"{T6_BASE_WINS} de los {len(REF6)} tamaños y es la más barata de entrenar. Mantener su decoder hace que la "
       "comparación mida solo el efecto del encoder.", {})])

# ================================================================ 3. entrenamiento
doc.add_heading("3. Entrenamiento", level=1)
cfg = R["train_cfg"]
para([(f"Entrenamos cuatro modelos con exactamente la configuración de la Base de la Tarea 6: los mismos datos "
       f"({miles(R['n_train'])} instancias de 10×8 resueltas por el experto y "
       f"{miles(R['n_val_decisions'] // (R['n_train_decisions'] // R['n_train']))} de validación), el mismo optimizador y "
       f"la misma tasa de aprendizaje, hasta {cfg['max_epochs']} épocas con parada temprana y {NSEED} semillas por modelo. "
       "Los modelos son la Base de la Tarea 6, re-entrenada aquí como referencia, y la EGAT con 1, 2 y 3 capas. " +
       ("La Base re-entrenada reproduce exactamente los resultados de la Tarea 6, lo que confirma que el código y los "
        "datos son los mismos." if R["reproduce_t6"] else ""), {})])

# ================================================================ 4. ablación de profundidad
doc.add_heading("4. Ablación de profundidad (punto 4)", level=1)
para([("Cada capa EGAT lleva la información un salto más lejos. En nuestro grafo, esto significa:", {})], keep=True)
table([["Capas", "Lo que sabe un arma al salir del encoder", "Lo que sabe un objetivo"],
       ["1", "Los objetivos que puede atacar: sus valores y sus propias probabilidades",
        "Las armas que lo pueden atacar, a través de sus probabilidades"],
       ["2", "Además, las otras armas que compiten por esos mismos objetivos",
        "Además, los otros objetivos que esas armas pueden atacar"],
       ["3", "Nada nuevo que alcanzar (el diámetro es 2): solo combina una vez más la misma información", "Ídem"]],
      [1.4, 8.0, 7.8], size=9)
c1, c2 = CAMPO["1"], CAMPO["2"]
para([("Lo verificamos midiendo de qué datos depende el vector de un arma: con 1 capa, su derivada respecto de las "
       "probabilidades de las otras armas es exactamente 0 (no las ve); con 2 capas ya no lo es "
       f"({num(c2['arma_otras_armas'], 3)}). Además, el decoder mezcla información de todas las armas y objetivos en cada "
       "paso, así que incluso con 1 capa el modelo completo puede relacionar armas entre sí: lo que cambia es si esa "
       "relación ya viene en los vectores del encoder.", {})])
rows = [["Modelo", "Capas", "Parámetros del encoder", "Segundos por época", "Gap de validación",
         "Gap de prueba 10×8", f"Gap promedio en los {len(SIZES)} tamaños"]]
for n in MODELS:
    a = ABL[n]
    rows.append([short(n), str(a["capas"]) + (" (T6)" if n == T6 else ""), miles(a["params_encoder"]),
                 num(a["sec_per_epoch"], 1), pm(np.mean(a["best_val_gap"]), np.std(a["best_val_gap"])),
                 pm(np.mean(a["test10_seed_means"]), np.std(a["test10_seed_means"])),
                 pm(np.mean(a["avg_sizes_seed_means"]), np.std(a["avg_sizes_seed_means"]))])
table(rows, [2.4, 1.5, 2.3, 2.0, 2.9, 2.9, 3.2], size=8.8, fills=[FILL[n] for n in MODELS],
      aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 6)
note(f"Gaps en %: promedio ± desviación entre las {NSEED} semillas. Validación: mejor época de cada semilla; prueba: "
     "las 50 instancias 10×8; promedio en los tamaños: media de los 7 tamaños de la sección 6.")
figure(F("profundidad.png"), TEXT_W,
       "Figura 3. Gap según el número de capas EGAT (puntos = semillas; línea = promedio). La línea azul punteada y la "
       "banda muestran la Base de la Tarea 6 (promedio y rango entre semillas).")
for txt in ("ANALISIS_ABLACION",):
    para([(txt, {})])

# ================================================================ 5. convergencia
doc.add_heading("5. Convergencia (punto 5)", level=1)
para([("La Figura 4 muestra las curvas de entrenamiento de los cuatro modelos y la tabla resume su velocidad:", {})],
     keep=True)
rows = [["Modelo", "Gap antes de entrenar", "Épocas hasta gap < 3 %", "Mejor pérdida de validación",
         "Segundos por época"]]
for n in MODELS:
    cv = CONV[n]
    rows.append([short(n), pct(cv["gap_init"], 0 if cv["gap_init"] > 100 else 1), ep_rng(cv["epoch_gap_below_3"]),
                 pm(cv["best_val_loss_mean"], cv["best_val_loss_std"], 3), num(cv["sec_per_epoch"], 1)])
table(rows, [3.4, 3.2, 3.4, 4.2, 3.0], size=9, fills=[FILL[n] for n in MODELS],
      aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 4)
note(f"Épocas: rango entre las {NSEED} semillas. Pérdida: promedio ± desviación entre semillas (menor es mejor).")
for txt in ("ANALISIS_CONVERGENCIA",):
    para([(txt, {})])
figure(F("convergencia.png"), TEXT_W * 0.88,
       f"Figura 4. Curvas de entrenamiento por época (línea = promedio de {NSEED} semillas; banda = mínimo y máximo). Abajo "
       "a la derecha, el gap de las soluciones que construye cada modelo en 100 instancias de validación; la línea naranja "
       "es MMR.")

# ================================================================ 6. evaluación
doc.add_heading("6. Evaluación con instancias de distintos tamaños (punto 6)", level=1)
para([(f"Los modelos se entrenaron solo con instancias de 10×8 y los probamos, sin cambios, en {len(SIZES)} tamaños, de "
       f"{SIZES[0]} a {SIZES[-1]}, con las mismas instancias de la Tarea 6 y del proyecto del módulo 2 (30 por tamaño; "
       "las 50 de prueba en 10×8). Cada valor es el gap respecto del experto: promedio ± desviación estándar sobre las "
       f"{NSEED} semillas y todas las instancias.", {})], keep=True)
rows = [["Tamaño"] + [short(n) for n in MODELS] + ["MMR", "Módulo 2"]]
for l in SIZES:
    best = best_model(l)
    rows.append([l] + [[(pm(gm(l, n), EV[l][n]["std"]), {"bold": n == best})] for n in MODELS] +
                [pm(gm(l, "MMR"), EV[l]["MMR"]["std"]), pm(*REF2[l]) if l in REF2 else "—"])
table(rows, [1.9, 2.6, 2.6, 2.6, 2.6, 2.6, 2.6], size=8.8,
      aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 6)
rand = [gm(l, "Aleatorio") for l in SIZES]
note(f"Gap (%) respecto del experto ILS; menor es mejor. En negrita, el mejor de los cuatro modelos en cada tamaño. Como "
     f"referencia, asignar las armas al azar da gaps de {num(min(rand), 0)} a {num(max(rand), 0)} %. La Figura 5, al "
     "final de la sección, muestra los mismos datos en barras.")
for txt in ("ANALISIS_EVALUACION",):
    para([(txt, {})])
figure(F("tamanos.png"), TEXT_W,
       "Figura 5. Gap por tamaño: la barra es el promedio, la línea vertical ±1 desviación estándar y los puntos negros, el "
       "promedio de cada semilla.")

# ================================================================ 7. conclusiones
doc.add_heading("7. Conclusiones", level=1)
for txt in ("CONCLUSIONES",):
    para([(txt, {})])

doc.add_heading("Referencias", level=1)
refs = ["Gasse, M., Chételat, D., Ferroni, N., Charlin, L. y Lodi, A. (2019). Exact combinatorial optimization with "
        "graph convolutional neural networks. NeurIPS.",
        "Veličković, P., Cucurull, G., Casanova, A., Romero, A., Liò, P. y Bengio, Y. (2018). Graph attention networks. "
        "ICLR.",
        "Material del curso (TSP-Framework): redes atencionales de grafos (GAT y EGAT)."]
for r_ in refs:
    p = para(r_, size=9, space_after=3)
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.6)

# pie de página con número
for s_ in doc.sections:
    fp = s_.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ = fp.add_run("OII462 · Tarea 7 · WTAP  —  ")
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
