"""Diagramas de arquitectura de la Tarea 6 (base rediseñada y variantes A y B) en figuras_t6/.

Uso (desde la raíz del repositorio):
    python scripts/make_diagramas_t6.py
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG = os.path.join(ROOT, "figuras_t6")
os.makedirs(FIG, exist_ok=True)

NAVY, MUTED = "#1B2A41", "#6B7585"
BLUE, BLUE_D, BLUE_L, BLUE_P = "#1F77B4", "#174A7C", "#DCEAF7", "#F2F7FC"
ORANGE, ORANGE_D, ORANGE_P = "#E8710A", "#8A4B0F", "#FFF8F1"
PURPLE, PURPLE_L = "#7B3FB0", "#EFE4F8"
GREEN, GREEN_L = "#23893A", "#DFF2E3"
RED, MASK = "#C0392B", "#2CA02C"
CHG = "#D35400"                      # marcadores de "cambio respecto del módulo 2"

plt.rcParams.update({"font.family": "DejaVu Sans", "mathtext.fontset": "dejavusans"})


def rbox(ax, x, y, w, h, fc, ec, lw=1.4, r=0.12, z=2, ls="-"):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}", fc=fc, ec=ec, lw=lw,
                                zorder=z, linestyle=ls))


def block(ax, x, y, w, h, title, lines=(), fc=BLUE, tc="white", ts=11, ls=9.2, ec=None, lw=1.2, z=3, title_dy=None):
    """Caja con título y líneas de detalle; las líneas se reparten bajo el título con posiciones explícitas."""
    rbox(ax, x, y, w, h, fc, ec or fc, lw=lw, z=z)
    if not lines:
        ax.text(x + w / 2, y + h / 2, title, ha="center", va="center", fontsize=ts, color=tc, weight="bold", zorder=z + 1)
        return
    n = len(lines)
    top = y + h - (title_dy if title_dy is not None else min(0.28, h / (n + 1.6)))
    ax.text(x + w / 2, top, title, ha="center", va="center", fontsize=ts, color=tc, weight="bold", zorder=z + 1)
    step = (top - y - 0.12) / (n + 0.35)
    for k, line in enumerate(lines, start=1):
        ax.text(x + w / 2, top - k * step, line, ha="center", va="center", fontsize=ls, color=tc, zorder=z + 1)


def arrow(ax, x1, y1, x2, y2, color=MUTED, lw=1.6, rad=0.0, z=2, ls="-"):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), zorder=z,
                arrowprops=dict(arrowstyle="-|>", color=color, lw=lw, shrinkA=0, shrinkB=0, linestyle=ls,
                                connectionstyle=f"arc3,rad={rad}", mutation_scale=14))


def marker(ax, x, y, k, color=CHG):
    ax.text(x, y, str(k), ha="center", va="center", fontsize=9.5, color="white", weight="bold", zorder=9,
            bbox=dict(boxstyle="circle,pad=0.22", fc=color, ec="white", lw=1.0))


def tag(ax, x, y, label, color):
    ax.text(x, y, label, ha="center", va="center", fontsize=9.5, color="white", weight="bold", zorder=9,
            bbox=dict(boxstyle="round,pad=0.32,rounding_size=0.25", fc=color, ec=color))


def label(ax, x, y, s, color=BLUE_D, size=9.5):
    ax.text(x, y, s, ha="center", va="center", fontsize=size, color=color, weight="bold", zorder=8,
            bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none"))


def draw(variant="base", fname="arq_base.png"):
    fig = plt.figure(figsize=(16, 9))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 16); ax.set_ylim(0, 9); ax.axis("off")
    A, B = variant == "A", variant == "B"

    # ------------------------------------------------------------------ marcos
    rbox(ax, 0.25, 0.85, 7.25, 7.8, BLUE_P, BLUE, lw=2.2, r=0.25, z=1)
    rbox(ax, 8.05, 0.85, 7.7, 7.8, ORANGE_P, ORANGE, lw=2.2, r=0.25, z=1)
    ax.text(0.55, 8.25, "ENCODER", fontsize=17, weight="bold", color=BLUE, va="center")
    ax.text(7.25, 8.25, "estático · se ejecuta 1 vez por instancia", fontsize=10.5, color=MUTED, va="center", ha="right")
    ax.text(8.35, 8.25, "DECODER", fontsize=17, weight="bold", color=ORANGE, va="center")
    ax.text(15.5, 8.25, "dinámico · se ejecuta en cada uno de los W pasos", fontsize=10.5, color=MUTED, va="center",
            ha="right")

    # ------------------------------------------------------------------ encoder: entradas
    ax.text(0.5, 7.75, "ENTRA (solo información de la instancia)", fontsize=9, color=BLUE_D, weight="bold")
    block(ax, 0.5, 6.55, 2.2, 1.0, "Armas i = 1..W", ["token aprendible $\\mathbf{w}$", "(sin información propia)"],
          fc=NAVY, ts=10.5, ls=8.8)
    block(ax, 2.85, 6.55, 2.05, 1.0, "Matriz $p_{ij}$ (W×T)", ["aristas del grafo", "bipartito ($\\tilde p_{ij}$)"],
          fc=NAVY, ts=10.5, ls=8.8)
    block(ax, 5.05, 6.55, 2.2, 1.0, "Objetivos j = 1..T", ["valor $V_j$:", "Linear$(V_j/\\max\\,V)$"],
          fc=NAVY, ts=10.5, ls=8.8)
    if variant == "base":
        marker(ax, 0.62, 7.47, 1)
    for x in (1.6, 3.87, 6.15):
        arrow(ax, x, 6.55, x, 6.22)

    # ------------------------------------------------------------------ encoder: 2 capas bipartitas
    g_lo = 3.0
    rbox(ax, 0.5, g_lo, 6.75, 6.2 - g_lo, "white", BLUE, lw=1.4, r=0.15, z=2, ls="--")
    ax.text(7.08, 5.98, "× 2", fontsize=14, color=BLUE, weight="bold", ha="right", va="center")
    if variant == "base":
        marker(ax, 0.62, 6.08, 2)
    block(ax, 0.75, 4.45, 3.05, 1.42, "Armas ← objetivos",
          ["MHA de 4 cabezas con aristas", "$s_{ij}=q_i\\cdot k_j/\\sqrt{d_h}+\\beta_h\\,\\tilde p_{ij}$",
           "mensaje: $v_j + w_e\\,\\tilde p_{ij}$"], fc=BLUE, ts=10.5, ls=8.9)
    block(ax, 3.95, 4.45, 3.05, 1.42, "Objetivos ← armas",
          ["MHA de 4 cabezas con aristas", "$s_{ji}=q_j\\cdot k_i/\\sqrt{d_h}+\\beta_h\\,\\tilde p_{ij}$",
           "mensaje: $v_i + w_e\\,\\tilde p_{ij}$"], fc=BLUE, ts=10.5, ls=8.9)
    if A:
        block(ax, 0.75, 3.8, 6.25, 0.5, "", fc=PURPLE_L, ec=PURPLE, lw=1.8)
        ax.text(3.875, 4.05, "sesgo atencional:  $+\\,\\alpha^{(l,\\,dir,\\,h)}\\cdot\\ln(V_j\\,p_{ij}\\,/\\,\\max\\,V p)$"
                "  en ambos puntajes  (α inician en 0)", ha="center", va="center", fontsize=9.3, color=PURPLE,
                weight="bold", zorder=6)
        tag(ax, 5.7, 6.0, "VARIANTE A", PURPLE)
        block(ax, 0.75, 3.15, 6.25, 0.5, "residual + LayerNorm · FF 64→128→64 · residual + LayerNorm",
              fc=BLUE_L, tc=BLUE_D, ts=9.6)
    else:
        block(ax, 0.75, 3.3, 6.25, 0.85, "residual + LayerNorm · FF 64→128→64 · residual + LayerNorm",
              fc=BLUE_L, tc=BLUE_D, ts=9.6)
    arrow(ax, 3.875, g_lo, 3.875, 2.67)
    block(ax, 0.5, 1.6, 6.75, 1.05, "SALE: $H^w\\in\\mathbb{R}^{W\\times d}$ y $H^t\\in\\mathbb{R}^{T\\times d}$",
          ["y precalcula las partes estáticas del puntero: $W_A h^w_i$ y $W_E\\,\\tilde p_{ij}$"], fc=BLUE_D, ts=11,
          ls=9.1)
    ax.text(3.875, 1.15, "Costo: $O(W T d + (W{+}T)\\,d^2)$ por capa → lineal en $n = W\\cdot T$",
            fontsize=10, color=BLUE_D, ha="center", va="center", style="italic")

    # ------------------------------------------------------------------ decoder
    X0, WD = 8.3, 7.2
    xc = X0 + WD / 2
    ax.text(X0, 7.75, "ENTRA (estado de la solución parcial)", fontsize=9, color=ORANGE_D, weight="bold")
    block(ax, X0, 6.72, WD, 0.88, "Estado $s(t)$: armas libres (máscara) · por objetivo: $q_j$, $v_j q_j$, $\\ln(v_j q_j)$",
          ["progreso: $k/W$ armas libres y fracción del valor que aún sobrevive $\\Sigma_j\\, v_j q_j\\,/\\,\\Sigma_j\\, v_j$"],
          fc=ORANGE_D, ts=10, ls=8.9)
    if variant == "base":
        marker(ax, X0 + 0.12, 7.52, 3)
    arrow(ax, xc, 6.72, xc, 6.5)
    block(ax, X0, 5.9, WD, 0.6, "Objetivos con su estado:  $h_j(t)=\\mathrm{LN}(h^t_j+\\mathrm{MLP}(s_j(t)))$", fc=ORANGE,
          ts=10.2)
    arrow(ax, xc, 5.9, xc, 5.62)

    if B:
        rbox(ax, X0, 3.88, WD, 1.74, GREEN_L, GREEN, lw=1.8, r=0.12, z=2)
        tag(ax, X0 + WD - 0.75, 5.37, "VARIANTE B", GREEN)
        ax.text(X0 + 0.15, 5.37, "pares factibles $Z_{ij}(t)=\\mathrm{ReLU}(W_A h^w_i+W_B h_j(t)+W_E\\,\\tilde p_{ij})$",
                fontsize=9.6, color=GREEN, weight="bold", va="center", zorder=6)
        block(ax, X0 + 0.15, 4.5, 3.3, 0.72, "ISAB · inducción",
              ["$m=16$ puntos $I$:  $I'=\\mathrm{LN}(I+\\mathrm{MHA}(I,Z,Z))$ + FF"], fc=GREEN, ts=9.8, ls=8.4)
        block(ax, X0 + 3.75, 4.5, 3.3, 0.72, "ISAB · actualización",
              ["cada par lee el resumen:  $Z'=\\mathrm{LN}(Z+\\mathrm{MHA}(Z,I',I'))$ + FF"], fc=GREEN, ts=9.8, ls=8.1)
        arrow(ax, X0 + 3.45, 4.86, X0 + 3.75, 4.86, color=GREEN)
        block(ax, X0 + 0.15, 3.97, WD - 0.3, 0.42,
              "contexto $c = W_c[\\,\\overline{I'}\\,;\\,\\mathrm{progreso}\\,]$  ·  costo $O(n\\,m\\,d)$ en lugar de $O(n^2 d)$",
              fc="white", tc=GREEN, ec=GREEN, ts=9.4)
        ptr_lines = ["$u_{ij}=C\\,\\tanh((W_r c)^\\top W_k Z'_{ij}/\\sqrt{d})$"]
        arrow(ax, xc, 3.88, xc, 3.7)
    else:
        block(ax, X0, 5.02, WD, 0.6, "Contexto:  $c=W_c[\\,\\overline{h^w}_{\\mathrm{libres}}\\,;\\,\\overline{h(t)}\\,;\\,"
              "\\mathrm{progreso}\\,]$", fc=ORANGE, ts=10.2)
        if variant == "base":
            marker(ax, X0 + 0.12, 5.52, 4)
        arrow(ax, xc, 5.02, xc, 4.78)
        block(ax, X0, 3.92, WD, 0.86, "Glimpse (MHA de 4 cabezas)",
              ["$c$ atiende a las armas libres y a los objetivos $h_j(t)$; residual + LN + FF"], fc=ORANGE, ts=10.2,
              ls=8.9)
        ptr_lines = ["$u_{ij}=C\\,\\tanh((W_r c)^\\top\\,\\mathrm{ReLU}(W_A h^w_i+W_B h_j(t)+W_E\\,\\tilde p_{ij})/\\sqrt{d})$"]
        arrow(ax, xc, 3.92, xc, 3.7)
    block(ax, X0, 2.95, WD, 0.75, "Puntero aditivo por pares (con la arista $p_{ij}$)", ptr_lines, fc=RED, ts=10, ls=9.6)
    if variant == "base":
        marker(ax, X0 + 0.12, 3.62, 5)
    if A:
        block(ax, X0, 2.4, WD, 0.42, "", fc=PURPLE_L, ec=PURPLE, lw=1.8)
        ax.text(xc, 2.61, "sesgo de salida:  $+\\,\\alpha\\cdot\\ln(g_{ij}(t)/g_{\\max}(t))$,  $g_{ij}=V_j\\,q_j\\,p_{ij}$"
                " (ganancia de MMR; α inicia en 1)", ha="center", va="center", fontsize=9.3, color=PURPLE,
                weight="bold", zorder=6)
        tag(ax, X0 + WD - 0.75, 3.45, "VARIANTE A", PURPLE)
        arrow(ax, xc, 2.95, xc, 2.82)
        y_mask = 1.86
        arrow(ax, xc, 2.4, xc, y_mask + 0.42)
    else:
        y_mask = 2.25
        arrow(ax, xc, 2.95, xc, y_mask + 0.42)
    block(ax, X0, y_mask, WD, 0.42, "Máscara (armas ya usadas → $-10^9$) · softmax sobre los W·T pares", fc=MASK, ts=9.8)
    y_act = y_mask - 0.36
    ax.text(xc, y_act, "acción $(i^*, j^*)$:  $q_{j^*}\\leftarrow q_{j^*}(1-p_{i^*j^*})$,  arma $i^*$ usada  →  siguiente paso (W veces)",
            fontsize=9.6, color=ORANGE_D, ha="center", va="center")
    arrow(ax, xc, y_mask, xc, y_act + 0.16)
    arrow(ax, X0 + WD + 0.12, y_act, X0 + WD + 0.12, 7.16, color=ORANGE, lw=1.6, ls="--")
    cost = ("Costo por paso: $O(n\\,m\\,d + n\\,d^2)$ → lineal en $n$" if B
            else "Costo por paso: $O(n\\,d+(W{+}T)\\,d^2)$ → lineal en $n = W\\cdot T$")
    ax.text(xc, 1.1, cost, fontsize=10, color=ORANGE_D, ha="center", va="center", style="italic")

    # ------------------------------------------------------------------ conexiones encoder -> decoder
    arrow(ax, 7.25, 2.45, X0, 6.2, color=BLUE_D, lw=1.8, rad=-0.22)
    arrow(ax, 7.25, 2.12, X0, 5.32 if not B else 5.37, color=BLUE_D, lw=1.8, rad=-0.16)
    arrow(ax, 7.25, 1.8, X0, 3.3, color=BLUE_D, lw=1.8, rad=-0.1)
    label(ax, 7.78, 5.3, "$H^t$")
    label(ax, 7.78, 4.15, "$H^w$")
    label(ax, 7.74, 2.62, "$W_A h^w$\n$W_E\\tilde p$", size=8.6)

    if variant == "base":
        ax.text(0.3, 0.5, "Cambios respecto del módulo 2:  ① tokens armas + objetivos y $p_{ij}$ como aristas (antes 1+W·T filas con "
                "16 features)   ② encoder 1 vez por instancia, atención bipartita con aristas, 4 cabezas (antes en cada paso)",
                fontsize=8.7, color=CHG, va="center")
        ax.text(0.3, 0.2, "③ el estado dinámico entra solo al decoder, sin features heurísticas   ④ contexto con promedios + glimpse "
                "MHA (antes fila global + promedio)   ⑤ puntero aditivo por pares con la arista $p_{ij}$ y recorte $C\\,\\tanh$",
                fontsize=8.7, color=CHG, va="center")
    else:
        txt = ("Variante A = base + sesgo inductivo con el criterio de MMR: atencional en el encoder y de salida en el puntero"
               if A else "Variante B = base con el contexto (promedios + glimpse) reemplazado por un bloque ISAB sobre los pares "
                         "factibles")
        ax.text(8.0, 0.38, txt, fontsize=10.5, color=PURPLE if A else GREEN, va="center", ha="center", weight="bold")
    fig.savefig(os.path.join(FIG, fname), dpi=170)
    plt.close(fig)
    print("✔", fname)


if __name__ == "__main__":
    draw("base", "arq_base.png")
    draw("A", "arq_variante_A.png")
    draw("B", "arq_variante_B.png")
