"""Diagrama de la arquitectura de la Tarea 7 (encoder EGAT + decoder de la Base de la Tarea 6) en figuras_t7/.

El diagrama del grafo (figuras_t7/grafo_wtap.png) lo genera el notebook, a partir de una instancia real.

Uso (desde la raíz del repositorio):
    python scripts/make_diagramas_t7.py
"""
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_diagramas_t6 import (BLUE, BLUE_D, BLUE_L, BLUE_P, CHG, MASK, MUTED, NAVY, ORANGE, ORANGE_D, ORANGE_P,  # noqa: E402
                               RED, arrow, block, label, marker, rbox, tag)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG = os.path.join(ROOT, "figuras_t7")
os.makedirs(FIG, exist_ok=True)
GREEN_D = "#1E7B34"


def draw(fname="arq_egat.png"):
    fig = plt.figure(figsize=(16, 9))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 16); ax.set_ylim(0, 9); ax.axis("off")

    # ------------------------------------------------------------------ marcos
    rbox(ax, 0.25, 0.85, 7.25, 7.8, BLUE_P, BLUE, lw=2.2, r=0.25, z=1)
    rbox(ax, 8.05, 0.85, 7.7, 7.8, ORANGE_P, ORANGE, lw=2.2, r=0.25, z=1)
    ax.text(0.55, 8.25, "ENCODER EGAT", fontsize=17, weight="bold", color=BLUE, va="center")
    ax.text(7.25, 8.25, "estático · se ejecuta 1 vez por instancia", fontsize=10.5, color=MUTED, va="center", ha="right")
    ax.text(8.35, 8.25, "DECODER", fontsize=17, weight="bold", color=ORANGE, va="center")
    ax.text(15.5, 8.25, "dinámico · se ejecuta en cada uno de los W pasos", fontsize=10.5, color=MUTED, va="center",
            ha="right")

    # ------------------------------------------------------------------ encoder: el grafo bipartito
    ax.text(0.5, 7.75, "ENTRA: el grafo bipartito de la instancia", fontsize=9, color=BLUE_D, weight="bold")
    block(ax, 0.5, 6.55, 2.2, 1.0, "Nodos arma i", ["token aprendible común", "(sin features propias)"],
          fc=NAVY, ts=10.5, ls=8.8)
    block(ax, 2.85, 6.55, 2.05, 1.0, "Aristas $K_{W,T}$", ["atributo $\\tilde p_{ij}$", "+ auto-bucles"],
          fc=NAVY, ts=10.5, ls=8.8)
    block(ax, 5.05, 6.55, 2.2, 1.0, "Nodos objetivo j", ["valor $V_j$:", "Linear$(V_j/\\max\\,V)$"],
          fc=NAVY, ts=10.5, ls=8.8)
    marker(ax, 2.97, 7.47, 1)
    for x in (1.6, 3.87, 6.15):
        arrow(ax, x, 6.55, x, 6.22)

    # ------------------------------------------------------------------ encoder: L capas EGAT
    g_lo = 3.0
    rbox(ax, 0.5, g_lo, 6.75, 6.2 - g_lo, "white", BLUE, lw=1.4, r=0.15, z=2, ls="--")
    ax.text(7.1, 6.03, "× L capas  (L = 1, 2, 3)", fontsize=12.5, color=BLUE, weight="bold", ha="right", va="center")
    marker(ax, 4.62, 6.03, 5)
    lines_w = ["$z = W\\,h$ por tipo de nodo",
               "$e_{ij}=\\mathrm{LReLU}(a_L\\!\\cdot\\! z_i+a_R\\!\\cdot\\! z_j+a_E\\!\\cdot\\! E_{ij})$",
               "softmax: vecinos + auto-bucle", "mensaje: $z_j + E_{ij}$,  $E_{ij}=W_e\\,\\tilde p_{ij}$"]
    block(ax, 0.75, 4.25, 3.05, 1.62, "Armas ← objetivos", lines_w, fc=BLUE, ts=10.5, ls=8.5)
    lines_t = ["$z = W\\,h$ por tipo de nodo",
               "$e_{ji}=\\mathrm{LReLU}(a_L\\!\\cdot\\! z_j+a_R\\!\\cdot\\! z_i+a_E\\!\\cdot\\! E_{ij})$",
               "softmax: vecinos + auto-bucle", "mensaje: $z_i + E_{ij}$"]
    block(ax, 3.95, 4.25, 3.05, 1.62, "Objetivos ← armas", lines_t, fc=BLUE, ts=10.5, ls=8.5)
    marker(ax, 0.87, 5.74, 2)
    marker(ax, 0.87, 5.27, 3)
    marker(ax, 0.87, 4.42, 4)
    block(ax, 0.75, 3.3, 6.25, 0.75, "residual + LayerNorm · FF 64→128→64 · residual + LayerNorm",
          ["4 cabezas · 1 capa = 1 salto del campo receptivo"], fc=BLUE_L, tc=BLUE_D, ts=9.6, ls=8.8)
    arrow(ax, 3.875, g_lo, 3.875, 2.67)
    block(ax, 0.5, 1.6, 6.75, 1.05, "SALE: $H^w\\in\\mathbb{R}^{W\\times d}$ y $H^t\\in\\mathbb{R}^{T\\times d}$",
          ["y precalcula las partes estáticas del puntero: $W_A h^w_i$ y $W_E\\,\\tilde p_{ij}$"], fc=BLUE_D, ts=11,
          ls=9.1)
    ax.text(3.875, 1.15, "Costo: $O(W T d + (W{+}T)\\,d^2)$ por capa → lineal en el número de aristas",
            fontsize=10, color=BLUE_D, ha="center", va="center", style="italic")

    # ------------------------------------------------------------------ decoder (el de la Base de la Tarea 6)
    X0, WD = 8.3, 7.2
    xc = X0 + WD / 2
    ax.text(X0, 7.75, "ENTRA (estado de la solución parcial)", fontsize=9, color=ORANGE_D, weight="bold")
    tag(ax, X0 + WD - 2.05, 7.8, "SIN CAMBIOS: decoder de la Base (Tarea 6)", GREEN_D)
    block(ax, X0, 6.72, WD, 0.88, "Estado $s(t)$: armas libres (máscara) · por objetivo: $q_j$, $v_j q_j$, $\\ln(v_j q_j)$",
          ["progreso: $k/W$ armas libres y fracción del valor que aún sobrevive $\\Sigma_j\\, v_j q_j\\,/\\,\\Sigma_j\\, v_j$"],
          fc=ORANGE_D, ts=10, ls=8.9)
    arrow(ax, xc, 6.72, xc, 6.5)
    block(ax, X0, 5.9, WD, 0.6, "Objetivos con su estado:  $h_j(t)=\\mathrm{LN}(h^t_j+\\mathrm{MLP}(s_j(t)))$", fc=ORANGE,
          ts=10.2)
    arrow(ax, xc, 5.9, xc, 5.62)
    block(ax, X0, 5.02, WD, 0.6, "Contexto:  $c=W_c[\\,\\overline{h^w}_{\\mathrm{libres}}\\,;\\,\\overline{h(t)}\\,;\\,"
          "\\mathrm{progreso}\\,]$", fc=ORANGE, ts=10.2)
    arrow(ax, xc, 5.02, xc, 4.78)
    block(ax, X0, 3.92, WD, 0.86, "Glimpse (MHA de 4 cabezas)",
          ["$c$ atiende a las armas libres y a los objetivos $h_j(t)$; residual + LN + FF"], fc=ORANGE, ts=10.2, ls=8.9)
    arrow(ax, xc, 3.92, xc, 3.7)
    block(ax, X0, 2.95, WD, 0.75, "Puntero aditivo por pares (con la arista $p_{ij}$)",
          ["$u_{ij}=C\\,\\tanh((W_r c)^\\top\\,\\mathrm{ReLU}(W_A h^w_i+W_B h_j(t)+W_E\\,\\tilde p_{ij})/\\sqrt{d})$"],
          fc=RED, ts=10, ls=9.6)
    y_mask = 2.25
    arrow(ax, xc, 2.95, xc, y_mask + 0.42)
    block(ax, X0, y_mask, WD, 0.42, "Máscara (armas ya usadas → $-10^9$) · softmax sobre los W·T pares", fc=MASK, ts=9.8)
    y_act = y_mask - 0.36
    ax.text(xc, y_act, "acción $(i^*, j^*)$:  $q_{j^*}\\leftarrow q_{j^*}(1-p_{i^*j^*})$,  arma $i^*$ usada  →  "
            "siguiente paso (W veces)", fontsize=9.6, color=ORANGE_D, ha="center", va="center")
    arrow(ax, xc, y_mask, xc, y_act + 0.16)
    arrow(ax, X0 + WD + 0.12, y_act, X0 + WD + 0.12, 7.16, color=ORANGE, lw=1.6, ls="--")
    ax.text(xc, 1.1, "Costo por paso: $O(n\\,d+(W{+}T)\\,d^2)$ → lineal en $n = W\\cdot T$", fontsize=10,
            color=ORANGE_D, ha="center", va="center", style="italic")

    # ------------------------------------------------------------------ conexiones encoder -> decoder
    arrow(ax, 7.25, 2.45, X0, 6.2, color=BLUE_D, lw=1.8, rad=-0.22)
    arrow(ax, 7.25, 2.12, X0, 5.32, color=BLUE_D, lw=1.8, rad=-0.16)
    arrow(ax, 7.25, 1.8, X0, 3.3, color=BLUE_D, lw=1.8, rad=-0.1)
    label(ax, 7.78, 5.3, "$H^t$")
    label(ax, 7.78, 4.15, "$H^w$")
    label(ax, 7.74, 2.62, "$W_A h^w$\n$W_E\\tilde p$", size=8.6)

    ax.text(0.3, 0.5, "Cambios respecto de la Tarea 6:  ① grafo bipartito explícito $K_{W,T}$ con atributo $\\tilde p_{ij}$ y "
            "auto-bucles   ② atención aditiva con LeakyReLU (antes producto punto $q_i\\cdot k_j/\\sqrt{d_h}$)",
            fontsize=8.7, color=CHG, va="center")
    ax.text(0.3, 0.2, "③ una proyección $W$ por tipo de nodo (antes $W_Q, W_K, W_V$)   ④ arista como vector $E_{ij}$ en el "
            "puntaje y en el mensaje   ⑤ profundidad L = 1, 2, 3 capas (ablación; antes 2)   ·   el decoder no cambia",
            fontsize=8.7, color=CHG, va="center")
    fig.savefig(os.path.join(FIG, fname), dpi=170)
    plt.close(fig)
    print("✔", fname)


if __name__ == "__main__":
    draw()
