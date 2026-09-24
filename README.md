# OII462 · Tarea 4 — Imitando un algoritmo de trayectoria para el WTAP

Modelo de capas densas (MLP en PyTorch) que aprende a construir soluciones del
**Weapon-Target Assignment Problem (WTAP)** imitando, paso a paso, las decisiones de un
algoritmo experto (Búsqueda Local Iterada, verificada contra un método exacto).

## Entregables

| Entregable | Archivo |
|---|---|
| Notebook (Colab) | [`Tarea4_WTAP_imitacion.ipynb`](Tarea4_WTAP_imitacion.ipynb) — [abrir en Colab](https://colab.research.google.com/github/Tomas-pucv/OII462-1/blob/claude/dense-model-wtap-rcjk1o/Tarea4_WTAP_imitacion.ipynb) |
| Presentación | [`Tarea4_WTAP_presentacion.pptx`](Tarea4_WTAP_presentacion.pptx) · [`Tarea4_WTAP_presentacion.pdf`](Tarea4_WTAP_presentacion.pdf) |
| Figuras y resultados | [`figuras/`](figuras) (generadas por el notebook, incluye `resultados.json`) |

> Si la rama se fusiona a `main`, reemplaza `claude/dense-model-wtap-rcjk1o` por `main` en el link de Colab.

## Resumen

- **Problema:** W = 10 armas, T = 8 objetivos; minimizar el valor esperado que sobrevive
  `Σ_j V_j · Π_i (1 − p_ij)^{x_ij}` (NP-hard).
- **Acción (constructiva):** a qué objetivo se asigna la arma actual (armas en orden fijo) → 8 clases.
- **Estado (`state2vec`):** matriz `8 × 15`, fila `j` = objetivo `j` (valor, prob. de sobrevivir,
  ganancia marginal de la arma actual, características de anticipación, …).
- **Experto:** Búsqueda Local Iterada (inicio MMR + búsqueda local `reassign`/`swap` + perturbación),
  que alcanza el óptimo en las instancias verificadas con un método exacto *meet-in-the-middle*.
- **Datos:** 6 000 instancias resueltas → 60 000 pares (estado → objetivo).
- **Modelo:** MLP con Dropout, entropía cruzada, Adam y early stopping.
- **Rollout:** el modelo reemplaza a la función de evaluación del `GreedyAgent` de la librería
  [`eda`](https://github.com/rilianx/eda) y se compara contra la heurística clásica **MMR**
  (*Maximum Marginal Return*), el greedy secuencial y el experto en 50 instancias nuevas.

## Reproducir

El notebook corre de principio a fin en Colab (CPU); clona la librería `eda` del curso en la
primera celda. La presentación se regenera a partir de `figuras/`:

```bash
python scripts/make_presentation.py
soffice --headless --convert-to pdf Tarea4_WTAP_presentacion.pptx
```
