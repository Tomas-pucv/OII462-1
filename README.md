# OII462 · Imitación de algoritmos para el WTAP

Modelos neuronales que aprenden a construir soluciones del **Weapon-Target Assignment Problem (WTAP)**
imitando, paso a paso, las decisiones de un algoritmo experto (Búsqueda Local Iterada, verificada
contra un método exacto).

- **Tarea 5:** encoder con self-attention y capa de decisión tipo *pointer* (acciones = pares arma → objetivo).
- **Tarea 4:** MLP de capas densas (armas en orden fijo).

---

## Tarea 5 — Datos de imitación y modelo de atención con puntero

### Entregables

| Entregable | Archivo |
|---|---|
| Informe (puntos 1–3) | [`Tarea5_WTAP_informe.pdf`](Tarea5_WTAP_informe.pdf) · [`.docx` editable](Tarea5_WTAP_informe.docx) |
| Notebook (Colab, puntos 4–5 y opcionales 7–8) | [`Tarea5_WTAP_puntero.ipynb`](Tarea5_WTAP_puntero.ipynb) — [abrir en Colab](https://colab.research.google.com/github/Tomas-pucv/OII462-1/blob/claude/festive-pasteur-se68vm/Tarea5_WTAP_puntero.ipynb) |
| Presentación (punto 6) | [`Tarea5_WTAP_presentacion.pptx`](Tarea5_WTAP_presentacion.pptx) · [`Tarea5_WTAP_presentacion.pdf`](Tarea5_WTAP_presentacion.pdf) |
| Figuras y resultados | [`figuras_t5/`](figuras_t5) (generadas por el notebook, incluye `resultados.json`) |

> Si la rama se fusiona, reemplaza `claude/festive-pasteur-se68vm` por la rama destino en el link de Colab.
> El nombre del grupo aparece como `[Nombre(s) del grupo]` en el informe y la presentación.

### Resumen

- **Decisión puntual:** dado el estado, *¿qué arma asigno ahora y a qué objetivo?* Cada acción posible es un
  **par (arma i → objetivo j)**, lo mismo que evalúa la heurística clásica **MMR**.
- **Descomposición (punto 1):** acciones constructivas. Se parte de la asignación vacía y, en cada paso, se agrega el
  par del experto con mayor ganancia marginal `V_j·q_j·p_ij` entre las armas libres. Hay W decisiones por instancia.
- **Entrada (punto 2):** matriz `(1 + W·T) × 16`. La fila 0 es una fila global `[CLS]` (flag `es_global`, forma la
  query) y la fila `1 + i·T + j` es el par (i, j). Las columnas son propias del par, flags (`no_disponible` = máscara)
  y relativas al estado, tomadas de MMR: ganancia, arrepentimiento y completar con MMR.
- **Salida (punto 3):** un logit por fila, logit −10⁹ en las filas no disponibles, etiqueta = índice de la fila
  escogida y `CrossEntropyLoss`.
- **Modelo (punto 5):** `WTAP_Attention`, adaptado de `TSP_Attention`: embedding compartido, 2 bloques de
  self-attention con LayerNorm y Dropout, query `Wc[h_global ; promedio de pares disponibles]` y puntero con máscara
  en la columna 4. Ningún peso depende de W ni de T.
- **Evaluación (opcional):** se compara en las mismas 50 instancias de prueba de la Tarea 4 contra MMR, el experto y
  la MLP de la Tarea 4, y se prueba sin reentrenar en instancias de 15×10 y 20×15.

### Reproducir

El notebook corre de principio a fin en Colab. Clona la librería [`eda`](https://github.com/rilianx/eda) en la
primera celda y usa GPU si está disponible. El informe y la presentación se regeneran a partir de `figuras_t5/`:

```bash
python scripts/make_informe_t5.py
python scripts/make_presentation_t5.py
soffice --headless --convert-to pdf Tarea5_WTAP_informe.docx Tarea5_WTAP_presentacion.pptx
```

---

## Tarea 4 — Imitando un algoritmo de trayectoria para el WTAP

Modelo de capas densas (MLP en PyTorch) que aprende a construir soluciones del WTAP imitando, paso a paso, las
decisiones de un algoritmo experto.

### Entregables

| Entregable | Archivo |
|---|---|
| Notebook (Colab) | [`Tarea4_WTAP_imitacion.ipynb`](Tarea4_WTAP_imitacion.ipynb) — [abrir en Colab](https://colab.research.google.com/github/Tomas-pucv/OII462-1/blob/claude/dense-model-wtap-rcjk1o/Tarea4_WTAP_imitacion.ipynb) |
| Presentación | [`Tarea4_WTAP_presentacion.pptx`](Tarea4_WTAP_presentacion.pptx) · [`Tarea4_WTAP_presentacion.pdf`](Tarea4_WTAP_presentacion.pdf) |
| Figuras y resultados | [`figuras/`](figuras) (generadas por el notebook, incluye `resultados.json`) |

> Si la rama se fusiona a `main`, reemplaza `claude/dense-model-wtap-rcjk1o` por `main` en el link de Colab.

### Resumen

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

### Reproducir

El notebook corre de principio a fin en Colab (CPU); clona la librería `eda` del curso en la
primera celda. La presentación se regenera a partir de `figuras/`:

```bash
python scripts/make_presentation.py
soffice --headless --convert-to pdf Tarea4_WTAP_presentacion.pptx
```
