# OII462 · Imitación de algoritmos para el WTAP

Modelos neuronales que aprenden a construir soluciones del **Weapon-Target Assignment Problem (WTAP)**
imitando, paso a paso, las decisiones de un algoritmo experto (Búsqueda Local Iterada, verificada
contra un método exacto).

- **Proyecto del módulo:** presentación final (13–15 min) del Transformer encoder–decoder con puntero.
- **Tarea 5:** encoder con self-attention y capa de decisión tipo *pointer* (acciones = pares arma → objetivo).
- **Tarea 4:** MLP de capas densas (armas en orden fijo).
- **Laboratorio 2:** Transformer implementado desde sus componentes para clasificar el sentimiento de reseñas IMDB.

---

## Laboratorio 2 — Transformers para clasificación de sentimientos (IMDB)

### Entregable

| Entregable | Archivo |
|---|---|
| Notebook (Colab) | [`Lab2_2026_2.ipynb`](Lab2_2026_2.ipynb) — [abrir en Colab](https://colab.research.google.com/github/Tomas-pucv/OII462-1/blob/claude/relaxed-wozniak-0hk3aa/Lab2_2026_2.ipynb) |

### Resumen

- **EDA:** longitudes sesgadas a la derecha (mediana 173 palabras, cola hasta 2.470), etiquetas balanceadas y
  porcentaje de reseñas truncadas según `max_length`.
- **Tokenizador BPE de 30.000 fragmentos** (cap. 12.5): gráfico del tamaño del vocabulario por iteración; las
  fusiones se aplican en el orden en que se aprendieron y la puntuación se separa de las palabras.
- **Atención** (cap. 12.2–12.3): `ScaledDotProductAttention`, `SingleHeadAttention` (Q/K/V) y
  `MultiHeadAttention`, con máscara de padding y verificada contra `F.scaled_dot_product_attention`.
- **Capa Transformer** (cap. 12.4): atención + residual + LayerNorm, FFN (ReLU, Dropout) + residual + LayerNorm.
- **Entrenamiento:** early stopping sobre la precisión de validación, AdamW, warm-up y comparación de 5 tamaños
  (0,97 M a 10,8 M parámetros).
- **Evaluación:** precisión por cuartil de longitud; bonus con 2 capas, 4 cabezas y `max_length = 512`:
  **79,0% en prueba** (no alcanza el 90%; el modelo base provisto da 63,0% y una regresión logística TF-IDF de
  referencia, 84,4%).
- **Despliegue:** casos de falla (reseñas muy cortas, sarcasmo, contraste, doble negación) y reseñas en español
  (BPE no genera `[UNK]`, pero divide las palabras en fragmentos sin significado para el modelo).

### Reproducir

El notebook corre de principio a fin en Colab (GPU T4 recomendada) y descarga `stanfordnlp/imdb` desde
Hugging Face. Las salidas guardadas son de una corrida en una GPU T4 de Colab; al re-ejecutar, las cifras del
entrenamiento pueden variar algunos puntos (dropout en GPU, orden de las sumas y desempates del tokenizador).

---

## Proyecto del módulo — Transformer con puntero para el WTAP

### Entregables

| Entregable | Archivo |
|---|---|
| Presentación (14 diapositivas, con notas del orador y tiempos) | [`Proyecto_WTAP_presentacion.pptx`](Proyecto_WTAP_presentacion.pptx) · [`Proyecto_WTAP_presentacion.pdf`](Proyecto_WTAP_presentacion.pdf) |
| Notebook (Colab) | [`Proyecto_WTAP_transformer.ipynb`](Proyecto_WTAP_transformer.ipynb) — [abrir en Colab](https://colab.research.google.com/github/Tomas-pucv/OII462-1/blob/claude/dense-model-wtap-rcjk1o/Proyecto_WTAP_transformer.ipynb) |
| Figuras y resultados | [`figuras_proyecto/`](figuras_proyecto) (generadas por el notebook, incluye `resultados.json`) |

La presentación sigue la guía del módulo: (1) problema con definición formal y ejemplo visual, (2) secuencia de
entrada y justificación de cada una de las 16 features, (3) generación de datos (experto ILS, descomposición y
codificación de la salida), (4) diagrama de la arquitectura con encoder y decoder diferenciados, y (5) experimentos:
tabla de hiperparámetros, convergencia, evaluación con media ± desviación contra MMR y un constructivo aleatorio,
desempeño por tamaño de instancia (5×4 a 30×20) y discusión.

### Modelo

- **Encoder:** embedding lineal compartido 16 → 64 y 2 capas de self-attention (1 cabeza) + feed-forward 64 → 128 → 64,
  con residual y LayerNorm. Entrada: la secuencia de `1 + W·T` filas (fila global + un par arma → objetivo por fila).
- **Decoder:** contexto `[h_global ; promedio de pares disponibles]` → query → puntero `q·h_r/√d` → máscara de pares no
  disponibles → softmax sobre las acciones (sin glimpse; ver la comparación con Kool et al. en la presentación).
- Mismo modelo y pipeline de la Tarea 5, reentrenado en el notebook del proyecto.

### Reproducir

```bash
python scripts/make_presentation_proyecto.py
soffice --headless --convert-to pdf Proyecto_WTAP_presentacion.pptx
```

La diapositiva 2 usa además `figuras/instancia.png` (la misma instancia de ejemplo de la Tarea 4).

---

## Tarea 5 — Datos de imitación y modelo de atención con puntero

### Entregables

| Entregable | Archivo |
|---|---|
| Informe (puntos 1–3) | [`Tarea5_WTAP_informe.pdf`](Tarea5_WTAP_informe.pdf) · [`.docx` editable](Tarea5_WTAP_informe.docx) |
| Notebook (Colab, puntos 4–5 y opcionales 7–8) | [`Tarea5_WTAP_puntero.ipynb`](Tarea5_WTAP_puntero.ipynb) — [abrir en Colab](https://colab.research.google.com/github/Tomas-pucv/OII462-1/blob/claude/dense-model-wtap-rcjk1o/Tarea5_WTAP_puntero.ipynb) |
| Presentación (punto 6) | [`Tarea5_WTAP_presentacion.pptx`](Tarea5_WTAP_presentacion.pptx) · [`Tarea5_WTAP_presentacion.pdf`](Tarea5_WTAP_presentacion.pdf) |
| Figuras y resultados | [`figuras_t5/`](figuras_t5) (generadas por el notebook, incluye `resultados.json`) |

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
