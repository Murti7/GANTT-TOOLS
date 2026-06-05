"""
Generación del diagrama temporal de implantación por bandas operativas.

El módulo mantiene la función pública generar_diagrama_red(planificacion)
para no romper el pipeline existente, aunque la representación ya no es una
red PERT clásica, sino un diagrama temporal por bandas.
"""

import io
import textwrap
from collections import defaultdict
from datetime import date

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch

from gantt.planning.analyser import calcular_holguras, dias_habiles_entre
from gantt.planning.models import PlanificacionProyecto, TareaGantt


COLOR_CRITICO = "#FFD6D6"
COLOR_CRITICO_BORDE = "#FF0000"
COLOR_NORMAL = "#DDEEFF"
COLOR_NORMAL_BORDE = "#3A6F9F"
COLOR_HITO = "#FFE66D"
COLOR_HITO_BORDE = "#333333"
COLOR_FIN_PLAZO = "#7FD37F"
COLOR_POST = "#D9D9D9"
COLOR_POST_BORDE = "#666666"

COLOR_DEP = "#111111"
COLOR_DEP_CRITICA = "#FF0000"
COLOR_DOC = "#666666"
COLOR_GRID = "#CCCCCC"
COLOR_TEXTO = "#111111"

ANCHO_PANEL: float = 20.0
# Unidades de datos reservadas para el panel de títulos (izquierda).
# La línea H1 (día 0) queda siempre en x = ANCHO_PANEL.

BANDAS = [
    (
        "Inicio / documentación contractual",
        "INICIO / DOCUMENTACIÓN\nCONTRACTUAL",
        "\nDocumentación previa al inicio de la obra,\nPSS, PGR, CAE, etc.",
        "#EAF4FF",
    ),
    (
        "Auditorías técnicas",
        "AUDITORÍAS TÉCNICAS",
        "\nCaracterización inicial de UTs\ny deteccción de deficiencias",
        "#F3FAEA",
    ),
    (
        "Validación de inventarios y tipologías",
        "VALIDACIÓN DE INVENTARIOS\nY TIPOLOGÍAS",
        "\nRevisión STI de los inventarios técnicos de auditoría\n y autorización técnica de actuaciones de corrección propuestas.",
        "#FFF6DD",
    ),
    (
        "Actuaciones habilitadoras sistémicas",
        "ACTUACIONES HABILITADORAS\nSISTÉMICAS",
        "\nPurgadores y sectorización vertical\n, adecuación de la sala de máquinas de clima",
        "#F4E9FF",
    ),
    (
        "Ejecución correctiva por bloques",
        "EJECUCIÓN CORRECTIVA\nPOR BLOQUES",
        "\nCorrecciones principales de control,\nventilación, hidráulica y accesos.",
        "#FFF0E6",
    ),
    (
        "Validación funcional por bloques",
        "VALIDACIÓN FUNCIONAL\nPOR BLOQUES",
        "\nComprobación funcional de cada\nbloque antes de puesta en servicio.",
        "#ECF8EE",
    ),
    (
        "Puesta en servicio funcional",
        "PUESTA EN SERVICIO\nFUNCIONAL",
        "\nPruebas integradas y verificación\noperativa del sistema reformado.",
        "#E9F5FF",
    ),
    (
        "Cierre contractual",
        "CIERRE CONTRACTUAL",
        "\nRecepción operativa y fin del\nplazo de ejecución de las actuaciones de reforma.",
        "#FFF7D8",
    ),
    (
        "Validaciones estacionales diferidas",
        "VALIDACIONES ESTACIONALES\nDIFERIDAS",
        "\nEnsayos térmicos en régimen real\nde calefacción y refrigeración.",
        "#F2F2F2",
    ),
    (
        "Proyecto As Built",
        "PROYECTO AS BUILT",
        "\nDocumentación final, planos\nactualizados y cierre documental.",
        "#F7F7F7",
    ),
]


def fecha_base(tareas: list[TareaGantt]) -> date:
    h1 = next((t for t in tareas if t.id == "H1" and t.fecha_inicio), None)
    if h1 is not None:
        return h1.fecha_inicio

    fechas = [t.fecha_inicio for t in tareas if t.fecha_inicio is not None]
    if not fechas:
        raise ValueError("La planificación no contiene fechas calculadas.")
    return min(fechas)


def x_plot(fecha: date | None, base: date) -> float:
    """
    Transforma una fecha a coordenada X del plot.
    x = ANCHO_PANEL + días_hábiles_entre(base, fecha)
    Las tareas previas a H1 (base) tienen x < ANCHO_PANEL.
    """
    if fecha is None:
        return ANCHO_PANEL
    return ANCHO_PANEL + float(dias_habiles_entre(base, fecha))


def grupo(tarea: TareaGantt) -> str:
    if tarea.grupo_visual:
        return tarea.grupo_visual
    if tarea.estacional:
        return "Validaciones estacionales diferidas"
    if tarea.es_fin_plazo:
        return "Cierre contractual"
    return "Inicio / documentación contractual"


def agrupar_tareas(tareas: list[TareaGantt]) -> dict[str, list[TareaGantt]]:
    grupos: dict[str, list[TareaGantt]] = defaultdict(list)

    for tarea in tareas:
        grupos[grupo(tarea)].append(tarea)

    for tareasgrupo in grupos.values():
        tareasgrupo.sort(
            key=lambda t: (
                t.orden_visual if t.orden_visual is not None else 999,
                t.fecha_inicio or date.min,
                t.id,
            )
        )

    return grupos


def calcular_y(tareas: list[TareaGantt]) -> tuple[dict[str, float], dict[str, tuple[float, float]]]:
    grupos = agrupar_tareas(tareas)
    y_tareas: dict[str, float] = {}
    bandas: dict[str, tuple[float, float]] = {}

    y_actual = 0.0
    alto_min_banda = 1.85
    separacion_fila = 1
    separacion_banda = 0.22

    for nombre_banda, *_ in BANDAS:
        tareas_banda = grupos.get(nombre_banda, [])
        num_filas = max(len(tareas_banda), 1)
        alto_banda = max(alto_min_banda, num_filas * separacion_fila + 0.65)

        y_sup = y_actual
        y_inf = y_actual - alto_banda
        bandas[nombre_banda] = (y_inf, y_sup)

        if tareas_banda:
            if len(tareas_banda) == 1:
                y_tareas[tareas_banda[0].id] = (y_sup + y_inf) / 2
            else:
                margen = 0.50
                y0 = y_sup - margen
                for i, tarea in enumerate(tareas_banda):
                    y_tareas[tarea.id] = y0 - i * separacion_fila

        y_actual = y_inf - separacion_banda

    return y_tareas, bandas


def nodos_post_plazo(tareas: list[TareaGantt]) -> set[str]:
    sucesores: dict[str, list[str]] = {t.id: [] for t in tareas}

    for tarea in tareas:
        for dep in tarea.dependencias:
            if dep in sucesores:
                sucesores[dep].append(tarea.id)

    hito_fin = next((t for t in tareas if t.es_fin_plazo), None)
    if hito_fin is None:
        return set()

    post: set[str] = set()
    pila = list(sucesores[hito_fin.id])

    while pila:
        tid = pila.pop()
        if tid in post:
            continue
        post.add(tid)
        pila.extend(sucesores.get(tid, []))

    return post


def es_critica(tarea: TareaGantt, holguras: dict[str, float]) -> bool:
    return holguras.get(tarea.id, 1.0) <= 0


def color_tarea(tarea: TareaGantt, holguras: dict[str, float], post: set[str]) -> tuple[str, str]:
    if tarea.id in post or tarea.estacional:
        return COLOR_POST, COLOR_POST_BORDE
    if tarea.es_fin_plazo:
        return COLOR_FIN_PLAZO, COLOR_HITO_BORDE
    if tarea.tipo == "hito":
        return COLOR_HITO, COLOR_HITO_BORDE
    if es_critica(tarea, holguras):
        return COLOR_CRITICO, COLOR_CRITICO_BORDE
    return COLOR_NORMAL, COLOR_NORMAL_BORDE


def wrap(texto: str, ancho: int) -> str:
    lineas: list[str] = []

    for linea in texto.split("\n"):
        lineas.extend(textwrap.wrap(linea, width=ancho) or [""])

    return "\n".join(lineas)


def dibujar_fondo_bandas(
    ax, bandas: dict[str, tuple[float, float]], x_min: float, x_max: float
) -> None:
    """
    Dibuja los fondos de color de las bandas en toda la zona temporal.
    Solo fondos y separadores horizontales, sin texto.
    """
    for clave, _, _, color in BANDAS:
        y_inf, y_sup = bandas[clave]
        ax.axhspan(y_inf, y_sup, xmin=0, xmax=1, color=color, zorder=0)
        ax.plot([x_min, x_max], [y_inf, y_inf], color="#BDBDBD", lw=0.7, zorder=1)


def dibujar_panel_bandas(
    ax, bandas: dict[str, tuple[float, float]]
) -> None:
    """
    Dibuja el panel izquierdo fijo: número, título y descripción de cada banda.
    Todo el contenido se posiciona entre x=0 y x=ANCHO_PANEL.
    Sin eje temporal, sin grid, sin flechas.
    """
    for idx, (clave, titulo, descripcion, _) in enumerate(BANDAS, start=1):
        y_inf, y_sup = bandas[clave]
        y_centro = (y_inf + y_sup) / 2

        ax.text(
            1.2,
            y_centro,
            str(idx),
            ha="center",
            va="center",
            fontsize=12,
            fontweight="bold",
            color="#333333",
            zorder=2,
        )
        ax.text(
            2.5,
            y_centro + 0.28,
            titulo,
            ha="left",
            va="center",
            fontsize=9,
            fontweight="bold",
            color=COLOR_TEXTO,
            zorder=2,
        )
        ax.text(
            2.5,
            y_centro - 0.35,
            descripcion,
            ha="left",
            va="center",
            fontsize=7,
            color=COLOR_TEXTO,
            zorder=2,
        )

    ax.axvline(ANCHO_PANEL, color="#BDBDBD", lw=1.0, zorder=3)


def dibujar_barra(
    ax,
    tarea: TareaGantt,
    x0: float,
    x1: float,
    y: float,
    color: str,
    borde: str,
) -> None:
    """
    Dibuja la barra de una tarea con texto adaptativo según ancho disponible.
    x0 y x1 ya incluyen la transformación ANCHO_PANEL + días_relativos.
    """
    alto = 0.56
    ancho_real = x1 - x0

    nombre = tarea.nombre_corto or tarea.nombre
    dur = f"{tarea.duracion_dias:.1f} d" if tarea.duracion_dias else ""

    if ancho_real < 3.0:
        texto = nombre
    elif ancho_real < 8.0:
        texto = f"{nombre}\n{dur}" if dur else nombre
    else:
        texto = wrap(f"{nombre}\n{dur}", 18) if dur else wrap(nombre, 18)

    ancho_draw = max(ancho_real, 3.5)

    rect = FancyBboxPatch(
        (x0, y - alto / 2),
        ancho_draw,
        alto,
        boxstyle="round,pad=0.025,rounding_size=0.04",
        facecolor=color,
        edgecolor=borde,
        linewidth=1.6 if borde == COLOR_CRITICO_BORDE else 1.0,
        zorder=5,
    )
    ax.add_patch(rect)

    ax.text(
        x0 + ancho_draw / 2,
        y,
        texto,
        ha="center",
        va="center",
        fontsize=5.8,
        fontweight="bold" if borde == COLOR_CRITICO_BORDE else "normal",
        color=COLOR_TEXTO,
        zorder=6,
        clip_on=True,
    )


def dibujar_hito(
    ax,
    tarea: TareaGantt,
    x: float,
    y: float,
    color: str,
    borde: str,
) -> None:
    """
    Dibuja el rombo de un hito y su nombre corto a la derecha.
    x ya incluye la transformación ANCHO_PANEL + días_relativos.
    """
    radio = 0.50

    rombo = mpatches.RegularPolygon(
        (x, y),
        numVertices=4,
        radius=radio,
        orientation=0.785398,
        facecolor=color,
        edgecolor=borde,
        linewidth=1.2,
        zorder=7,
    )
    ax.add_patch(rombo)

    ax.text(
        x,
        y,
        tarea.id,
        ha="center",
        va="center",
        fontsize=6.5,
        fontweight="bold",
        color=COLOR_TEXTO,
        zorder=8,
    )

    if tarea.nombre_corto:
        ax.text(
            x + radio + 0.35,
            y,
            tarea.nombre_corto,
            ha="left",
            va="center",
            fontsize=6.0,
            color=COLOR_TEXTO,
            zorder=8,
        )


def dibujar_entregables(ax, tarea: TareaGantt, x: float, y: float, idx: int) -> None:
    if not tarea.entregables:
        return

    for i, entregable in enumerate(tarea.entregables):
        x_box = x + 0.75
        y_box = y - 0.48 - i * 0.32

        ax.plot(
            [x, x_box - 0.08],
            [y - 0.22, y_box],
            color=COLOR_DOC,
            lw=0.55,
            linestyle="--",
            zorder=4,
        )

        ax.text(
            x_box,
            y_box,
            entregable,
            ha="left",
            va="center",
            fontsize=5.5,
            color=COLOR_TEXTO,
            bbox={
                "boxstyle": "round,pad=0.12",
                "facecolor": "#FFFFFF",
                "edgecolor": COLOR_DOC,
                "linewidth": 0.55,
            },
            zorder=9,
        )


def dibujar_dependencias(
    ax,
    tareas: list[TareaGantt],
    y_tareas: dict[str, float],
    base: date,
    holguras: dict[str, float],
) -> None:
    tarea_map = {t.id: t for t in tareas}

    for tarea in tareas:
        if tarea.id not in y_tareas:
            continue

        x_dest = x_plot(tarea.fecha_inicio, base)
        y_dest = y_tareas[tarea.id]

        for dep_id in tarea.dependencias:
            origen = tarea_map.get(dep_id)
            if origen is None or origen.id not in y_tareas:
                continue

            x_orig = x_plot(origen.fecha_fin, base)
            y_orig = y_tareas[origen.id]

            critica = es_critica(tarea, holguras) and es_critica(origen, holguras)
            color = COLOR_DEP_CRITICA if critica else COLOR_DEP
            lw = 1.05 if critica else 0.55

            ax.annotate(
                "",
                xy=(x_dest, y_dest),
                xytext=(x_orig, y_orig),
                arrowprops={
                    "arrowstyle": "->",
                    "color": color,
                    "lw": lw,
                    "shrinkA": 10,
                    "shrinkB": 10,
                    "connectionstyle": "angle3,angleA=0,angleB=90",
                },
                zorder=3,
            )


def dibujar_eje(ax, x_min: float, x_max: float, y_min: float) -> None:
    """
    Dibuja el eje temporal con etiquetas de días hábiles relativos a H1.
    Las marcas se posicionan en x = ANCHO_PANEL + día_relativo.
    """
    ax.set_xlim(x_min, x_max)
    ax.set_xlabel(
        "DÍAS HÁBILES DESDE EL INICIO DE OBRA (H1)",
        fontsize=9,
        fontweight="bold",
    )

    dia_min_rel = int((x_min - ANCHO_PANEL) // 10) * 10
    dia_max_rel = int(x_max - ANCHO_PANEL) + 1

    ticks = [ANCHO_PANEL + d for d in range(dia_min_rel, dia_max_rel, 10)]
    labels = [str(d) for d in range(dia_min_rel, dia_max_rel, 10)]

    ax.set_xticks(ticks)
    ax.set_xticklabels(labels)
    ax.tick_params(axis="x", labelsize=8)

    for d in range(dia_min_rel, dia_max_rel, 10):
        x = ANCHO_PANEL + d
        if x >= ANCHO_PANEL:
            ax.axvline(x, color=COLOR_GRID, linestyle="--", lw=0.55, zorder=1)

    ax.axvline(ANCHO_PANEL, color="#333333", linewidth=1.1, zorder=2)
    ax.text(
        ANCHO_PANEL,
        y_min + 0.15,
        "H1",
        ha="center",
        va="bottom",
        fontsize=7,
        fontweight="bold",
        color="#333333",
        zorder=30,
    )


def dibujar_leyenda(ax, x: float, y: float) -> None:
    def caja(titulo: str, y_top: float, alto: float) -> None:
        rect = FancyBboxPatch(
            (x, y_top - alto),
            11.2,
            alto,
            boxstyle="round,pad=0.28,rounding_size=0.06",
            facecolor="#FFFFFF",
            edgecolor="#555555",
            linewidth=0.8,
            zorder=20,
        )
        ax.add_patch(rect)

        ax.text(
            x + 0.45,
            y_top - 0.38,
            titulo,
            ha="left",
            va="top",
            fontsize=7.5,
            fontweight="bold",
            zorder=21,
        )

    caja("LEYENDA GRÁFICA", y, 3.45)
    y0 = y - 0.9

    rect = FancyBboxPatch(
        (x + 0.45, y0 - 0.11),
        0.85,
        0.22,
        boxstyle="round,pad=0.03,rounding_size=0.03",
        facecolor=COLOR_NORMAL,
        edgecolor=COLOR_NORMAL_BORDE,
        linewidth=0.8,
        zorder=21,
    )
    ax.add_patch(rect)
    ax.text(x + 1.65, y0, "Tarea", va="center", fontsize=6.6, zorder=21)

    y0 -= 0.43
    rombo = mpatches.RegularPolygon(
        (x + 0.88, y0),
        4,
        radius=0.18,
        orientation=0.785398,
        facecolor=COLOR_HITO,
        edgecolor=COLOR_HITO_BORDE,
        linewidth=0.8,
        zorder=21,
    )
    ax.add_patch(rombo)
    ax.text(x + 1.65, y0, "Hito contractual", va="center", fontsize=6.6, zorder=21)

    y0 -= 0.43
    entregable_box = FancyBboxPatch(
        (x + 0.45, y0 - 0.11),
        0.85,
        0.22,
        boxstyle="round,pad=0.03",
        facecolor="#FFFFFF",
        edgecolor=COLOR_DOC,
        linewidth=0.7,
        zorder=21,
    )
    ax.add_patch(entregable_box)
    ax.text(x + 1.65, y0, "Entregable", va="center", fontsize=6.6, zorder=21)

    y0 -= 0.43
    ax.annotate(
        "",
        xy=(x + 1.3, y0),
        xytext=(x + 0.45, y0),
        arrowprops={"arrowstyle": "->", "lw": 0.8, "color": COLOR_DEP},
        zorder=21,
    )
    ax.text(x + 1.65, y0, "Dependencia temporal", va="center", fontsize=6.6, zorder=21)

    y0 -= 0.43
    ax.plot([x + 0.45, x + 1.3], [y0, y0], "--", lw=0.8, color=COLOR_DOC, zorder=21)
    ax.text(x + 1.65, y0, "Relación documental", va="center", fontsize=6.6, zorder=21)

    y2 = y - 4.35
    caja("LEYENDA DE COLORES", y2, 3.55)

    items = [
        ("Camino crítico", COLOR_CRITICO, COLOR_CRITICO_BORDE),
        ("Tarea con holgura", COLOR_NORMAL, COLOR_NORMAL_BORDE),
        ("Hito contractual", COLOR_HITO, COLOR_HITO_BORDE),
        ("Fin plazo contractual", COLOR_FIN_PLAZO, COLOR_HITO_BORDE),
        ("Actividad diferida / post-plazo", COLOR_POST, COLOR_POST_BORDE),
    ]

    yc = y2 - 0.9
    for texto, face, edge in items:
        rect = FancyBboxPatch(
            (x + 0.45, yc - 0.11),
            0.85,
            0.22,
            boxstyle="round,pad=0.03,rounding_size=0.03",
            facecolor=face,
            edgecolor=edge,
            linewidth=0.8,
            zorder=21,
        )
        ax.add_patch(rect)
        ax.text(x + 1.65, yc, texto, va="center", fontsize=6.6, zorder=21)
        yc -= 0.44

    y3 = y2 - 4.25
    caja("LEYENDA METODOLÓGICA", y3, 3.75)

    texto = (
        "Las bandas agrupan las actividades\n"
        "por fase operativa.\n"
        "El eje X representa días hábiles\n"
        "desde el inicio de obra H1.\n"
        "La posición temporal es global\n"
        "y común a todas las bandas."
    )

    ax.text(
        x + 0.45,
        y3 - 0.9,
        texto,
        ha="left",
        va="top",
        fontsize=6.5,
        linespacing=1.25,
        zorder=21,
    )


def generar_diagrama_red(planificacion: PlanificacionProyecto) -> bytes:
    """
    Genera el diagrama temporal de implantación como PNG.
    """
    tareas = planificacion.tareas
    if not tareas:
        raise ValueError("La planificación no contiene tareas.")

    base = fecha_base(tareas)
    holguras = calcular_holguras(tareas)
    post = nodos_post_plazo(tareas)

    y_tareas, bandas = calcular_y(tareas)

    x_max_tareas = max(x_plot(t.fecha_fin, base) for t in tareas if t.fecha_fin)
    x_min_tareas = min(x_plot(t.fecha_inicio, base) for t in tareas if t.fecha_inicio)
    x_min = min(x_min_tareas - 2.0, ANCHO_PANEL - 20.0)
    x_max = x_max_tareas + 18.0

    y_min = min(y_inf for y_inf, _ in bandas.values()) - 0.6
    y_max = max(y_sup for _, y_sup in bandas.values()) + 0.6

    fig, ax = plt.subplots(figsize=(24, 14))

    dibujar_fondo_bandas(ax, bandas, x_min, x_max)
    dibujar_panel_bandas(ax, bandas)
    dibujar_eje(ax, x_min, x_max, y_min)
    dibujar_dependencias(ax, tareas, y_tareas, base, holguras)

    for idx, tarea in enumerate(tareas):
        if tarea.id not in y_tareas:
            continue

        y = y_tareas[tarea.id]
        x0 = x_plot(tarea.fecha_inicio, base)
        x1 = x_plot(tarea.fecha_fin, base)
        color, borde = color_tarea(tarea, holguras, post)

        if tarea.tipo == "hito":
            dibujar_hito(ax, tarea, x0, y, color, borde)
            dibujar_entregables(ax, tarea, x0 + 0.4, y, idx)
        else:
            dibujar_barra(ax, tarea, x0, x1, y, color, borde)
            dibujar_entregables(ax, tarea, x1, y, idx)

    dibujar_leyenda(ax, x_max - 11.8, y_max - 0.6)

    ax.set_title(
        "DIAGRAMA TEMPORAL DE IMPLANTACIÓN, DEPENDENCIAS PRINCIPALES Y CAMINO CRÍTICO",
        fontsize=15,
        fontweight="bold",
        pad=18,
    )

    ax.set_ylim(y_min, y_max)
    ax.set_yticks([])
    ax.spines["left"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["top"].set_visible(False)

    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=220, bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)

    return buffer.getvalue()
