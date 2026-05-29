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

BANDAS = [
    (
        "Inicio / documentación contractual",
        "INICIO / DOCUMENTACIÓN\nCONTRACTUAL",
        "Documentación previa,\nplanes de obra e inicio formal.",
        "#EAF4FF",
    ),
    (
        "Auditorías técnicas",
        "AUDITORÍAS TÉCNICAS",
        "Caracterización inicial de UTs\ny subsistemas afectados.",
        "#F3FAEA",
    ),
    (
        "Validación de inventarios y tipologías",
        "VALIDACIÓN DE INVENTARIOS\nY TIPOLOGÍAS",
        "Revisión STI, clasificación de\ndeficiencias y autorización técnica.",
        "#FFF6DD",
    ),
    (
        "Actuaciones habilitadoras sistémicas",
        "ACTUACIONES HABILITADORAS\nSISTÉMICAS",
        "Trabajos que condicionan pruebas\nposteriores y estabilidad del sistema.",
        "#F4E9FF",
    ),
    (
        "Ejecución correctiva por bloques",
        "EJECUCIÓN CORRECTIVA\nPOR BLOQUES",
        "Correcciones principales de control,\nventilación, hidráulica y accesos.",
        "#FFF0E6",
    ),
    (
        "Validación funcional por bloques",
        "VALIDACIÓN FUNCIONAL\nPOR BLOQUES",
        "Comprobación funcional de cada\nbloque antes de puesta en servicio.",
        "#ECF8EE",
    ),
    (
        "Puesta en servicio funcional",
        "PUESTA EN SERVICIO\nFUNCIONAL",
        "Pruebas integradas y verificación\noperativa del sistema reformado.",
        "#E9F5FF",
    ),
    (
        "Cierre contractual",
        "CIERRE CONTRACTUAL",
        "Recepción operativa y fin del\nplazo contractual.",
        "#FFF7D8",
    ),
    (
        "Validaciones estacionales diferidas",
        "VALIDACIONES ESTACIONALES\nDIFERIDAS",
        "Ensayos térmicos en régimen real\nde calefacción y refrigeración.",
        "#F2F2F2",
    ),
    (
        "Proyecto As Built",
        "PROYECTO AS BUILT",
        "Documentación final, planos\nactualizados y cierre documental.",
        "#F7F7F7",
    ),
]


def _fecha_base(tareas: list[TareaGantt]) -> date:
    fechas = [t.fecha_inicio for t in tareas if t.fecha_inicio is not None]
    if not fechas:
        raise ValueError("La planificación no contiene fechas calculadas.")
    return min(fechas)


def _x_inicio(tarea: TareaGantt, base: date) -> float:
    if tarea.fecha_inicio is None:
        return 0.0
    return float(dias_habiles_entre(base, tarea.fecha_inicio))


def _x_fin(tarea: TareaGantt, base: date) -> float:
    if tarea.fecha_fin is None:
        return _x_inicio(tarea, base)
    return float(dias_habiles_entre(base, tarea.fecha_fin))


def _grupo(tarea: TareaGantt) -> str:
    if tarea.grupo_visual:
        return tarea.grupo_visual
    if tarea.estacional:
        return "Validaciones estacionales diferidas"
    if tarea.es_fin_plazo:
        return "Cierre contractual"
    return "Inicio / documentación contractual"


def _agrupar_tareas(tareas: list[TareaGantt]) -> dict[str, list[TareaGantt]]:
    grupos: dict[str, list[TareaGantt]] = defaultdict(list)
    for tarea in tareas:
        grupos[_grupo(tarea)].append(tarea)

    for tareas_grupo in grupos.values():
        tareas_grupo.sort(
            key=lambda t: (
                t.orden_visual if t.orden_visual is not None else 999,
                t.fecha_inicio or date.min,
                t.id,
            )
        )

    return grupos


def _calcular_y(tareas: list[TareaGantt]) -> tuple[dict[str, float], dict[str, tuple[float, float]]]:
    grupos = _agrupar_tareas(tareas)
    y_tareas: dict[str, float] = {}
    bandas: dict[str, tuple[float, float]] = {}

    y_actual = 0.0
    alto_min_banda = 1.75
    separacion_fila = 0.85
    separacion_banda = 0.25

    for nombre_banda, *_ in BANDAS:
        tareas_banda = grupos.get(nombre_banda, [])
        num_filas = max(len(tareas_banda), 1)
        alto_banda = max(alto_min_banda, num_filas * separacion_fila + 0.55)

        y_sup = y_actual
        y_inf = y_actual - alto_banda
        bandas[nombre_banda] = (y_inf, y_sup)

        if tareas_banda:
            if len(tareas_banda) == 1:
                y_tareas[tareas_banda[0].id] = (y_sup + y_inf) / 2
            else:
                margen = 0.45
                y0 = y_sup - margen
                for i, tarea in enumerate(tareas_banda):
                    y_tareas[tarea.id] = y0 - i * separacion_fila

        y_actual = y_inf - separacion_banda

    return y_tareas, bandas


def _nodos_post_plazo(tareas: list[TareaGantt]) -> set[str]:
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


def _es_critica(tarea: TareaGantt, holguras: dict[str, float]) -> bool:
    return holguras.get(tarea.id, 1.0) <= 0


def _color_tarea(tarea: TareaGantt, holguras: dict[str, float], post: set[str]) -> tuple[str, str]:
    if tarea.id in post or tarea.estacional:
        return COLOR_POST, COLOR_POST_BORDE
    if tarea.es_fin_plazo:
        return COLOR_FIN_PLAZO, COLOR_HITO_BORDE
    if tarea.tipo == "hito":
        return COLOR_HITO, COLOR_HITO_BORDE
    if _es_critica(tarea, holguras):
        return COLOR_CRITICO, COLOR_CRITICO_BORDE
    return COLOR_NORMAL, COLOR_NORMAL_BORDE


def _texto_tarea(tarea: TareaGantt, duracion: bool = True) -> str:
    nombre = tarea.nombre_corto or tarea.nombre
    if tarea.tipo == "hito":
        return f"{tarea.id}"

    if duracion and tarea.duracion_dias is not None:
        return f"{tarea.id}\n{nombre}\n{tarea.duracion_dias:.1f} d"

    return f"{tarea.id}\n{nombre}"


def _wrap(texto: str, ancho: int) -> str:
    lineas: list[str] = []
    for linea in texto.split("\n"):
        lineas.extend(textwrap.wrap(linea, width=ancho) or [""])
    return "\n".join(lineas)


def _dibujar_bandas(ax, bandas: dict[str, tuple[float, float]], x_min: float, x_max: float) -> None:
    ancho_panel = 18.0

    for idx, (clave, titulo, descripcion, color) in enumerate(BANDAS, start=1):
        y_inf, y_sup = bandas[clave]

        ax.axhspan(y_inf, y_sup, xmin=0, xmax=1, color=color, zorder=0)
        ax.plot([x_min, x_max], [y_inf, y_inf], color="#BDBDBD", lw=0.7, zorder=1)

        ax.text(
            x_min + 0.8,
            (y_inf + y_sup) / 2,
            str(idx),
            ha="center",
            va="center",
            fontsize=12,
            fontweight="bold",
            color="#333333",
            bbox={
                "boxstyle": "round,pad=0.18",
                "facecolor": "#FFFFFFAA",
                "edgecolor": "none",
            },
            zorder=2,
        )

        ax.text(
            x_min + 2.0,
            (y_inf + y_sup) / 2 + 0.28,
            titulo,
            ha="left",
            va="center",
            fontsize=9,
            fontweight="bold",
            color=COLOR_TEXTO,
            zorder=2,
        )

        ax.text(
            x_min + 2.0,
            (y_inf + y_sup) / 2 - 0.35,
            descripcion,
            ha="left",
            va="center",
            fontsize=7,
            color=COLOR_TEXTO,
            zorder=2,
        )

    ax.axvline(x_min + ancho_panel, color="#BDBDBD", lw=0.8, zorder=2)


def _dibujar_barra(
    ax,
    tarea: TareaGantt,
    x0: float,
    x1: float,
    y: float,
    color: str,
    borde: str,
) -> tuple[float, float, float, float]:
    alto = 0.42
    ancho = max(x1 - x0, 1.8)

    rect = FancyBboxPatch(
        (x0, y - alto / 2),
        ancho,
        alto,
        boxstyle="round,pad=0.02,rounding_size=0.03",
        facecolor=color,
        edgecolor=borde,
        linewidth=1.4 if borde == COLOR_CRITICO_BORDE else 0.9,
        zorder=5,
    )
    ax.add_patch(rect)

    texto = _texto_tarea(tarea)
    texto = _wrap(texto, 18)

    ax.text(
        x0 + ancho / 2,
        y,
        texto,
        ha="center",
        va="center",
        fontsize=6.4,
        fontweight="bold" if borde == COLOR_CRITICO_BORDE else "normal",
        color=COLOR_TEXTO,
        zorder=6,
    )

    return x0, x1, y - alto / 2, y + alto / 2


def _dibujar_hito(
    ax,
    tarea: TareaGantt,
    x: float,
    y: float,
    color: str,
    borde: str,
) -> tuple[float, float, float, float]:
    radio = 0.48

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
        fontsize=7,
        fontweight="bold",
        color=COLOR_TEXTO,
        zorder=8,
    )

    if tarea.nombre_corto:
        ax.text(
            x,
            y + 0.62,
            tarea.nombre_corto,
            ha="center",
            va="bottom",
            fontsize=6.2,
            color=COLOR_TEXTO,
            zorder=8,
        )

    return x - radio, x + radio, y - radio, y + radio


def _dibujar_entregables(ax, tarea: TareaGantt, x: float, y: float, idx: int) -> None:
    if not tarea.entregables:
        return

    for i, entregable in enumerate(tarea.entregables):
        x_box = x + 0.55 + i * 2.25
        y_box = y - 0.42

        ax.plot(
            [x, x_box],
            [y - 0.18, y_box + 0.12],
            color=COLOR_DOC,
            lw=0.6,
            linestyle="--",
            zorder=4,
        )

        ax.text(
            x_box,
            y_box,
            entregable,
            ha="left",
            va="center",
            fontsize=5.8,
            color=COLOR_TEXTO,
            bbox={
                "boxstyle": "round,pad=0.15",
                "facecolor": "#FFFFFF",
                "edgecolor": COLOR_DOC,
                "linewidth": 0.6,
            },
            zorder=9,
        )


def _dibujar_dependencias(
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

        x_dest = _x_inicio(tarea, base)
        y_dest = y_tareas[tarea.id]

        for dep_id in tarea.dependencias:
            origen = tarea_map.get(dep_id)
            if origen is None or origen.id not in y_tareas:
                continue

            x_orig = _x_fin(origen, base)
            y_orig = y_tareas[origen.id]

            critica = _es_critica(tarea, holguras) and _es_critica(origen, holguras)
            color = COLOR_DEP_CRITICA if critica else COLOR_DEP
            lw = 1.0 if critica else 0.55

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


def _dibujar_eje(ax, x_min: float, x_max: float, y_min: float) -> None:
    ax.set_xlim(x_min, x_max)
    ax.set_xlabel("DÍAS HÁBILES DESDE LA ADJUDICACIÓN", fontsize=9, fontweight="bold")
    ax.set_xticks(range(0, int(x_max) + 1, 10))
    ax.tick_params(axis="x", labelsize=8)

    for x in range(0, int(x_max) + 1, 10):
        ax.axvline(x, color=COLOR_GRID, linestyle="--", lw=0.55, zorder=1)


def _dibujar_leyenda(ax, x: float, y: float) -> None:
    def caja(titulo: str, y_top: float, alto: float) -> None:
        rect = FancyBboxPatch(
            (x, y_top - alto),
            10.0,
            alto,
            boxstyle="round,pad=0.25,rounding_size=0.05",
            facecolor="#FFFFFF",
            edgecolor="#555555",
            linewidth=0.8,
            zorder=20,
        )
        ax.add_patch(rect)
        ax.text(x + 0.45, y_top - 0.35, titulo, ha="left", va="top", fontsize=8, fontweight="bold", zorder=21)

    caja("LEYENDA GRÁFICA", y, 3.0)
    y0 = y - 0.85

    rect = FancyBboxPatch((x + 0.45, y0 - 0.12), 1.0, 0.25, facecolor=COLOR_NORMAL, edgecolor=COLOR_NORMAL_BORDE, zorder=21)
    ax.add_patch(rect)
    ax.text(x + 1.8, y0, "Tarea", va="center", fontsize=7, zorder=21)

    rombo = mpatches.RegularPolygon((x + 0.95, y0 - 0.45), 4, radius=0.22, orientation=0.785398,
                                    facecolor=COLOR_HITO, edgecolor=COLOR_HITO_BORDE, zorder=21)
    ax.add_patch(rombo)
    ax.text(x + 1.8, y0 - 0.45, "Hito contractual", va="center", fontsize=7, zorder=21)

    ax.text(x + 0.45, y0 - 0.9, "▭", fontsize=10, va="center", zorder=21)
    ax.text(x + 1.8, y0 - 0.9, "Entregable", va="center", fontsize=7, zorder=21)

    ax.annotate("", xy=(x + 1.4, y0 - 1.35), xytext=(x + 0.45, y0 - 1.35),
                arrowprops={"arrowstyle": "->", "lw": 0.8, "color": COLOR_DEP}, zorder=21)
    ax.text(x + 1.8, y0 - 1.35, "Dependencia temporal", va="center", fontsize=7, zorder=21)

    ax.plot([x + 0.45, x + 1.4], [y0 - 1.8, y0 - 1.8], "--", lw=0.8, color=COLOR_DOC, zorder=21)
    ax.text(x + 1.8, y0 - 1.8, "Relación documental", va="center", fontsize=7, zorder=21)

    y2 = y - 3.8
    caja("LEYENDA DE COLORES", y2, 3.2)
    items = [
        ("Camino crítico", COLOR_CRITICO, COLOR_CRITICO_BORDE),
        ("Tarea con holgura", COLOR_NORMAL, COLOR_NORMAL_BORDE),
        ("Hito contractual", COLOR_HITO, COLOR_HITO_BORDE),
        ("Fin plazo contractual", COLOR_FIN_PLAZO, COLOR_HITO_BORDE),
        ("Actividad diferida /\npost-plazo contractual", COLOR_POST, COLOR_POST_BORDE),
    ]

    yc = y2 - 0.85
    for texto, face, edge in items:
        rect = FancyBboxPatch((x + 0.45, yc - 0.12), 1.0, 0.25, facecolor=face, edgecolor=edge, zorder=21)
        ax.add_patch(rect)
        ax.text(x + 1.8, yc, texto, va="center", fontsize=7, zorder=21)
        yc -= 0.48

    y3 = y2 - 4.0
    caja("LEYENDA METODOLÓGICA", y3, 3.0)
    texto = (
        "Las bandas agrupan las\n"
        "actividades por fase\n"
        "operativa. El eje X\n"
        "representa días hábiles\n"
        "desde la adjudicación.\n"
        "La posición temporal es\n"
        "global y común a todas\n"
        "las bandas."
    )
    ax.text(x + 0.45, y3 - 0.85, texto, ha="left", va="top", fontsize=7, zorder=21)


def generar_diagrama_red(planificacion: PlanificacionProyecto) -> bytes:
    """
    Genera el diagrama temporal de implantación como PNG.
    """
    tareas = planificacion.tareas
    if not tareas:
        raise ValueError("La planificación no contiene tareas.")

    base = _fecha_base(tareas)
    holguras = calcular_holguras(tareas)
    post = _nodos_post_plazo(tareas)

    y_tareas, bandas = _calcular_y(tareas)

    x_max_tareas = max(_x_fin(t, base) for t in tareas)
    x_min = -18.0
    x_max = x_max_tareas + 14.0

    y_min = min(y_inf for y_inf, _ in bandas.values()) - 0.6
    y_max = max(y_sup for _, y_sup in bandas.values()) + 0.4

    fig, ax = plt.subplots(figsize=(24, 14))

    _dibujar_bandas(ax, bandas, x_min, x_max)
    _dibujar_eje(ax, x_min, x_max, y_min)
    _dibujar_dependencias(ax, tareas, y_tareas, base, holguras)

    for idx, tarea in enumerate(tareas):
        if tarea.id not in y_tareas:
            continue

        y = y_tareas[tarea.id]
        x0 = _x_inicio(tarea, base)
        x1 = _x_fin(tarea, base)
        color, borde = _color_tarea(tarea, holguras, post)

        if tarea.tipo == "hito":
            _dibujar_hito(ax, tarea, x0, y, color, borde)
            _dibujar_entregables(ax, tarea, x0 + 0.4, y, idx)
        else:
            _dibujar_barra(ax, tarea, x0, x1, y, color, borde)
            _dibujar_entregables(ax, tarea, x1, y, idx)

    _dibujar_leyenda(ax, x_max - 10.7, y_max - 0.2)

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