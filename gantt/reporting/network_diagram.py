"""
Módulo responsable de generar el diagrama de red del proyecto como imagen PNG.

Responsabilidad: construir un grafo dirigido de dependencias entre tareas
y renderizarlo en memoria para embeber en el Excel.
"""

import io
from collections import defaultdict, deque

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import networkx as nx

from gantt.planning.analyser import calcular_holguras
from gantt.planning.models import PlanificacionProyecto, TareaGantt
from gantt.reporting.styles import (
    COLOR_AMARILLO, COLOR_GRIS_NEUTRO, COLOR_ROJO, COLOR_VERDE,
)

# Azul para tareas paralelas — específico del diagrama, sin constante global
_COLOR_PARALELA = '#6BB5FF'


def calcular_layout(tareas: list[TareaGantt], G: nx.DiGraph) -> dict:
    """
    Calcula posiciones jerárquicas izquierda a derecha por profundidad topológica.
    El eje X es la profundidad; el eje Y distribuye nodos del mismo nivel.
    """
    profundidad: dict[str, int] = {}
    in_degree = {t.id: len(t.dependencias) for t in tareas}
    queue = deque([t.id for t in tareas if in_degree[t.id] == 0])

    for tid in queue:
        profundidad[tid] = 0

    orden = []
    while queue:
        nid = queue.popleft()
        orden.append(nid)
        for suc in G.successors(nid):
            in_degree[suc] -= 1
            prof_nueva = profundidad[nid] + 1
            profundidad[suc] = max(profundidad.get(suc, 0), prof_nueva)
            if in_degree[suc] == 0:
                queue.append(suc)

    nivel_nodos: dict[int, list[str]] = defaultdict(list)
    for tid, prof in profundidad.items():
        nivel_nodos[prof].append(tid)

    pos = {}
    for prof, nodos_nivel in nivel_nodos.items():
        n = len(nodos_nivel)
        for j, tid in enumerate(nodos_nivel):
            pos[tid] = (prof * 4, -(j - n / 2) * 2)

    return pos


def generar_diagrama_red(planificacion: PlanificacionProyecto) -> bytes:
    """
    Genera el diagrama de red del proyecto como imagen PNG en memoria.
    Colores definidos en styles.py; azul paralela es específico del diagrama.
    """
    tareas    = planificacion.tareas
    tarea_map = {t.id: t for t in tareas}
    holguras  = calcular_holguras(tareas)

    hito_fin = next((t for t in tareas if t.es_fin_plazo), None)

    graph = nx.DiGraph()
    for task in tareas:
        graph.add_node(task.id)
    for task in tareas:
        for dep_id in task.dependencias:
            if dep_id in tarea_map:
                graph.add_edge(dep_id, task.id)

    pos = calcular_layout(tareas, graph)

    # Identificar nodos posteriores al fin de plazo via BFS en el grafo de sucesores
    nodos_post_plazo: set[str] = set()
    if hito_fin:
        cola = list(graph.successors(hito_fin.id))
        visitados: set[str] = set()
        while cola:
            nid = cola.pop()
            if nid in visitados:
                continue
            visitados.add(nid)
            nodos_post_plazo.add(nid)
            cola.extend(graph.successors(nid))

    node_colors = []
    for nid in graph.nodes():
        tarea = tarea_map[nid]
        if nid in nodos_post_plazo:
            node_colors.append(f'#{COLOR_GRIS_NEUTRO}')   # fuera de plazo contractual
        elif tarea.es_fin_plazo:
            node_colors.append(f'#{COLOR_VERDE}')          # hito fin de plazo
        elif tarea.tipo == 'hito':
            node_colors.append(f'#{COLOR_AMARILLO}')       # hito normal
        elif holguras.get(nid, 1) <= 0:
            node_colors.append(f'#{COLOR_ROJO}')           # camino crítico
        else:
            node_colors.append(_COLOR_PARALELA)            # tarea paralela

    edge_colors = []
    for u, v in graph.edges():
        es_post = (u in nodos_post_plazo or v in nodos_post_plazo
                   or (hito_fin is not None and u == hito_fin.id))
        if es_post:
            edge_colors.append(f'#{COLOR_GRIS_NEUTRO}')
        elif holguras.get(u, 1) <= 0 and holguras.get(v, 1) <= 0:
            edge_colors.append(f'#{COLOR_ROJO}')
        else:
            edge_colors.append('#555555')

    etiquetas = {}
    for t in tareas:
        if t.tipo == 'hito':
            etiquetas[t.id] = t.id
        else:
            dias = f'{t.duracion_dias:.1f}d' if t.duracion_dias else '0d'
            etiquetas[t.id] = f'{t.id}\n{dias}'

    fig, ax = plt.subplots(figsize=(22, 12))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')

    nx.draw_networkx_edges(
        graph,
        pos,
        ax=ax,
        edge_color=edge_colors,
        arrows=True,
        arrowsize=15,
        width=1.5,
        alpha=0.8,
        node_size=2000,
    )

    nx.draw_networkx_nodes(
        graph,
        pos,
        ax=ax,
        node_color=node_colors,
        node_size=2000,
        alpha=0.9,
    )

    nx.draw_networkx_labels(
        graph,
        pos,
        etiquetas,
        ax=ax,
        font_size=8,
        font_weight='bold',
    )

    ax.legend(
        handles=[
            mpatches.Patch(color=f'#{COLOR_ROJO}',       label='Camino critico'),
            mpatches.Patch(color=_COLOR_PARALELA,         label='Tarea paralela'),
            mpatches.Patch(color=f'#{COLOR_AMARILLO}',    label='Hito'),
            mpatches.Patch(color=f'#{COLOR_VERDE}',       label='Fin de plazo contractual'),
            mpatches.Patch(color=f'#{COLOR_GRIS_NEUTRO}', label='Fuera de plazo contractual'),
        ],
        loc='upper left',
        fontsize=9,
    )

    ax.set_title(
        f'Diagrama de Red - {planificacion.parametros.nombre}',
        fontsize=13,
        fontweight='bold',
        pad=10,
    )

    ax.axis('off')

    buffer = io.BytesIO()
    plt.savefig(
        buffer,
        format='png',
        dpi=120,
        bbox_inches='tight',
        facecolor='white',
    )
    plt.close(fig)
    buffer.seek(0)

    return buffer.read()
