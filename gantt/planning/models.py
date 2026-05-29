"""
Modelos Pydantic que representan la planificación temporal del proyecto.

Responsabilidad: definir los tipos de datos del cronograma y cargar
la configuración desde el archivo planificacion.yaml.
"""

from datetime import date
from pathlib import Path

import yaml
from pydantic import BaseModel, Field


class ParametrosProyecto(BaseModel):
    """
    Parámetros globales del proyecto extraídos del yaml.
    Define el calendario laboral y el plazo contractual.
    """

    nombre: str
    fecha_inicio: date
    horas_dia: int
    dias_semana: int
    plazo_contractual_dias: int


class EscenarioRecursos(BaseModel):
    """
    Escenario de recursos definido en el yaml.
    Mapea cada código de recurso MO al número de operarios asignados.
    """

    nombre: str
    operarios_por_recurso: dict[str, int]  # clave: codigo MO, valor: operarios


class TareaGantt(BaseModel):
    """
    Tarea o hito del cronograma del proyecto.

    Campos de definición (vienen del yaml, obligatorios):
        id, nombre, tipo, capitulos_bc3, duracion_dias_fija,
        dependencias, estacional, es_fin_plazo.

    Campos calculados (los rellena calculator, None hasta que se calculan):
        duracion_dias, fecha_inicio, fecha_fin,
        perfil_limitante, horas_mo_total, horas_por_recurso.
    """

    # Campos de definición — vienen del yaml
    id: str
    nombre: str
    tipo: str                           # "tarea" o "hito"
    capitulos_bc3: list[str]            # códigos BC3 asociados
    duracion_dias_fija: int | None = None  # None si la duración se calcula desde BC3
    dependencias: list[str]             # IDs de tareas predecesoras
    estacional: bool = False
    es_fin_plazo: bool = False          # True en el hito que marca el fin del plazo contractual

    # Campos opcionales de representación gráfica
    grupo_visual: str | None = None     # banda funcional del diagrama de red
    nombre_corto: str | None = None     # etiqueta compacta para el nodo
    entregables: list[str] = Field(default_factory=list)
    orden_visual: int | None = None     # ajuste manual dentro de la banda

    # Campos calculados — None hasta que calculator los rellena
    duracion_dias: float | None = None
    fecha_inicio: date | None = None
    fecha_fin: date | None = None
    perfil_limitante: str | None = None  # descripción del recurso MO más limitante
    horas_mo_total: float | None = None
    horas_por_recurso: dict[str, float] = Field(default_factory=dict)


class PlanificacionProyecto(BaseModel):
    """
    Planificación completa del proyecto para un escenario de recursos dado.
    Contiene los parámetros del proyecto, el escenario aplicado
    y la lista ordenada de tareas con sus fechas calculadas.

    tareas_originales conserva las tareas sin calcular tal como vienen del yaml,
    necesarias para recalcular la planificación con distintos escenarios de recursos
    sin mutar el estado de la planificación de referencia.
    """

    parametros: ParametrosProyecto
    escenario: EscenarioRecursos
    tareas: list[TareaGantt]
    tareas_originales: list[TareaGantt] = Field(default_factory=list)


# Loader que preserva como string los escalares YAML con cero inicial (ej: 07, 08).
# yaml.safe_load convertiría '07' → int 7; este loader devuelve '07' como str.
# Los enteros sin cero inicial (8, 15, 65...) se cargan normalmente como int.
class _StrPreservingLoader(yaml.SafeLoader):
    pass

def _construct_int_preserving_zeros(loader: yaml.SafeLoader, node) -> int | str:
    value = loader.construct_scalar(node)
    if value.startswith('0') and len(value) > 1 and value[1:].isdigit():
        return value  # preservar '07', '08', etc. como string
    return loader.construct_yaml_int(node)

_StrPreservingLoader.add_constructor(
    'tag:yaml.org,2002:int',
    _construct_int_preserving_zeros,
)


def cargar_planificacion_yaml(
    yaml_path: Path,
) -> tuple[ParametrosProyecto, list[EscenarioRecursos], list[TareaGantt]]:
    """
    Lee el archivo planificacion.yaml y construye los objetos de planificación.
    Usa un loader personalizado para preservar claves con cero inicial (ej: '07').

    Parámetros
    ----------
    yaml_path : Path
        Ruta al archivo planificacion.yaml del proyecto.

    Retorna
    -------
    tuple[ParametrosProyecto, list[EscenarioRecursos], list[TareaGantt]]
        Parámetros del proyecto, lista de escenarios y lista de tareas.
    """
    with open(yaml_path, encoding='utf-8') as f:
        data = yaml.load(f, Loader=_StrPreservingLoader)

    proy = data['proyecto']
    parametros = ParametrosProyecto(
        nombre=proy['nombre'],
        fecha_inicio=date.fromisoformat(proy['fecha_inicio']),
        horas_dia=proy['horas_dia'],
        dias_semana=proy['dias_semana'],
        plazo_contractual_dias=proy['plazo_contractual_dias'],
    )

    escenarios = [
        EscenarioRecursos(nombre=nombre, operarios_por_recurso=operarios)
        for nombre, operarios in data.get('escenarios', {}).items()
    ]

    # Fuerza las claves a str antes de iterar: yaml.safe_load convierte '07' → int 7
    tareas_raw = {str(k): v for k, v in data.get('tareas', {}).items()}

    tareas = [
        TareaGantt(
            id=tarea_id,
            nombre=tarea_data['nombre'],
            tipo=tarea_data['tipo'],
            capitulos_bc3=tarea_data.get('capitulos_bc3', []),
            duracion_dias_fija=tarea_data.get('duracion_dias_fija'),
            dependencias=[str(d) for d in tarea_data.get('dependencias', [])],
            estacional=tarea_data.get('estacional', False),
            es_fin_plazo=tarea_data.get('es_fin_plazo', False),
            grupo_visual=tarea_data.get('grupo_visual'),
            nombre_corto=tarea_data.get('nombre_corto'),
            entregables=tarea_data.get('entregables', []),
            orden_visual=tarea_data.get('orden_visual'),
        )
        for tarea_id, tarea_data in tareas_raw.items()
    ]

    return parametros, escenarios, tareas


if __name__ == '__main__':
    import sys
    from pathlib import Path

    project_name = sys.argv[1] if len(sys.argv) > 1 else 'Complexe-Balear'
    input_dir = Path('projects') / project_name / 'input'

    yaml_files = list(input_dir.glob('planificacion.yaml'))
    if not yaml_files:
        raise FileNotFoundError(f'No se encontró planificacion.yaml en {input_dir}')

    parametros, escenarios, tareas = cargar_planificacion_yaml(yaml_files[0])

    print(f'Proyecto: {parametros.nombre}')
    print(f'Fecha inicio: {parametros.fecha_inicio}')
    print(f'Plazo contractual: {parametros.plazo_contractual_dias} días hábiles')
    print(f'Escenarios: {[e.nombre for e in escenarios]}')
    print(f'Tareas cargadas: {len(tareas)}')
    print()
    for tarea in tareas:
        dep_str = ', '.join(tarea.dependencias) if tarea.dependencias else '—'
        bc3_str = ', '.join(tarea.capitulos_bc3) if tarea.capitulos_bc3 else 'duración fija'
        print(f'  {tarea.id:<12} [{tarea.tipo:<5}] {tarea.nombre[:40]:<40} '
              f'dep: {dep_str:<20} bc3: {bc3_str}')
        