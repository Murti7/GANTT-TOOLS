"""
Módulo responsable de leer y parsear un archivo BC3 (FIEBDC-3).

Responsabilidad única: transformar el contenido de texto de un archivo .bc3
en un objeto Presupuesto con la estructura de capítulos y partidas correspondiente.
"""

from pathlib import Path

from gantt.bc3.models import Capitulo, LineaDescompuesto, Partida, Presupuesto, ProjectConfig, RecursoElemental, RecursoMO


def read_bc3_records(
    filepath: Path,
) -> tuple[dict[str, dict[str, str]], dict[str, list[tuple[str, float]]], dict[str, str]]:
    """
    Lee el archivo BC3 en una sola pasada y extrae los registros ~C, ~D y ~T.
    Retorna (conceptos, descompuestos, textos) indexados por código.
    El campo 'tipo' del registro ~C se almacena en conceptos[codigo]['tipo'].
    """
    conceptos: dict[str, dict[str, str]] = {}
    descompuestos: dict[str, list[tuple[str, float]]] = {}
    textos: dict[str, str] = {}
    current_parent: str | None = None
    current_items: list[tuple[str, float]] = []
    current_texto: str | None = None
    current_texto_lines: list[str] = []

    with open(filepath, encoding='latin-1') as f:
        for raw in f:
            line = raw.rstrip('\r\n')

            if line.startswith('~C|'):
                if current_parent is not None:
                    descompuestos[current_parent] = current_items
                    current_parent = None
                    current_items = []
                if current_texto is not None:
                    textos[current_texto] = ' '.join(current_texto_lines).strip()
                    current_texto = None
                    current_texto_lines = []
                parts = line.split('|')
                if len(parts) >= 5:
                    codigo = parts[1].strip()
                    conceptos[codigo] = {
                        'unidad':      parts[2].strip(),
                        'descripcion': parts[3].strip(),
                        'precio_str':  parts[4].strip(),
                        'tipo':        parts[6].strip().rstrip('|') if len(parts) >= 7 else '',
                    }

            elif line.startswith('~D|'):
                if current_parent is not None:
                    descompuestos[current_parent] = current_items
                if current_texto is not None:
                    textos[current_texto] = ' '.join(current_texto_lines).strip()
                    current_texto = None
                    current_texto_lines = []
                current_parent = line[3:].strip().rstrip('|').strip()
                current_items = []

            elif line.startswith('~T|'):
                if current_parent is not None:
                    descompuestos[current_parent] = current_items
                    current_parent = None
                    current_items = []
                if current_texto is not None:
                    textos[current_texto] = ' '.join(current_texto_lines).strip()
                parts_t = line[3:].split('|', 1)
                current_texto = parts_t[0].strip()
                current_texto_lines = []
                if len(parts_t) > 1:
                    first = parts_t[1].strip()
                    if first:
                        current_texto_lines.append(first)

            elif current_parent is not None:
                if line.startswith('|') or line.startswith('\\'):
                    content = line.lstrip('|\\')
                    if not content or content == '|':
                        descompuestos[current_parent] = current_items
                        current_parent = None
                        current_items = []
                    else:
                        parts = content.split('\\\\')
                        if len(parts) >= 2:
                            codigo_hijo = parts[0].strip()
                            if codigo_hijo:
                                try:
                                    cantidad = float(parts[1].strip())
                                except ValueError:
                                    cantidad = 0.0
                                current_items.append((codigo_hijo, cantidad))
                else:
                    descompuestos[current_parent] = current_items
                    current_parent = None
                    current_items = []

            elif current_texto is not None:
                if line.startswith('~'):
                    textos[current_texto] = ' '.join(current_texto_lines).strip()
                    current_texto = None
                    current_texto_lines = []
                else:
                    stripped = line.strip()
                    if stripped:
                        current_texto_lines.append(stripped)

    if current_parent is not None:
        descompuestos[current_parent] = current_items
    if current_texto is not None:
        textos[current_texto] = ' '.join(current_texto_lines).strip()

    return conceptos, descompuestos, textos


def expandir_descompuesto(
    codigo: str,
    descompuestos: dict[str, list[tuple[str, float]]],
    pa_codigos: set[str],
    cant_factor: float = 1.0,
    depth: int = 0,
    vistos: frozenset[str] = frozenset(),
) -> dict[str, float]:
    """
    Retorna {codigo_recurso: cantidad_efectiva} expandiendo recursivamente
    las partidas alzadas indicadas en pa_codigos.
    Los costes indirectos (%) dentro de PAs (depth > 0) se omiten.
    Función pública: también la usa presupuesto_exporter.
    """
    if codigo in vistos:
        return {}
    vistos = vistos | {codigo}
    acumulado: dict[str, float] = {}
    for cod_rec, cant in descompuestos.get(codigo, []):
        cant_ef = cant * cant_factor
        if cod_rec in pa_codigos:
            for cod, c in expandir_descompuesto(
                cod_rec, descompuestos, pa_codigos, cant_ef, depth + 1, vistos
            ).items():
                acumulado[cod] = acumulado.get(cod, 0.0) + c
        elif cod_rec.startswith('%') and depth > 0:
            pass  # % dentro de PA: omitir
        else:
            acumulado[cod_rec] = acumulado.get(cod_rec, 0.0) + cant_ef
    return acumulado


def build_partida(
    codigo: str,
    cantidad: float,
    conceptos: dict[str, dict[str, str]],
    descompuestos: dict[str, list[tuple[str, float]]],
    textos: dict[str, str],
    mo_codigos: set[str],
    pa_codigos: set[str],
) -> Partida:
    """
    Construye una Partida. Las lineas_mo se obtienen expandiendo recursivamente
    PAs mediante expandir_descompuesto, filtrando después por mo_codigos.
    """
    concepto = conceptos.get(codigo, {})
    todos = expandir_descompuesto(codigo, descompuestos, pa_codigos)
    lineas_mo = [
        LineaDescompuesto(codigo_recurso=cod, cantidad=round(cant, 6))
        for cod, cant in todos.items()
        if cod in mo_codigos
    ]
    return Partida(
        codigo=codigo,
        descripcion=concepto.get('descripcion', ''),
        descripcion_larga=textos.get(codigo, ''),
        unidad=concepto.get('unidad', ''),
        precio_unitario=float(concepto.get('precio_str', '') or '0'),
        cantidad=cantidad,
        lineas_mo=lineas_mo,
    )


def build_capitulo(
    codigo: str,
    conceptos: dict[str, dict[str, str]],
    descompuestos: dict[str, list[tuple[str, float]]],
    textos: dict[str, str],
    mo_codigos: set[str],
    todos_recursos: set[str],
    pa_codigos: set[str],
) -> Capitulo:
    """
    Construye recursivamente un Capitulo con sus subcapítulos y partidas.
    Ítems terminados en '#' son subcapítulos; los que no están en todos_recursos son partidas.
    """
    concepto = conceptos.get(codigo, {})
    subcapitulos = []
    partidas = []

    for codigo_hijo, cantidad in descompuestos.get(codigo, []):
        if codigo_hijo.endswith('#'):
            subcapitulos.append(
                build_capitulo(codigo_hijo, conceptos, descompuestos, textos,
                               mo_codigos, todos_recursos, pa_codigos)
            )
        elif codigo_hijo not in todos_recursos:
            partidas.append(
                build_partida(codigo_hijo, cantidad, conceptos, descompuestos, textos,
                              mo_codigos, pa_codigos)
            )

    return Capitulo(
        codigo=codigo,
        descripcion=concepto.get('descripcion', ''),
        descripcion_larga=textos.get(codigo, ''),
        importe_total=float(concepto.get('precio_str', '') or '0'),
        partidas=partidas,
        subcapitulos=subcapitulos,
    )


def parse_bc3(filepath: Path) -> Presupuesto:
    """
    Lee el archivo BC3 y retorna un Presupuesto poblado.

    Orden de clasificación de recursos:
      1. Campo tipo del registro ~C: '1'→MO, '2'→MQ, '3'→MT  (fuente más fiable)
      2. Prefijo del código: MO-→MO, MT-→MT, MQ-→MQ           (convención estándar)
      3. Unidad para elementales sin tipo ni prefijo: 'h'→MO, resto→MT

    Los códigos con descompuesto propio que no son hijos directos de capítulos
    se registran como partidas alzadas (recursos_pa).
    """
    conceptos, descompuestos, textos = read_bc3_records(filepath)

    recursos_mo: dict[str, RecursoMO] = {}
    recursos_mt: dict[str, RecursoElemental] = {}
    recursos_mq: dict[str, RecursoElemental] = {}

    # ── Nivel 1: campo tipo del registro ~C ──────────────────────────────────
    for codigo, datos in conceptos.items():
        tipo_bc3 = datos.get('tipo', '')
        if tipo_bc3 == '1':
            recursos_mo[codigo] = RecursoMO(
                codigo=codigo,
                descripcion=datos['descripcion'],
                precio_hora=float(datos['precio_str'] or '0'),
            )
        elif tipo_bc3 == '2':
            recursos_mq[codigo] = RecursoElemental(
                codigo=codigo,
                descripcion=datos['descripcion'],
                unidad=datos['unidad'],
                precio_unidad=float(datos['precio_str'] or '0'),
            )
        elif tipo_bc3 == '3':
            recursos_mt[codigo] = RecursoElemental(
                codigo=codigo,
                descripcion=datos['descripcion'],
                unidad=datos['unidad'],
                precio_unidad=float(datos['precio_str'] or '0'),
            )

    # ── Nivel 2: prefijo del código para los que no tienen tipo ──────────────
    for codigo, datos in conceptos.items():
        if codigo in recursos_mo or codigo in recursos_mt or codigo in recursos_mq:
            continue
        if codigo.startswith('MO-'):
            recursos_mo[codigo] = RecursoMO(
                codigo=codigo,
                descripcion=datos['descripcion'],
                precio_hora=float(datos['precio_str'] or '0'),
            )
        elif codigo.startswith('MT-'):
            recursos_mt[codigo] = RecursoElemental(
                codigo=codigo,
                descripcion=datos['descripcion'],
                unidad=datos['unidad'],
                precio_unidad=float(datos['precio_str'] or '0'),
            )
        elif codigo.startswith('MQ-'):
            recursos_mq[codigo] = RecursoElemental(
                codigo=codigo,
                descripcion=datos['descripcion'],
                unidad=datos['unidad'],
                precio_unidad=float(datos['precio_str'] or '0'),
            )

    # ── Nivel 3: elementales sin tipo ni prefijo que aparecen en descompuestos
    ya_registrados = set(recursos_mo) | set(recursos_mt) | set(recursos_mq)
    todos_hijos = {cod for items in descompuestos.values() for cod, _ in items}

    # Partidas directas (hijos de capítulos): no son recursos ni PAs
    partidas_directas: set[str] = set()
    codigo_raiz = next((c for c in conceptos if c.endswith('##')), '')

    def recoger_hijos_capitulo(cod: str) -> None:
        for cod_hijo, _ in descompuestos.get(cod, []):
            if cod_hijo.endswith('#'):
                recoger_hijos_capitulo(cod_hijo)
            else:
                partidas_directas.add(cod_hijo)

    for cod_cap in [c for c, _ in descompuestos.get(codigo_raiz, []) if c.endswith('#')]:
        recoger_hijos_capitulo(cod_cap)

    recursos_pa: dict[str, RecursoElemental] = {}

    for codigo in todos_hijos - ya_registrados:
        if (codigo.endswith('#')
                or codigo not in conceptos
                or codigo.startswith('%')
                or codigo in partidas_directas):
            continue
        datos = conceptos[codigo]
        unidad = datos['unidad']
        tipo_bc3 = datos.get('tipo', '')

        if codigo in descompuestos:
            # Tiene descompuesto propio → partida alzada / unidad auxiliar
            recursos_pa[codigo] = RecursoElemental(
                codigo=codigo,
                descripcion=datos['descripcion'],
                unidad=unidad,
                precio_unidad=float(datos['precio_str'] or '0'),
            )
        elif tipo_bc3 == '1' or unidad == 'h':
            recursos_mo[codigo] = RecursoMO(
                codigo=codigo,
                descripcion=datos['descripcion'],
                precio_hora=float(datos['precio_str'] or '0'),
            )
        elif tipo_bc3 == '2':
            recursos_mq[codigo] = RecursoElemental(
                codigo=codigo,
                descripcion=datos['descripcion'],
                unidad=unidad,
                precio_unidad=float(datos['precio_str'] or '0'),
            )
        elif unidad and unidad != '%':
            recursos_mt[codigo] = RecursoElemental(
                codigo=codigo,
                descripcion=datos['descripcion'],
                unidad=unidad,
                precio_unidad=float(datos['precio_str'] or '0'),
            )

    # Auxiliares mal clasificados: recursos en MT/MQ que tienen descompuesto propio
    # y no son partidas directas → reclasificar como PA (unidades auxiliares)
    for banco in (recursos_mt, recursos_mq):
        for codigo in list(banco.keys()):
            if codigo in descompuestos and codigo not in partidas_directas:
                datos = conceptos[codigo]
                recursos_pa[codigo] = RecursoElemental(
                    codigo=codigo,
                    descripcion=datos['descripcion'],
                    unidad=datos['unidad'],
                    precio_unidad=float(datos['precio_str'] or '0'),
                )
                del banco[codigo]

    todos_recursos = set(recursos_mo) | set(recursos_mt) | set(recursos_mq)
    mo_codigos   = set(recursos_mo)
    pa_codigos   = set(recursos_pa)

    # ── Estructura de capítulos ───────────────────────────────────────────────
    concepto_raiz = conceptos.get(codigo_raiz, {})
    codigos_capitulos = [
        c for c, _ in descompuestos.get(codigo_raiz, [])
        if c.endswith('#')
    ]

    # ── Metadatos del proyecto extraídos del BC3 ──────────────────────────────
    config = ProjectConfig(proyecto=concepto_raiz.get('descripcion', ''))

    return Presupuesto(
        codigo=codigo_raiz,
        descripcion=concepto_raiz.get('descripcion', ''),
        importe_total=float(concepto_raiz.get('precio_str', '') or '0'),
        capitulos=[
            build_capitulo(c, conceptos, descompuestos, textos,
                           mo_codigos, todos_recursos, pa_codigos)
            for c in codigos_capitulos
        ],
        recursos_mo=recursos_mo,
        recursos_mt=recursos_mt,
        recursos_mq=recursos_mq,
        recursos_pa=recursos_pa,
        descompuestos_raw=descompuestos,
        config=config,
    )


if __name__ == '__main__':
    import sys
    from pathlib import Path

    if len(sys.argv) < 2:
        print('Uso: python -m gantt.bc3.parser <nombre-proyecto> [archivo.bc3]')
        sys.exit(1)

    project_name = sys.argv[1]
    bc3_filename  = sys.argv[2] if len(sys.argv) > 2 else None
    input_dir    = Path('projects') / project_name / 'input'

    if bc3_filename:
        bc3_path = input_dir / bc3_filename
        if not bc3_path.exists():
            raise FileNotFoundError(f'No se encontró: {bc3_path}')
    else:
        bc3_files = list(input_dir.glob('*.bc3'))
        if not bc3_files:
            raise FileNotFoundError(f'No se encontró ningún .bc3 en {input_dir}')
        if len(bc3_files) > 1:
            raise ValueError(
                f'Multiples .bc3 en {input_dir}. '
                f'Especifica el archivo como segundo argumento.\n'
                f'Disponibles: {[f.name for f in bc3_files]}'
            )
        bc3_path = bc3_files[0]

    presupuesto = parse_bc3(bc3_path)
    print(f'Capítulos principales: {len(presupuesto.capitulos)}')
    print(f'Recursos MO: {len(presupuesto.recursos_mo)}')
    print(f'Recursos MT: {len(presupuesto.recursos_mt)}')
    print(f'Recursos MQ: {len(presupuesto.recursos_mq)}')
    print(f'Partidas alzadas: {len(presupuesto.recursos_pa)}')

    def print_capitulo(cap: Capitulo, indent: int = 0) -> None:
        prefix = '  ' * indent
        print(f'{prefix}{cap.codigo:<10} {cap.descripcion[:45]:<45} '
              f'{len(cap.partidas)} partidas / {len(cap.subcapitulos)} subcapítulos')
        for sub in cap.subcapitulos:
            print_capitulo(sub, indent + 1)

    for cap in presupuesto.capitulos:
        print_capitulo(cap)
