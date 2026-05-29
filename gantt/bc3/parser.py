"""
Módulo responsable de leer y parsear un archivo BC3 (FIEBDC-3).

Responsabilidad única: transformar el contenido de texto de un archivo .bc3
en un objeto Presupuesto con la estructura de capítulos y partidas correspondiente.
"""

from pathlib import Path

from gantt.bc3.models import Capitulo, LineaDescompuesto, Partida, Presupuesto, RecursoElemental, RecursoMO

RESOURCE_PREFIXES = ('MO-', 'MT-', 'MQ-')


def read_bc3_records(
    filepath: Path,
) -> tuple[dict[str, dict[str, str]], dict[str, list[tuple[str, float]]], dict[str, str]]:
    """
    Lee el archivo BC3 en una sola pasada y extrae los registros ~C, ~D y ~T.
    Retorna (conceptos, descompuestos, textos) indexados por código.
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
                        'unidad': parts[2].strip(),
                        'descripcion': parts[3].strip(),
                        'precio_str': parts[4].strip(),
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
                # Formato ~T|codigo| o ~T|codigo|texto_primera_linea
                parts_t = line[3:].split('|', 1)
                current_texto = parts_t[0].strip()
                current_texto_lines = []
                if len(parts_t) > 1:
                    first = parts_t[1].strip()
                    if first:
                        current_texto_lines.append(first)

            elif current_parent is not None:
                if line.startswith('|') or line.startswith('\\'):
                    # Eliminar delimitadores iniciales (| en el primer ítem, \ o \\ en los siguientes)
                    content = line.lstrip('|\\')
                    # Terminador: \| o \\| → tras lstrip queda solo |
                    if not content or content == '|':
                        descompuestos[current_parent] = current_items
                        current_parent = None
                        current_items = []
                    else:
                        # Separador de campo en el archivo: \ o \\ según versión del BC3.
                        # Filtrando partes vacías se soportan ambas convenciones.
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
                    # Registro no continuación: cierra el descompuesto en curso
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


def build_partida(
    codigo: str,
    cantidad: float,
    conceptos: dict[str, dict[str, str]],
    descompuestos: dict[str, list[tuple[str, float]]],
    textos: dict[str, str],
) -> Partida:
    """
    Construye una Partida desde el catálogo de conceptos, su descompuesto y textos ~T.
    La cantidad proviene del registro ~D del capítulo padre.
    """
    concepto = conceptos.get(codigo, {})
    lineas_mo = [
        LineaDescompuesto(codigo_recurso=r_code, cantidad=r_cant)
        for r_code, r_cant in descompuestos.get(codigo, [])
        if r_code.startswith('MO-')
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
) -> Capitulo:
    """
    Construye recursivamente un Capitulo con sus subcapítulos y partidas.
    Ítems con código terminado en '#' son subcapítulos; el resto son partidas.
    """
    concepto = conceptos.get(codigo, {})
    subcapitulos = []
    partidas = []

    for codigo_hijo, cantidad in descompuestos.get(codigo, []):
        if codigo_hijo.endswith('#'):
            subcapitulos.append(build_capitulo(codigo_hijo, conceptos, descompuestos, textos))
        elif not codigo_hijo.startswith(RESOURCE_PREFIXES):
            partidas.append(build_partida(codigo_hijo, cantidad, conceptos, descompuestos, textos))

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
    Lee el archivo BC3 indicado y retorna un objeto Presupuesto poblado
    con la jerarquía completa de capítulos, partidas y recursos MO.
    """
    conceptos, descompuestos, textos = read_bc3_records(filepath)

    # Recursos MO: conceptos cuyo código empieza por 'MO-'
    recursos_mo = {
        codigo: RecursoMO(
            codigo=codigo,
            descripcion=datos['descripcion'],
            precio_hora=float(datos['precio_str'] or '0'),
        )
        for codigo, datos in conceptos.items()
        if codigo.startswith('MO-')
    }

    # Recursos MT: materiales (MT- y porcentajes %MT)
    recursos_mt = {
        codigo: RecursoElemental(
            codigo=codigo,
            descripcion=datos['descripcion'],
            unidad=datos['unidad'],
            precio_unidad=float(datos['precio_str'] or '0'),
        )
        for codigo, datos in conceptos.items()
        if codigo.startswith('MT-') or codigo.startswith('%MT')
    }

    # Recursos MQ: maquinaria
    recursos_mq = {
        codigo: RecursoElemental(
            codigo=codigo,
            descripcion=datos['descripcion'],
            unidad=datos['unidad'],
            precio_unidad=float(datos['precio_str'] or '0'),
        )
        for codigo, datos in conceptos.items()
        if codigo.startswith('MQ-')
    }

    # Código raíz del presupuesto: termina en '##'
    codigo_raiz = next((c for c in conceptos if c.endswith('##')), '')
    concepto_raiz = conceptos.get(codigo_raiz, {})

    # Capítulos de primer nivel: hijos directos del código raíz que terminan en '#'
    codigos_capitulos = [
        c for c, _ in descompuestos.get(codigo_raiz, [])
        if c.endswith('#')
    ]

    return Presupuesto(
        codigo=codigo_raiz,
        descripcion=concepto_raiz.get('descripcion', ''),
        importe_total=float(concepto_raiz.get('precio_str', '') or '0'),
        capitulos=[build_capitulo(c, conceptos, descompuestos, textos) for c in codigos_capitulos],
        recursos_mo=recursos_mo,
        recursos_mt=recursos_mt,
        recursos_mq=recursos_mq,
        descompuestos_raw=descompuestos,
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

    def print_capitulo(cap: Capitulo, indent: int = 0) -> None:
        prefix = '  ' * indent
        print(f'{prefix}{cap.codigo:<10} {cap.descripcion[:45]:<45} '
              f'{len(cap.partidas)} partidas / {len(cap.subcapitulos)} subcapítulos')
        for sub in cap.subcapitulos:
            print_capitulo(sub, indent + 1)

    for cap in presupuesto.capitulos:
        print_capitulo(cap)
