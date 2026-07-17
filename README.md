# Gantt Tools — Generador de documentos de presupuesto y planificación

Sistema de procesado de archivos BC3 (FIEBDC-3) para la generación automática de
documentos de presupuesto en formato Excel, gráficos de análisis económico y
diagramas de planificación temporal (Gantt).

---

## Requisitos

- **Python 3.12 o superior**
- **Fuente UIBSans** instalada en el sistema (corporativa UIB).
  Si no está disponible, los documentos usarán la fuente por defecto del sistema;
  el layout puede variar ligeramente pero los datos son correctos.

### Instalación de dependencias

Con el entorno virtual activado:

```bash
pip install -r requirements.txt
```

### Crear y activar el entorno virtual (primera vez)

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# Mac / Linux
python -m venv .venv
source .venv/bin/activate
```

---

## Estructura de carpetas

```
gantt-tools/
├── projects/                    # Un subdirectorio por proyecto
│   └── NombreProyecto/
│       ├── input/
│       │   ├── archivo.bc3        ← presupuesto (obligatorio)
│       │   ├── config.yaml        ← configuración del proyecto (recomendado)
│       │   └── planificacion.yaml ← cronograma (opcional)
│       └── output/                ← documentos generados aquí
├── gantt/
│   ├── bc3/                     # Lectura y parseo del BC3
│   ├── planning/                # Cálculo de duraciones y análisis CPM
│   └── reporting/               # Generación de Excel, gráficos y diagramas
│       ├── palette.py           # Colores y fuentes (modificar para otro cliente)
│       ├── styles.py            # Objetos openpyxl y paletas matplotlib
│       ├── presupuesto_exporter.py
│       ├── excel_exporter.py
│       ├── analysis_charts.py
│       └── network_diagram.py
├── main.py                      # Punto de entrada del pipeline
├── requirements.txt
└── README.md
```

---

## Uso

Desde la raíz del repositorio, con el entorno virtual activado:

```bash
python main.py <nombre-proyecto>
python main.py <nombre-proyecto> <archivo.bc3>
```

- Si hay un solo `.bc3` en `input/`, no hace falta especificarlo.
- Si hay más de uno, es obligatorio indicar cuál usar.

### Ejemplos

```bash
python main.py Complexe-Balear
python main.py Complexe-Balear pressupost_v02.bc3
python main.py 4t_2t
```

---

## Documentos generados

### Siempre (solo con BC3 y config.yaml)

| Archivo | Contenido |
|---|---|
| `PRES.01_Cuadro_Oferta.xlsx` | Cuadro de oferta con columnas editables para el licitador |
| `PRES.02.01_Cuadro_Precios_1.xlsx` | Cuadro de precios n.º 1 — precios unitarios por partida |
| `PRES.02.02_Cuadro_Precios_2.xlsx` | Cuadro de precios n.º 2 — descomposición por recurso |
| `PRES.02.03_Presupuesto_Descompuesto.xlsx` | Presupuesto descompuesto y mediciones con precios |
| `PRES.02.04_Resumen_Capitulos.xlsx` | Resumen por capítulos con cascada financiera PEM → PGL |
| `PRES.03_Mediciones_Ciegas.xlsx` | Mediciones sin precios ni horas MO (para solicitar ofertas) |
| `PRES.05_VEC_Liquidacion.xlsx` | VEC, liquidación máxima y cascada de IVA |
| `JUST_PRECIOS_Recursos.xlsx` | Justificación de precios — totales de recursos por proyecto |
| `*_gantt.xlsx` | Excel de análisis económico (capítulos, MO, partidas detalladas) |
| `reporting/*.png` | Gráficos: importes por capítulo, horas MO, distribución por perfil |

### Solo si existe `planificacion.yaml`

| Archivo | Contenido |
|---|---|
| `*_gantt.xlsx` | Excel actualizado con datos de planificación temporal |
| `*_gantt.png` | Diagrama de red / Gantt |

---

## Configuración del proyecto — `config.yaml`

Crear en `projects/<nombre>/input/config.yaml`.
Si no existe, el sistema funciona con valores por defecto (porcentajes estándar
españoles, encabezados vacíos).

```yaml
# Metadatos documentales
entidad:            UNIVERSITAT DE LES ILLES BALEARS
proyecto:           Reforma del sistema de climatització
edificio:           COMPLEXE BALEAR DE RECERCA
numero_expediente:  2025/EXP-001   # aparece en el encabezado de todos los documentos
footer_org:         Complexe Balear de Recerca
footer_exp:         Expedient de licitació

# Parámetros financieros del contrato
porcentaje_gg:          0.13   # Gastos Generales (13 %)
porcentaje_bi:          0.06   # Beneficio Industrial (6 %)
iva_obra:               0.21   # IVA obra (21 %)
iva_gr:                 0.10   # IVA gestión de residuos (10 %)
porcentaje_liquidacion: 0.10   # Liquidación máxima (10 %)

# Estructura del presupuesto
# Código BC3 del capítulo de gestión de residuos para IVA diferenciado.
# Dejar vacío si el proyecto no tiene ese capítulo.
codigo_capitulo_gr: '13#'
```

Solo incluir los campos que necesites cambiar. Los omitidos usan valores por defecto.

### Referencia de campos

| Campo | Defecto | Descripción |
|---|---|---|
| `entidad` | *(vacío)* | Entidad contratante |
| `proyecto` | descripción del BC3 | Nombre del proyecto |
| `edificio` | *(vacío)* | Localización o nombre del inmueble |
| `numero_expediente` | *(vacío)* | Nº de expediente — aparece en el encabezado |
| `footer_org` | *(vacío)* | Pie de página izquierda en todos los documentos |
| `footer_exp` | *(vacío)* | Pie de página central en todos los documentos |
| `porcentaje_gg` | `0.13` | Gastos Generales |
| `porcentaje_bi` | `0.06` | Beneficio Industrial |
| `iva_obra` | `0.21` | IVA aplicable a la obra |
| `iva_gr` | `0.10` | IVA aplicable al capítulo de residuos |
| `porcentaje_liquidacion` | `0.10` | Base para cálculo de liquidación máxima |
| `codigo_capitulo_gr` | *(vacío)* | Código BC3 del capítulo de gestión de residuos |

---

## Planificación temporal — `planificacion.yaml` (opcional)

Permite generar el diagrama de Gantt y el análisis de ruta crítica (CPM).
Si no existe este archivo, el sistema genera igualmente todos los documentos
de presupuesto.

```yaml
proyecto:
  nombre:                  NombreProyecto
  fecha_inicio:            '2025-06-01'   # formato AAAA-MM-DD, entre comillas
  horas_dia:               8
  dias_semana:             5              # días laborables por semana (L-V = 5)
  plazo_contractual_dias:  120

# Cada escenario mapea código de recurso MO -> nº de operarios asignados
escenarios:
  base:
    MO-mo001: 1
    MO-mo102: 1
  optimizado:
    MO-mo001: 2
    MO-mo102: 1

# Diccionario de tareas indexado por ID (no una lista)
tareas:
  T01:
    nombre: Instalación cuadro eléctrico
    tipo: tarea
    capitulos_bc3: ['01.01']
    dependencias: []
  T02:
    nombre: Tendido de cable
    tipo: tarea
    capitulos_bc3: ['01.02']
    dependencias: [T01]
  FIN:
    nombre: Entrega
    tipo: hito
    capitulos_bc3: []
    duracion_dias_fija: 0
    dependencias: [T02]
    es_fin_plazo: true
```

**Notas:**
- `tareas` es un diccionario indexado por ID de tarea, no una lista.
- `tipo` es obligatorio en cada tarea: `tarea` o `hito`.
- Los códigos en `capitulos_bc3` deben coincidir exactamente con los del BC3
  (código de partida, o código de capítulo terminado en `#` para sumar todas
  sus partidas).
- Los códigos en cada escenario deben coincidir con los recursos MO del BC3,
  y su valor es el número de operarios asignados a ese recurso.
- La tarea con `es_fin_plazo: true` define el hito de fin de contrato.
- Puede haber varios escenarios para comparar diferentes intensidades de recurso.

---

## Añadir un proyecto nuevo

1. Crear la carpeta `projects/<NombreProyecto>/input/`
2. Copiar el archivo `.bc3` dentro
3. Copiar `config.yaml` de un proyecto existente y ajustar los valores
4. Ejecutar: `python main.py <NombreProyecto>`

Los resultados aparecen en `projects/<NombreProyecto>/output/`.

---

## Adaptar la identidad visual (otro cliente)

Los colores, fuentes y paleta de gráficos están centralizados en:

```
gantt/reporting/palette.py
```

Cambiar `FONT_PRES` para usar otra fuente corporativa.
Cambiar `GRIS_OSCURO` / `AZUL_OSCURO` para adaptar la paleta a otro cliente.
El cambio se propaga automáticamente a todos los documentos generados.

---

## Solución de problemas frecuentes

**"No se encontró ningún .bc3 en projects/..."**
El nombre del proyecto en el comando no coincide con el nombre de la carpeta en `projects/`.

**"Múltiples .bc3 en ..."**
Especificar el archivo: `python main.py NombreProyecto archivo.bc3`

**Los documentos usan Arial en lugar de UIBSans**
Instalar la fuente UIBSans en el sistema operativo y reiniciar el terminal.

**Los archivos Excel están bloqueados (PermissionError)**
Cerrar los archivos abiertos en Excel y volver a ejecutar.

**Los gráficos de horas MO están vacíos o a cero**
Verificar que el BC3 tiene recursos MO con horas definidas en los descompuestos.
El sistema clasifica MO por el campo `tipo=1` del BC3, o por prefijo `MO-`, o por
unidad `h` (en ese orden de prioridad).

**El campo `numero_expediente` del yaml no aparece en los documentos**
El campo aparece en el encabezado justo antes de la línea "ANEXO ECONÓMICO".
Si no se ve, verificar que no está vacío en `config.yaml`.

---

## Normas de programación

Convenciones definidas en `Formato-de-programación.txt`:

- Código en inglés; comentarios y docstrings en castellano.
- Modelos de datos con Pydantic.
- Funciones todas públicas
- Paleta visual centralizada en `palette.py`; ningún módulo define colores propios.
- Sin `try/except` profiláctico.
- Sin abstracciones prematuras.