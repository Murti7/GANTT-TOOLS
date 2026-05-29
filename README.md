# gantt-tools

## Descripción del proyecto

`gantt-tools` es una herramienta de línea de comandos que genera automáticamente
un diagrama de Gantt y un archivo Excel de planificación a partir de un presupuesto
de construcción en formato BC3 (FIEBDC-3) y una configuración de cuadrillas de trabajo.

El sistema lee el presupuesto, calcula las duraciones de cada capítulo y partida
en función de las horas de mano de obra y la composición de las cuadrillas disponibles,
y exporta los resultados a un Excel con cuatro hojas: resumen por capítulo,
partidas detalladas, escenarios de duración y tabla de importación para MS Project.

---

## Estructura de carpetas

```
gantt-tools/
├── projects/                    # Un subdirectorio por proyecto
│   └── Complexe-Balear/
│       ├── input/               # Archivos de entrada: .bc3 y planificacion.yaml
│       └── output/              # Excel y gráficos generados (ignorado por git)
├── gantt/                       # Paquete principal con toda la lógica
│   ├── bc3/                     # Lectura y parseo del archivo BC3
│   │   ├── models.py            # Modelos Pydantic: Partida, Capitulo, Presupuesto
│   │   └── parser.py            # parse_bc3(), build_capitulo(), read_bc3_records()
│   ├── planning/                # Cálculo de duraciones y análisis del cronograma
│   │   ├── models.py            # TareaGantt, PlanificacionProyecto, EscenarioRecursos
│   │   ├── calculator.py        # calcular_planificacion(), forward pass CPM
│   │   └── analyser.py          # calcular_holguras(), calcular_sensibilidad(), etc.
│   └── reporting/               # Generación de outputs visuales
│       ├── styles.py            # Constantes de estilo visual (paleta, fonts, formatos)
│       ├── excel_exporter.py    # exportar_analisis() — Excel multi-hoja con Dashboard
│       ├── network_diagram.py   # generar_diagrama_red() — PNG embebido en Excel
│       └── analysis_charts.py   # Gráficos de reporting por capítulo y perfil MO
├── main.py                      # Punto de entrada: orquesta el pipeline completo
├── requirements.txt             # Dependencias del proyecto
└── README.md
```

---

## Cómo añadir un proyecto nuevo

1. Crear la carpeta del proyecto dentro de `projects/`:
   ```
   projects/
   └── nombre-del-proyecto/
       ├── input/
       └── output/
   ```

2. Copiar el archivo BC3 en `input/` con el nombre `presupuesto.bc3`.

3. Crear el archivo `input/cuadrillas.yaml` con la configuración de cuadrillas
   para ese proyecto (recursos disponibles y su composición).

4. Ejecutar el pipeline (ver sección siguiente).

---

## Cómo ejecutar

Con el entorno virtual activado, desde la raíz del proyecto:

```bash
python main.py <nombre-proyecto>
python main.py <nombre-proyecto> <archivo.bc3>
```

Ejemplos:

```bash
python main.py Complexe-Balear
python main.py Complexe-Balear pressupost_original.bc3
python main.py Complexe-Balear pressupost_revisat.bc3
```

Si hay un solo `.bc3` en `input/`, no hace falta especificarlo.
Si hay más de uno, es obligatorio especificar cuál usar.

El Excel de resultados se escribirá en `projects/Complexe-Balear/output/`.

---

## Entorno virtual: creación y activación

Crear el entorno virtual (solo la primera vez):

```bash
python -m venv .venv
```

Activar el entorno:

- **Windows:**
  ```bash
  .venv\Scripts\activate
  ```
- **Mac / Linux:**
  ```bash
  source .venv/bin/activate
  ```

Para desactivarlo:

```bash
deactivate
```

---

## Dependencias e instalación

Con el entorno virtual activado:

```bash
pip install pydantic openpyxl pyyaml pytest
```

Para regenerar el archivo `requirements.txt` tras instalar nuevas dependencias:

```bash
pip freeze > requirements.txt
```

Para instalar desde `requirements.txt` en otro equipo:

```bash
pip install -r requirements.txt
```

---

## Normas de programación

Este proyecto sigue las convenciones definidas en `Formato-de-programación.txt`.

Resumen de las normas principales:

- **Idioma del código:** inglés (nombres de variables, funciones, clases).
- **Idioma de comentarios y docstrings:** castellano.
- **Modelos de datos:** siempre con Pydantic, sin dataclasses ni dicts desnudos.
- **Funciones públicas:** sin prefijo `_`. Solo se usan prefijos privados cuando la función
  es un detalle de implementación sin valor fuera de su módulo.
- **Constantes de estilo centralizadas:** toda la paleta, fonts y formatos de Excel
  están en `gantt/reporting/styles.py`. No hardcodear colores en otros módulos.
- **Imports al inicio del módulo:** nunca dentro de funciones, salvo en bloques `__main__`.
- **Sin referencias hardcodeadas a IDs de tareas:** usar campos semánticos del modelo
  (`es_fin_plazo`, `tipo`, etc.) en lugar de comparar `t.id == 'H11'`.
- **Sin abstracciones prematuras:** no crear clases artificiales para envolver funciones simples.
- **Responsabilidad única por módulo:** cada archivo tiene un propósito claro y acotado.
- **Sin try/except profiláctico:** solo capturar excepciones cuando se sabe qué hacer con ellas.
- **Constructores simples:** sin lógica pesada en `__init__`.
