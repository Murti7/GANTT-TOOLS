"""
Paleta visual del proyecto. Único archivo que hay que modificar para adaptar
el sistema a otro cliente o identidad corporativa.

Los módulos de generación (presupuesto_exporter, excel_exporter, analysis_charts)
leen de aquí; nunca definen colores o fuentes propios.
"""

# ── Tipografía ────────────────────────────────────────────────────────────────
FONT_PRES  = 'UIBSans'      # documentos de presupuesto; fallback manual: 'Arial'
FONT_EXCEL = 'Calibri'      # hojas de análisis y Gantt
FONT_CHART = 'DejaVu Sans'  # matplotlib (acepta solo fuentes del sistema disponibles)

# ── Grises — identidad presupuesto ───────────────────────────────────────────
NEGRO        = '000000'
BLANCO       = 'FFFFFF'
GRIS_OSCURO  = '404040'   # cabeceras principales presupuesto
GRIS_MEDIO   = 'D9D9D9'   # cabeceras de tabla / capítulos
GRIS_SUAVE   = 'F2F2F2'   # subcapítulos / filas alternas
GRIS_LINEA   = 'C8C8C8'   # bordes de celda
GRIS_TEXTO   = '404040'   # texto detalle

# ── Azules — identidad análisis / Gantt ──────────────────────────────────────
AZUL_OSCURO  = '1F3864'   # cabeceras principales análisis
AZUL_MEDIO   = '2E6DA4'   # subcabeceras análisis
AZUL_NEUTRO  = '9E9E9E'   # inactivo / fuera de plazo

# ── Semánticos funcionales ────────────────────────────────────────────────────
VERDE_EXITO       = '2D7D46'   # texto: cumple / positivo
VERDE_EXITO_FONDO = 'C6EFCE'   # fondo: cumple
ROJO_ALERTA       = 'C0392B'   # texto: no cumple / crítico
ROJO_ALERTA_FONDO = 'FFC7CE'   # fondo: alerta
AMARILLO_ADV      = 'D4A017'   # advertencia / estacional
AMARILLO_EDIT_PRES = 'FFF9E6'  # celdas editables presupuesto
AMARILLO_EDIT_GANTT = 'FFFFC0' # celdas editables Gantt

# ── Paleta matplotlib ─────────────────────────────────────────────────────────
CHART_AZUL        = '#1F3864'
CHART_AZUL_CLARO  = '#5B9BD5'   # variante clara para pares oscuro/claro
CHART_AZUL_MEDIO  = '#4472C4'   # azul medio (gráficos MO)
CHART_VERDE       = '#70AD47'   # verde accesible para gráficos
CHART_VERDE_CLARO = '#A9D18E'
CHART_NARANJA     = '#C55A11'
CHART_NARANJA_CL  = '#F4B183'
CHART_TIERRA      = '#7F6000'
CHART_ROSAOSCURO  = '#A64D79'
CHART_MORADO      = '#8E44AD'
CHART_CYAN        = '#16A085'
CHART_GRIS        = '#95A5A6'
CHART_GRIS_CLARO  = '#BFBFBF'   # "Otros" en gráficos apilados
CHART_BORDE       = '#DDDDDD'   # borde leyenda / elementos decorativos
CHART_GRID        = '#D0D0D0'
CHART_TEXTO       = '#222222'
CHART_MUTED       = '#666666'
