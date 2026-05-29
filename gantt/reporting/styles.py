"""
Constantes de estilo visual para los reportes Excel del proyecto.

Paleta monocromática azul con accents funcionales semánticos.
Modificar aquí para cambiar el estilo global de todos los reportes.
"""

from openpyxl.styles import Alignment, Font, PatternFill

# Paleta de colores — azul corporativo con accents funcionales
COLOR_AZUL_OSCURO = '1B3A5C'   # cabeceras principales
COLOR_AZUL_MEDIO  = '2E6DA4'   # cabeceras secundarias
COLOR_GRIS_CLARO  = 'F5F7FA'   # filas alternas
COLOR_BLANCO      = 'FFFFFF'   # filas normales
COLOR_VERDE       = '2D7D46'   # cumple / óptimo / positivo
COLOR_ROJO        = 'C0392B'   # no cumple / crítico / alerta
COLOR_AMARILLO    = 'D4A017'   # advertencia / estacional
COLOR_GRIS_NEUTRO = '9E9E9E'   # fuera de plazo / inactivo

# Fills
FILL_HEADER  = PatternFill('solid', fgColor='1F3864')
FILL_SUBHEAD = PatternFill('solid', fgColor=COLOR_AZUL_MEDIO)
FILL_CAP     = PatternFill('solid', fgColor='D9D9D9')
FILL_ALT     = PatternFill('solid', fgColor='F2F2F2')
FILL_WHITE   = PatternFill('solid', fgColor=COLOR_BLANCO)
FILL_GREEN   = PatternFill('solid', fgColor='C6EFCE')
FILL_RED     = PatternFill('solid', fgColor='FFC7CE')
FILL_YELLOW  = PatternFill('solid', fgColor='FFFFC0')

# Fonts
FONT_HEADER = Font(bold=True, color='FFFFFF', name='Calibri', size=10)
FONT_BOLD   = Font(bold=True, name='Calibri', size=10)
FONT_ITALIC = Font(italic=True, name='Calibri', size=10)
FONT_NORMAL = Font(name='Calibri', size=10)

# Alineaciones
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_LEFT   = Alignment(horizontal='left',   vertical='center')

# Formatos numéricos
FMT_EUROS  = '#,##0.00'
FMT_HORAS  = '#,##0.0'
FMT_ENTERO = '#,##0'
FMT_PCT    = '0.0'
