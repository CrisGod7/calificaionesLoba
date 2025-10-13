# app/styles_config.py
"""
Configuración de estilos para el Sistema de Calificaciones Lobatchewsky
"""

# Paleta de colores corporativa
COLORS = {
    # Colores principales
    'principal': '#005DAB',  # Azul corporativo Lobatchewsky
    'principal_hover': '#004080',  # Azul oscuro para hover
    'principal_claro': '#337AB7',  # Azul claro para detalles

    # Colores de fondo
    'blanco': '#FFFFFF',
    'gris_claro': '#F2F2F2',
    'gris_medio': '#CCCCCC',
    'gris_oscuro': '#666666',

    # Colores de texto
    'texto_principal': '#333333',
    'texto_secundario': '#666666',
    'texto_claro': '#999999',

    # Colores de estado
    'exito': '#28A745',
    'exito_hover': '#218838',
    'advertencia': '#FFC107',
    'advertencia_hover': '#E0A800',
    'error': '#DC3545',
    'error_hover': '#C82333',
    'info': '#17A2B8',
    'info_hover': '#138496',

    # Colores de rendimiento
    'excelente': '#C6EFCE',
    'muy_bien': '#FFEB9C',
    'bien': '#FFC7CE',
    'regular': '#FFD9D9',
    'bajo': '#FFC7CE',
}

# Configuración de fuentes
FONTS = {
    'titulo_principal': ('Poppins', 18, 'bold'),
    'titulo_seccion': ('Poppins', 14, 'bold'),
    'subtitulo': ('Poppins', 12, 'bold'),
    'subtitulo_seccion': ('Poppins', 11, 'bold'),
    'normal': ('Open Sans', 10),
    'normal_bold': ('Open Sans', 10, 'bold'),
    'pequeno': ('Open Sans', 9),
    'pequeno_bold': ('Open Sans', 9, 'bold'),
    'codigo': ('Courier New', 10),
    'codigo_pequeno': ('Courier New', 9),
}

# Tamaños de ventana
WINDOW_SIZES = {
    'main': '1400x800',
    'generator': '750x700',
    'preview': '700x550',
}

# Espaciado y padding
SPACING = {
    'padding_grande': 20,
    'padding_medio': 15,
    'padding_pequeno': 10,
    'padding_minimo': 5,
    'margen_seccion': 20,
    'margen_elemento': 10,
}

# Dimensiones de componentes
DIMENSIONS = {
    'panel_lateral_width': 350,
    'header_height': 80,
    'boton_height': 35,
    'boton_pequeno_height': 28,
    'border_radius': 8,
    'separator_height': 2,
}

# Iconos emoji para diferentes secciones
ICONS = {
    'logo': '🎓',
    'archivos': '📁',
    'clave': '🔑',
    'alumnos': '👥',
    'procesar': '⚙️',
    'calificar': '🚀',
    'reportes': '📊',
    'excel': '📄',
    'csv': '📋',
    'ver': '👁️',
    'crear': '➕',
    'editar': '✏️',
    'eliminar': '🗑️',
    'guardar': '💾',
    'buscar': '🔍',
    'configuracion': '⚙️',
    'ayuda': '❓',
    'exito': '✓',
    'error': '✗',
    'advertencia': '⚠',
    'info': 'ℹ',
    'cargando': '⏳',
}

# Mensajes del sistema
MESSAGES = {
    'bienvenida': '👈 Comience cargando los archivos',
    'procesando': '⏳ Procesando calificaciones...',
    'exito': '✅ Procesamiento completado',
    'error_archivos': '✗ Error al cargar archivos',
    'error_formato': '✗ Formato de archivo incorrecto',
    'archivo_cargado': '✓ Archivo cargado correctamente',
    'sin_resultados': 'No hay resultados para mostrar',
    'excel_generado': '✅ Reporte Excel generado correctamente',
}

# Configuración de tablas
TABLE_CONFIG = {
    'header_bg': COLORS['principal'],
    'header_fg': COLORS['blanco'],
    'row_height': 25,
    'header_height': 35,
    'alternating_colors': [COLORS['blanco'], COLORS['gris_claro']],
}

# Configuración de gráficos
CHART_CONFIG = {
    'bar_color': COLORS['principal'],
    'grid_color': COLORS['gris_medio'],
    'text_color': COLORS['texto_principal'],
    'background_color': COLORS['blanco'],
}


def obtener_color_rendimiento(porcentaje: float) -> str:
    """Retorna el color según el porcentaje de rendimiento."""
    if porcentaje >= 90:
        return COLORS['excelente']
    elif porcentaje >= 80:
        return COLORS['muy_bien']
    elif porcentaje >= 70:
        return COLORS['bien']
    elif porcentaje >= 60:
        return COLORS['regular']
    else:
        return COLORS['bajo']


def obtener_texto_rendimiento(porcentaje: float) -> str:
    """Retorna el texto descriptivo según el porcentaje."""
    if porcentaje >= 90:
        return "Excelente"
    elif porcentaje >= 80:
        return "Muy bien"
    elif porcentaje >= 70:
        return "Bien"
    elif porcentaje >= 60:
        return "Regular"
    else:
        return "Necesita mejorar"


def oscurecer_color(color: str, factor: float = 0.8) -> str:
    """
    Oscurece un color hexadecimal.

    Args:
        color: Color en formato hexadecimal (#RRGGBB)
        factor: Factor de oscurecimiento (0.0 a 1.0)

    Returns:
        Color oscurecido en formato hexadecimal
    """
    color = color.lstrip('#')
    rgb = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
    rgb = tuple(max(0, int(c * factor)) for c in rgb)
    return f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'


def aclarar_color(color: str, factor: float = 1.2) -> str:
    """
    Aclara un color hexadecimal.

    Args:
        color: Color en formato hexadecimal (#RRGGBB)
        factor: Factor de aclarado (1.0 a 2.0)

    Returns:
        Color aclarado en formato hexadecimal
    """
    color = color.lstrip('#')
    rgb = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
    rgb = tuple(min(255, int(c * factor)) for c in rgb)
    return f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'


# Configuración de animaciones (para futuras mejoras)
ANIMATION_CONFIG = {
    'duration': 200,  # milisegundos
    'easing': 'ease-in-out',
    'fade_steps': 10,
}

if __name__ == "__main__":
    print("Configuración de estilos para Sistema Lobatchewsky")
    print(f"\nColores principales:")
    print(f"  Azul corporativo: {COLORS['principal']}")
    print(f"  Éxito: {COLORS['exito']}")
    print(f"  Error: {COLORS['error']}")

    print(f"\nFuentes disponibles:")
    for nombre, fuente in FONTS.items():
        print(f"  {nombre}: {fuente}")