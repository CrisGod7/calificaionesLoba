# app/config.py

import os

# Rutas y configuración de archivos
RUTA_DATOS = os.path.join(os.path.dirname(__file__), '..', 'data')
RUTA_EXPORTACION_DEFAULT = os.path.join(RUTA_DATOS, 'resultados')

# Crear carpeta de datos si no existe
if not os.path.exists(RUTA_DATOS):
    os.makedirs(RUTA_DATOS)
    print(f"Carpeta '{RUTA_DATOS}' creada")

# Estructura del examen por materias (ajustar según tu examen)
MAPEO_MATERIAS = {
    'Fisica': range(1, 16),  # Preguntas 1-15
    'Historia': range(16, 31),  # Preguntas 16-30
    'Matematicas': range(31, 61),  # Preguntas 31-60
    'Biologia': range(61, 76),  # Preguntas 61-75
    'Razonamiento': range(76, 81),  # Preguntas 76-80
    'Quimica': range(81, 91),  # Preguntas 81-90
    'Espaniol': range(91, 111)  # Preguntas 91-110
}

# Nombres de columnas para Google Forms
# El sistema buscará flexiblemente estas columnas
COLUMNA_TIMESTAMP = 'Marca temporal'
COLUMNA_EMAIL = 'Nombre de usuario'  # Esta SÍ existe en tu archivo
COLUMNA_NOMBRE = 'Nombre completo'  # Buscará cualquier columna que contenga esto
COLUMNA_GRUPO = 'Grupo'  # Buscará cualquier columna que contenga esto
COLUMNA_PUNTUACION_TOTAL = 'Puntuación total'

# Total de preguntas
TOTAL_PREGUNTAS = 110

# Respuestas válidas
RESPUESTAS_VALIDAS = ['A', 'B', 'C', 'D', 'E']


def validar_configuracion():
    """Valida que la configuración sea consistente."""
    total_preguntas_config = sum(len(list(rango)) for rango in MAPEO_MATERIAS.values())

    if total_preguntas_config != TOTAL_PREGUNTAS:
        print(f"Advertencia: Total de preguntas configuradas ({total_preguntas_config}) "
              f"difiere de TOTAL_PREGUNTAS ({TOTAL_PREGUNTAS})")
        return False

    return True


def mostrar_configuracion():
    """Muestra la configuración actual."""
    print("\n" + "=" * 60)
    print("CONFIGURACION DEL SISTEMA")
    print("=" * 60)
    print(f"Carpeta de datos: {RUTA_DATOS}")
    print(f"\nMaterias configuradas: {len(MAPEO_MATERIAS)}")
    print("-" * 60)

    for materia, rango in MAPEO_MATERIAS.items():
        inicio = rango.start
        fin = rango.stop - 1
        total = len(list(rango))
        print(f"  {materia:15s}: Preguntas {inicio:3d} - {fin:3d} ({total:2d} preguntas)")

    print("-" * 60)
    print(f"Total de preguntas: {TOTAL_PREGUNTAS}")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    mostrar_configuracion()
    validar_configuracion()