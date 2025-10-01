# app/data_loader.py
import pandas as pd
from typing import Optional, Dict, List, Tuple
import re
import os


def cargar_datos(ruta_completa_archivo: str) -> Optional[pd.DataFrame]:
    """Carga los datos desde un archivo CSV y devuelve un DataFrame."""
    if not os.path.exists(ruta_completa_archivo):
        print(f"Error: Archivo no encontrado en '{ruta_completa_archivo}'")
        return None

    try:
        encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'iso-8859-1', 'cp1252']
        df = None

        for encoding in encodings:
            try:
                df = pd.read_csv(ruta_completa_archivo, encoding=encoding)
                print(f"Archivo cargado: {os.path.basename(ruta_completa_archivo)}")
                print(f"  Encoding: {encoding}")
                print(f"  Filas: {len(df)}, Columnas: {len(df.columns)}")
                break
            except (UnicodeDecodeError, UnicodeError):
                continue

        if df is None:
            print(f"Error: No se pudo decodificar el archivo")
            return None

        return df

    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return None


def extraer_columnas_respuestas(df: pd.DataFrame) -> Dict[int, str]:
    """
    Extrae las columnas de RESPUESTAS (no de puntuación ni comentarios).
    Detecta múltiples formatos:
    - "1.", "2.", "3." (Google Forms)
    - "pregunta_1", "pregunta_2", "pregunta_3" (CSV personalizado)
    - "P1", "P2", "P3"
    """
    columnas_respuestas = {}

    for columna in df.columns:
        columna_str = str(columna).strip()

        # Ignorar columnas que contienen [Puntuación] o [Comentarios]
        if '[Puntuación]' in columna_str or '[Comentarios]' in columna_str or '[Puntuacion]' in columna_str:
            continue

        # FORMATO 1: "1.", "2.", "10.", "110."
        match = re.match(r'^(\d+)\.?\s*$', columna_str)
        if match:
            num_pregunta = int(match.group(1))
            columnas_respuestas[num_pregunta] = columna
            continue

        # FORMATO 2: "pregunta_1", "pregunta_2", "pregunta_110"
        match = re.match(r'^pregunta[_\s]*(\d+)$', columna_str, re.IGNORECASE)
        if match:
            num_pregunta = int(match.group(1))
            columnas_respuestas[num_pregunta] = columna
            continue

        # FORMATO 3: "P1", "P2", "P110"
        match = re.match(r'^P(\d+)$', columna_str, re.IGNORECASE)
        if match:
            num_pregunta = int(match.group(1))
            columnas_respuestas[num_pregunta] = columna
            continue

    if columnas_respuestas:
        print(f"Se encontraron {len(columnas_respuestas)} columnas de respuestas")
        preguntas = sorted(columnas_respuestas.keys())
        print(f"  Rango: pregunta {preguntas[0]} a {preguntas[-1]}")
    else:
        print("ADVERTENCIA: No se encontraron columnas de respuestas")
        print("Formatos soportados: '1.', 'pregunta_1', 'P1'")
        print("\nPrimeras 20 columnas disponibles:")
        for i, col in enumerate(df.columns[:20], 1):
            print(f"  {i}. '{col}'")

    return columnas_respuestas


def obtener_respuestas_correctas(df_clave: pd.DataFrame, columnas_respuestas: Dict[int, str]) -> Dict[int, str]:
    """
    Extrae las respuestas correctas de la primera fila del CSV de clave.
    Busca la columna correcta incluso si el formato es diferente.
    """
    if len(df_clave) == 0:
        print("Error: El archivo de clave esta vacio")
        return {}

    respuestas_correctas = {}
    primera_fila = df_clave.iloc[0]

    # Primero extraer las columnas de respuestas del archivo de clave
    columnas_clave = extraer_columnas_respuestas(df_clave)

    print("\nExtrayendo respuestas correctas...")
    print("-" * 60)

    # Para cada pregunta que queremos calificar
    for num_pregunta in sorted(columnas_respuestas.keys()):
        # Buscar la columna correspondiente en la clave
        if num_pregunta in columnas_clave:
            nombre_columna_clave = columnas_clave[num_pregunta]

            if nombre_columna_clave in primera_fila.index:
                valor_raw = primera_fila[nombre_columna_clave]
                respuesta = str(valor_raw).strip().upper()

                # Limpiar la respuesta
                respuesta = respuesta.replace('.', '').replace(',', '').replace(';', '')

                if respuesta in ['A', 'B', 'C', 'D', 'E']:
                    respuestas_correctas[num_pregunta] = respuesta
                    if num_pregunta <= 5 or num_pregunta % 20 == 0:
                        print(f"  Pregunta {num_pregunta:3d}: {respuesta}")
                elif respuesta and respuesta not in ['', 'NAN', 'NONE']:
                    print(f"  Pregunta {num_pregunta:3d}: '{respuesta}' (INVALIDA - ignorada)")

    print("-" * 60)
    print(f"Total de respuestas correctas cargadas: {len(respuestas_correctas)}\n")

    if len(respuestas_correctas) == 0:
        print("ERROR CRITICO: No se cargaron respuestas correctas")
        print("Verifica que la primera fila del CSV de clave contenga las respuestas (A, B, C, D, E)")

    return respuestas_correctas


def validar_estructura_csv(df: pd.DataFrame, es_clave: bool = False) -> Tuple[bool, List[str]]:
    """Valida la estructura del CSV de Google Forms."""
    advertencias = []
    es_valido = True

    # Verificar columnas de preguntas (ignorando [Puntuación] y [Comentarios])
    columnas_pregunta = []
    for col in df.columns:
        col_str = str(col).strip()
        if '[Puntuación]' not in col_str and '[Comentarios]' not in col_str and '[Puntuacion]' not in col_str:
            # Detectar los 3 formatos
            if (re.match(r'^\d+\.?\s*$', col_str) or
                    re.match(r'^pregunta[_\s]*\d+$', col_str, re.IGNORECASE) or
                    re.match(r'^P\d+$', col_str, re.IGNORECASE)):
                columnas_pregunta.append(col)

    if len(columnas_pregunta) == 0:
        print("Error: No se encontraron columnas de preguntas")
        print("Formatos esperados: '1.', 'pregunta_1', 'P1', etc.")
        es_valido = False
    else:
        print(f"Se encontraron {len(columnas_pregunta)} columnas de preguntas")

    if es_clave:
        if len(df) == 0:
            print("Error: El archivo de clave esta vacio")
            es_valido = False
        elif len(df) > 1:
            advertencias.append(f"El archivo de clave tiene {len(df)} filas. Se usara solo la primera.")
            print(f"ADVERTENCIA: {advertencias[-1]}")
    else:
        if len(df) == 0:
            print("Error: El archivo de respuestas esta vacio")
            es_valido = False
        else:
            print(f"Se encontraron {len(df)} respuestas de alumnos")

    return es_valido, advertencias


def obtener_columna_flexible(df: pd.DataFrame, posibles_nombres: List[str]) -> Optional[str]:
    """
    Busca una columna probando varios nombres posibles.
    También busca en columnas que contienen el texto buscado.
    """
    # Primero buscar coincidencia exacta
    for nombre in posibles_nombres:
        if nombre in df.columns:
            return nombre
        nombre_limpio = nombre.strip()
        if nombre_limpio in df.columns:
            return nombre_limpio

    # Si no hay coincidencia exacta, buscar columnas que CONTIENEN el texto
    for nombre in posibles_nombres:
        nombre_lower = nombre.lower().strip()
        for col in df.columns:
            col_lower = str(col).lower().strip()
            # Buscar si la columna contiene el nombre buscado
            if nombre_lower in col_lower:
                # Pero evitar columnas de puntuación/comentarios
                if '[puntuación]' not in col_lower and '[comentarios]' not in col_lower and '[puntuacion]' not in col_lower:
                    return col

    return None


def limpiar_respuesta(respuesta: any) -> str:
    """Limpia y normaliza una respuesta."""
    if pd.isna(respuesta):
        return ''

    respuesta_str = str(respuesta).strip().upper()
    respuesta_str = respuesta_str.replace('.', '').replace(',', '').replace(';', '')

    return respuesta_str


def diagnosticar_csv(ruta_archivo: str):
    """Función de diagnóstico para entender la estructura del CSV."""
    print(f"\n{'=' * 60}")
    print(f"DIAGNOSTICO DEL ARCHIVO")
    print(f"{'=' * 60}")
    print(f"Archivo: {ruta_archivo}\n")

    df = cargar_datos(ruta_archivo)
    if df is None:
        return

    print(f"Informacion general:")
    print(f"  - Filas: {len(df)}")
    print(f"  - Columnas: {len(df.columns)}")

    print(f"\nPrimeras 20 columnas:")
    for i, col in enumerate(df.columns[:20], 1):
        muestra = df[col].iloc[0] if len(df) > 0 else 'N/A'
        print(f"  {i:3d}. '{col}' = '{muestra}'")

    # Detectar columnas de respuestas
    columnas_resp = extraer_columnas_respuestas(df)
    if columnas_resp:
        print(f"\nColumnas de respuestas detectadas: {len(columnas_resp)}")
        preguntas = sorted(columnas_resp.keys())
        print(f"  Primera pregunta: {preguntas[0]}")
        print(f"  Ultima pregunta: {preguntas[-1]}")

    print("=" * 60 + "\n")


if __name__ == "__main__":
    print("Modulo data_loader.py")
    print("Funciones para cargar y validar archivos CSV")