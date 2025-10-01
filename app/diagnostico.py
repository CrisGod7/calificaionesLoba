#!/usr/bin/env python3
# diagnostico.py - Script para diagnosticar archivos CSV

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

from data_loader import cargar_datos, extraer_columnas_respuestas
import pandas as pd


def diagnosticar_archivo_clave(ruta_archivo):
    """Diagnostica el archivo de clave de respuestas."""
    print("\n" + "=" * 70)
    print("DIAGNOSTICO DEL ARCHIVO DE CLAVE")
    print("=" * 70)
    print(f"Archivo: {ruta_archivo}\n")

    df = cargar_datos(ruta_archivo)
    if df is None:
        return

    print(f"\n1. INFORMACION GENERAL")
    print("-" * 70)
    print(f"  Filas: {len(df)}")
    print(f"  Columnas: {len(df.columns)}")

    print(f"\n2. PRIMERAS 30 COLUMNAS")
    print("-" * 70)
    for i, col in enumerate(df.columns[:30], 1):
        valor = df[col].iloc[0] if len(df) > 0 else 'N/A'
        print(f"  {i:3d}. '{col}' = '{valor}'")

    if len(df.columns) > 30:
        print(f"  ... y {len(df.columns) - 30} columnas mas")

    # Extraer columnas de respuestas
    print(f"\n3. COLUMNAS DE RESPUESTAS DETECTADAS")
    print("-" * 70)
    columnas_resp = extraer_columnas_respuestas(df)

    if columnas_resp:
        print(f"Total: {len(columnas_resp)} columnas")
        print("\nPrimeras 20 preguntas y sus respuestas:")
        for num_p in sorted(columnas_resp.keys())[:20]:
            col_name = columnas_resp[num_p]
            valor = df[col_name].iloc[0] if len(df) > 0 else 'N/A'
            print(f"  Pregunta {num_p:3d}: columna='{col_name}' valor='{valor}'")

    # Verificar respuestas validas
    print(f"\n4. VALIDACION DE RESPUESTAS")
    print("-" * 70)

    if len(df) == 0:
        print("  ERROR: El archivo esta vacio")
        return

    primera_fila = df.iloc[0]
    respuestas_validas = 0
    respuestas_invalidas = []
    respuestas_vacias = []

    for num_p, col_name in columnas_resp.items():
        if col_name in primera_fila.index:
            valor = str(primera_fila[col_name]).strip().upper()
            valor_limpio = valor.replace('.', '').replace(',', '')

            if valor_limpio in ['A', 'B', 'C', 'D', 'E']:
                respuestas_validas += 1
            elif not valor_limpio or valor_limpio in ['NAN', 'NONE', '']:
                respuestas_vacias.append(num_p)
            else:
                respuestas_invalidas.append((num_p, valor))

    print(f"  Respuestas validas (A-E): {respuestas_validas}")
    print(f"  Respuestas vacias: {len(respuestas_vacias)}")
    print(f"  Respuestas invalidas: {len(respuestas_invalidas)}")

    if respuestas_vacias:
        print(f"\n  Preguntas sin respuesta: {respuestas_vacias[:10]}")
        if len(respuestas_vacias) > 10:
            print(f"  ... y {len(respuestas_vacias) - 10} mas")

    if respuestas_invalidas:
        print(f"\n  Preguntas con respuestas invalidas:")
        for num_p, valor in respuestas_invalidas[:10]:
            print(f"    Pregunta {num_p}: '{valor}'")
        if len(respuestas_invalidas) > 10:
            print(f"  ... y {len(respuestas_invalidas) - 10} mas")

    print(f"\n5. RECOMENDACIONES")
    print("-" * 70)

    if respuestas_validas == 0:
        print("  ERROR CRITICO: No hay respuestas validas en el archivo de clave")
        print("\n  Para que funcione correctamente:")
        print("    1. La primera fila debe contener las respuestas correctas")
        print("    2. Las respuestas deben ser letras: A, B, C, D o E")
        print("    3. Cada pregunta (1., 2., 3., etc.) debe tener su respuesta")
        print("\n  Ejemplo de como debe verse:")
        print("    Columna '1.' -> Valor 'A'")
        print("    Columna '2.' -> Valor 'B'")
        print("    Columna '3.' -> Valor 'C'")
    elif respuestas_validas < len(columnas_resp) * 0.8:
        print(f"  ADVERTENCIA: Solo {respuestas_validas}/{len(columnas_resp)} preguntas tienen respuestas validas")
        print("  Verifica que todas las preguntas tengan respuesta en la clave")
    else:
        print(f"  OK: {respuestas_validas}/{len(columnas_resp)} preguntas tienen respuestas validas")

    print("\n" + "=" * 70 + "\n")


def diagnosticar_archivo_respuestas(ruta_archivo):
    """Diagnostica el archivo de respuestas de alumnos."""
    print("\n" + "=" * 70)
    print("DIAGNOSTICO DEL ARCHIVO DE RESPUESTAS")
    print("=" * 70)
    print(f"Archivo: {ruta_archivo}\n")

    df = cargar_datos(ruta_archivo)
    if df is None:
        return

    print(f"\n1. INFORMACION GENERAL")
    print("-" * 70)
    print(f"  Alumnos (filas): {len(df)}")
    print(f"  Columnas totales: {len(df.columns)}")

    # Buscar columnas de identificacion
    print(f"\n2. COLUMNAS DE IDENTIFICACION")
    print("-" * 70)

    posibles_nombres = ['Nombre completo', 'Nombre completo ', 'Nombre', 'Alumno']
    posibles_email = ['Nombre de usuario', 'Email', 'Correo', 'Correo electronico']
    posibles_grupo = ['Grupo', 'Grupo ', 'Seccion']

    col_nombre = None
    col_email = None
    col_grupo = None

    for col in df.columns:
        if any(p in col for p in posibles_nombres):
            col_nombre = col
        if any(p in col for p in posibles_email):
            col_email = col
        if any(p in col for p in posibles_grupo):
            col_grupo = col

    print(f"  Columna de nombre: {col_nombre if col_nombre else 'NO ENCONTRADA'}")
    print(f"  Columna de email: {col_email if col_email else 'NO ENCONTRADA'}")
    print(f"  Columna de grupo: {col_grupo if col_grupo else 'NO ENCONTRADA'}")

    # Extraer columnas de respuestas
    print(f"\n3. COLUMNAS DE RESPUESTAS")
    print("-" * 70)
    columnas_resp = extraer_columnas_respuestas(df)
    print(f"  Total: {len(columnas_resp)} columnas de respuestas")

    # Mostrar muestra de alumnos
    print(f"\n4. MUESTRA DE ALUMNOS")
    print("-" * 70)
    for idx in range(min(3, len(df))):
        alumno = df.iloc[idx]
        nombre = alumno[col_nombre] if col_nombre else f'Alumno_{idx + 1}'
        email = alumno[col_email] if col_email else 'Sin email'
        grupo = alumno[col_grupo] if col_grupo else 'Sin grupo'
        print(f"\n  Alumno {idx + 1}:")
        print(f"    Nombre: {nombre}")
        print(f"    Email: {email}")
        print(f"    Grupo: {grupo}")

        # Mostrar primeras respuestas
        if columnas_resp:
            print(f"    Primeras 5 respuestas:")
            for num_p in sorted(columnas_resp.keys())[:5]:
                col_name = columnas_resp[num_p]
                respuesta = alumno[col_name] if col_name in alumno.index else 'N/A'
                print(f"      P{num_p}: {respuesta}")

    print("\n" + "=" * 70 + "\n")


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        archivo = sys.argv[1]
        if 'clave' in archivo.lower():
            diagnosticar_archivo_clave(archivo)
        else:
            diagnosticar_archivo_respuestas(archivo)
    else:
        print("\nUso:")
        print("  python diagnostico.py <archivo.csv>")
        print("\nEjemplos:")
        print("  python diagnostico.py data/clave_respuestas.csv")
        print("  python diagnostico.py data/simulacro.csv")
        print()