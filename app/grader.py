# app/grader.py
import pandas as pd
from typing import List, Dict, Any, Optional
from data_loader import (extraer_columnas_respuestas, obtener_respuestas_correctas,
                         limpiar_respuesta, obtener_columna_flexible)


def procesar_calificaciones_google_forms(
        clave_df: pd.DataFrame,
        respuestas_df: pd.DataFrame,
        mapeo_materias: Dict[str, range],
        columna_nombre: str,
        columna_email: str = 'Nombre de usuario',
        columna_grupo: str = 'Grupo '
) -> List[Dict[str, Any]]:
    """
    Procesa las calificaciones de un CSV de Google Forms.

    Args:
        clave_df: DataFrame con las respuestas correctas
        respuestas_df: DataFrame con las respuestas de los alumnos
        mapeo_materias: Diccionario con los rangos de preguntas por materia
        columna_nombre: Nombre de la columna que contiene el nombre del alumno
        columna_email: Nombre de la columna que contiene el email
        columna_grupo: Nombre de la columna que contiene el grupo

    Returns:
        Lista de diccionarios con los resultados por alumno
    """
    print("\n" + "=" * 60)
    print("🔄 PROCESANDO CALIFICACIONES")
    print("=" * 60)

    # Extraer las columnas de respuestas
    columnas_respuestas = extraer_columnas_respuestas(respuestas_df)

    if not columnas_respuestas:
        print("❌ Error: No se encontraron columnas de respuestas")
        return []

    print(f"📋 Total de preguntas detectadas: {len(columnas_respuestas)}")

    # Obtener respuestas correctas
    respuestas_correctas = obtener_respuestas_correctas(clave_df, columnas_respuestas)

    if not respuestas_correctas:
        print("❌ Error: No se pudieron cargar las respuestas correctas")
        return []

    print(f"✅ Respuestas correctas cargadas: {len(respuestas_correctas)}")

    # Buscar columnas de forma flexible (con y sin espacios)
    col_nombre_encontrada = obtener_columna_flexible(respuestas_df,
                                                     [columna_nombre, columna_nombre.strip(), 'Nombre completo',
                                                      'Nombre'])
    col_email_encontrada = obtener_columna_flexible(respuestas_df,
                                                    [columna_email, columna_email.strip(), 'Email', 'Correo'])
    col_grupo_encontrada = obtener_columna_flexible(respuestas_df,
                                                    [columna_grupo, columna_grupo.strip(), 'Grupo'])

    print(f"\n📊 Procesando {len(respuestas_df)} alumno(s)...")
    print("-" * 60)

    resultados_finales = []

    # Procesar cada alumno
    for idx, alumno in respuestas_df.iterrows():
        # Obtener información del alumno
        nombre_alumno = obtener_valor_columna(alumno, col_nombre_encontrada, f'Alumno_{idx + 1}')
        email_alumno = obtener_valor_columna(alumno, col_email_encontrada, 'Sin email')
        grupo_alumno = obtener_valor_columna(alumno, col_grupo_encontrada, 'Sin grupo')

        print(f"  Procesando: {nombre_alumno}")

        # Calcular aciertos
        aciertos_totales = []
        errores = []
        sin_responder = []

        for num_pregunta, nombre_col in sorted(columnas_respuestas.items()):
            if num_pregunta in respuestas_correctas:
                respuesta_alumno = limpiar_respuesta(alumno.get(nombre_col, ''))
                respuesta_correcta = respuestas_correctas[num_pregunta]

                if not respuesta_alumno:
                    sin_responder.append(num_pregunta)
                elif respuesta_alumno == respuesta_correcta:
                    aciertos_totales.append(num_pregunta)
                else:
                    errores.append(num_pregunta)

        # Calcular calificaciones por materia
        reporte_alumno = {
            'nombre': nombre_alumno,
            'email': email_alumno,
            'grupo': grupo_alumno,
            'calificaciones': {},
            'estadisticas': {
                'aciertos': aciertos_totales,
                'errores': errores,
                'sin_responder': sin_responder
            }
        }

        # Procesar cada materia
        for materia, rango in mapeo_materias.items():
            # Filtrar solo las preguntas que existen en este examen
            preguntas_materia = [p for p in rango if p in respuestas_correctas]
            aciertos_materia = [p for p in aciertos_totales if p in preguntas_materia]
            errores_materia = [p for p in errores if p in preguntas_materia]
            sin_resp_materia = [p for p in sin_responder if p in preguntas_materia]

            total_preguntas = len(preguntas_materia)
            num_aciertos = len(aciertos_materia)
            porcentaje = (num_aciertos / total_preguntas * 100) if total_preguntas > 0 else 0

            # Calcular calificación numérica (escala 0-10)
            calificacion_numerica = (num_aciertos / total_preguntas * 10) if total_preguntas > 0 else 0

            reporte_alumno['calificaciones'][materia] = {
                'aciertos': num_aciertos,
                'errores': len(errores_materia),
                'sin_responder': len(sin_resp_materia),
                'total': total_preguntas,
                'porcentaje': round(porcentaje, 2),
                'calificacion': round(calificacion_numerica, 2),
                'preguntas_correctas': sorted(aciertos_materia),
                'preguntas_incorrectas': sorted(errores_materia),
                'preguntas_sin_responder': sorted(sin_resp_materia)
            }

        # Calcular totales generales
        total_preguntas_examen = len(respuestas_correctas)
        total_aciertos = len(aciertos_totales)
        total_errores = len(errores)
        total_sin_responder = len(sin_responder)

        reporte_alumno['total_aciertos'] = total_aciertos
        reporte_alumno['total_errores'] = total_errores
        reporte_alumno['total_sin_responder'] = total_sin_responder
        reporte_alumno['total_preguntas'] = total_preguntas_examen
        reporte_alumno['porcentaje_global'] = round((total_aciertos / total_preguntas_examen * 100),
                                                    2) if total_preguntas_examen > 0 else 0
        reporte_alumno['calificacion_global'] = round((total_aciertos / total_preguntas_examen * 10),
                                                      2) if total_preguntas_examen > 0 else 0

        resultados_finales.append(reporte_alumno)

    print("-" * 60)
    print(f"✅ Procesamiento completado: {len(resultados_finales)} alumno(s)")
    print("=" * 60 + "\n")

    return resultados_finales


def obtener_valor_columna(fila: pd.Series, nombre_columna: Optional[str], valor_default: str) -> str:
    """
    Obtiene el valor de una columna, manejando valores nulos y espacios.

    Args:
        fila: Fila del DataFrame (Series)
        nombre_columna: Nombre de la columna a buscar
        valor_default: Valor por defecto si no se encuentra

    Returns:
        Valor de la columna limpio o valor por defecto
    """
    if nombre_columna is None:
        return valor_default

    if nombre_columna in fila.index:
        valor = fila[nombre_columna]
        if pd.notna(valor):
            valor_str = str(valor).strip()
            if valor_str and valor_str.lower() not in ['nan', 'none', '']:
                return valor_str

    return valor_default


def calcular_estadisticas_grupo(resultados: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Calcula estadísticas generales del grupo.

    Args:
        resultados: Lista de resultados de alumnos

    Returns:
        Diccionario con estadísticas del grupo
    """
    if not resultados:
        return {}

    # Calcular promedios
    promedio_global = sum(r['porcentaje_global'] for r in resultados) / len(resultados)
    promedio_calificacion = sum(r['calificacion_global'] for r in resultados) / len(resultados)

    # Calcular máximos y mínimos
    max_porcentaje = max(r['porcentaje_global'] for r in resultados)
    min_porcentaje = min(r['porcentaje_global'] for r in resultados)

    # Encontrar mejor y peor alumno
    mejor_alumno = max(resultados, key=lambda x: x['porcentaje_global'])
    peor_alumno = min(resultados, key=lambda x: x['porcentaje_global'])

    # Calcular estadísticas por materia
    estadisticas_materias = {}
    if resultados:
        for materia in resultados[0]['calificaciones'].keys():
            porcentajes = [r['calificaciones'][materia]['porcentaje'] for r in resultados]
            estadisticas_materias[materia] = {
                'promedio': round(sum(porcentajes) / len(porcentajes), 2),
                'maximo': round(max(porcentajes), 2),
                'minimo': round(min(porcentajes), 2)
            }

    return {
        'total_alumnos': len(resultados),
        'promedio_global': round(promedio_global, 2),
        'promedio_calificacion': round(promedio_calificacion, 2),
        'porcentaje_maximo': round(max_porcentaje, 2),
        'porcentaje_minimo': round(min_porcentaje, 2),
        'mejor_alumno': mejor_alumno['nombre'],
        'peor_alumno': peor_alumno['nombre'],
        'materias': estadisticas_materias
    }


def generar_reporte_detallado(resultado: Dict[str, Any]) -> str:
    """
    Genera un reporte detallado en texto para un alumno.

    Args:
        resultado: Diccionario con los resultados de un alumno

    Returns:
        String con el reporte detallado
    """
    reporte = []
    reporte.append("=" * 70)
    reporte.append(f"REPORTE DETALLADO: {resultado['nombre']}")
    reporte.append("=" * 70)
    reporte.append(f"Email: {resultado['email']}")
    reporte.append(f"Grupo: {resultado['grupo']}")
    reporte.append("")
    reporte.append(f"Calificación Global: {resultado['calificacion_global']}/10 ({resultado['porcentaje_global']}%)")
    reporte.append(f"Aciertos: {resultado['total_aciertos']}/{resultado['total_preguntas']}")
    reporte.append(f"Errores: {resultado['total_errores']}")
    reporte.append(f"Sin responder: {resultado['total_sin_responder']}")
    reporte.append("")
    reporte.append("DESGLOSE POR MATERIA:")
    reporte.append("-" * 70)

    for materia, datos in sorted(resultado['calificaciones'].items()):
        reporte.append(f"\n{materia}:")
        reporte.append(f"  Calificación: {datos['calificacion']}/10 ({datos['porcentaje']}%)")
        reporte.append(f"  Aciertos: {datos['aciertos']}/{datos['total']}")
        reporte.append(f"  Errores: {datos['errores']}")
        reporte.append(f"  Sin responder: {datos['sin_responder']}")

    reporte.append("")
    reporte.append("=" * 70)

    return "\n".join(reporte)


if __name__ == "__main__":
    print("Módulo grader.py")
    print("Este módulo contiene funciones para calificar exámenes")