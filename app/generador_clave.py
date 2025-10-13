# app/generador_clave.py
import pandas as pd
import os
from typing import List, Optional
from datetime import datetime


class GeneradorClave:
    """Clase para generar archivos CSV de clave de respuestas."""

    def __init__(self, total_preguntas: int = 120):
        self.total_preguntas = total_preguntas
        self.respuestas = [''] * total_preguntas

    def establecer_respuestas_desde_texto(self, texto: str) -> int:
        """
        Establece las respuestas desde un texto.
        Acepta varios formatos:
        - Una línea por respuesta: "A\nB\nC\n..."
        - Respuestas con viñetas: "* A\n* B\n* C\n..."
        - Respuestas numeradas: "1. A\n2. B\n3. C\n..."
        - Respuestas separadas por comas: "A, B, C, ..."

        Returns:
            Número de respuestas válidas procesadas
        """
        import re

        # Limpiar el texto
        texto = texto.strip()

        if not texto:
            return 0

        # Detectar el formato
        lineas = texto.split('\n')
        respuestas_procesadas = []

        for linea in lineas:
            linea = linea.strip()
            if not linea:
                continue

            # Remover viñetas, números, asteriscos, puntos
            linea = re.sub(r'^[\*\-•]\s*', '', linea)  # Viñetas
            linea = re.sub(r'^\d+[\.\)]\s*', '', linea)  # Números
            linea = linea.strip()

            # Extraer solo la letra (A, B, C, D, E)
            match = re.search(r'^([A-Ea-e])\b', linea)
            if match:
                respuestas_procesadas.append(match.group(1).upper())

        # Si no se encontraron respuestas con el método anterior,
        # intentar separar por comas
        if not respuestas_procesadas:
            partes = re.split(r'[,;|\s]+', texto)
            for parte in partes:
                parte = parte.strip().upper()
                if parte in ['A', 'B', 'C', 'D', 'E']:
                    respuestas_procesadas.append(parte)

        # Asignar las respuestas
        for i, respuesta in enumerate(respuestas_procesadas):
            if i < self.total_preguntas:
                self.respuestas[i] = respuesta

        return len(respuestas_procesadas)

    def establecer_respuesta(self, numero_pregunta: int, respuesta: str) -> bool:
        """
        Establece una respuesta individual.

        Args:
            numero_pregunta: Número de pregunta (1-based)
            respuesta: Letra de la respuesta (A, B, C, D, E)

        Returns:
            True si se estableció correctamente
        """
        if not 1 <= numero_pregunta <= self.total_preguntas:
            return False

        respuesta = respuesta.strip().upper()
        if respuesta not in ['A', 'B', 'C', 'D', 'E']:
            return False

        self.respuestas[numero_pregunta - 1] = respuesta
        return True

    def obtener_respuesta(self, numero_pregunta: int) -> str:
        """Obtiene la respuesta de una pregunta específica."""
        if 1 <= numero_pregunta <= self.total_preguntas:
            return self.respuestas[numero_pregunta - 1]
        return ''

    def contar_respuestas_validas(self) -> int:
        """Cuenta cuántas respuestas válidas hay."""
        return sum(1 for r in self.respuestas if r)

    def obtener_respuestas_faltantes(self) -> List[int]:
        """Retorna lista de números de preguntas sin respuesta."""
        return [i + 1 for i, r in enumerate(self.respuestas) if not r]

    def limpiar_respuestas(self):
        """Limpia todas las respuestas."""
        self.respuestas = [''] * self.total_preguntas

    def generar_csv(self, ruta_salida: str, formato: str = 'pregunta_n') -> bool:
        """
        Genera el archivo CSV de clave de respuestas.

        Args:
            ruta_salida: Ruta donde guardar el archivo
            formato: Formato de columnas ('pregunta_n', 'Pn', 'n.')

        Returns:
            True si se generó correctamente
        """
        try:
            # Crear el diccionario de datos
            datos = {'Nombre': ['CLAVE']}

            for i in range(1, self.total_preguntas + 1):
                if formato == 'pregunta_n':
                    nombre_col = f'pregunta_{i}'
                elif formato == 'Pn':
                    nombre_col = f'P{i}'
                else:  # formato 'n.'
                    nombre_col = f'{i}.'

                datos[nombre_col] = [self.respuestas[i - 1]]

            # Crear DataFrame y exportar
            df = pd.DataFrame(datos)
            df.to_csv(ruta_salida, index=False, encoding='utf-8-sig')

            return True
        except Exception as e:
            print(f"Error al generar CSV: {e}")
            return False

    def generar_reporte_texto(self) -> str:
        """Genera un reporte en texto de las respuestas."""
        reporte = []
        reporte.append("=" * 70)
        reporte.append("CLAVE DE RESPUESTAS")
        reporte.append("=" * 70)
        reporte.append(f"Total de preguntas: {self.total_preguntas}")
        reporte.append(f"Respuestas establecidas: {self.contar_respuestas_validas()}")
        reporte.append(f"Respuestas faltantes: {self.total_preguntas - self.contar_respuestas_validas()}")
        reporte.append("")

        # Mostrar respuestas en columnas
        reporte.append("RESPUESTAS:")
        reporte.append("-" * 70)

        # Mostrar en 5 columnas
        for fila in range(0, self.total_preguntas, 5):
            linea_parts = []
            for col in range(5):
                idx = fila + col
                if idx < self.total_preguntas:
                    num_pregunta = idx + 1
                    respuesta = self.respuestas[idx] if self.respuestas[idx] else '?'
                    linea_parts.append(f"{num_pregunta:3d}. {respuesta}")
            reporte.append("  ".join(linea_parts))

        reporte.append("")

        # Mostrar respuestas faltantes si las hay
        faltantes = self.obtener_respuestas_faltantes()
        if faltantes:
            reporte.append("PREGUNTAS SIN RESPUESTA:")
            reporte.append("-" * 70)
            # Mostrar en grupos de 10
            for i in range(0, len(faltantes), 10):
                grupo = faltantes[i:i + 10]
                reporte.append("  " + ", ".join(str(n) for n in grupo))
            reporte.append("")

        reporte.append("=" * 70)
        return "\n".join(reporte)


def crear_clave_interactiva(total_preguntas: int = 120) -> Optional[GeneradorClave]:
    """
    Función interactiva para crear una clave de respuestas desde la consola.

    Returns:
        GeneradorClave con las respuestas establecidas o None si se cancela
    """
    print("\n" + "=" * 70)
    print("GENERADOR DE CLAVE DE RESPUESTAS")
    print("=" * 70)
    print(f"Total de preguntas: {total_preguntas}")
    print("\nOpciones:")
    print("  1. Ingresar respuestas en bloque (una por línea)")
    print("  2. Ingresar respuestas individualmente")
    print("  3. Cancelar")

    opcion = input("\nSeleccione una opción (1-3): ").strip()

    generador = GeneradorClave(total_preguntas)

    if opcion == '1':
        print("\nIngrese todas las respuestas (una por línea).")
        print("Puede usar formatos como:")
        print("  A")
        print("  * B")
        print("  1. C")
        print("Presione Ctrl+D (Linux/Mac) o Ctrl+Z (Windows) cuando termine:\n")

        lineas = []
        try:
            while True:
                linea = input()
                lineas.append(linea)
        except EOFError:
            pass

        texto = '\n'.join(lineas)
        num_procesadas = generador.establecer_respuestas_desde_texto(texto)

        print(f"\nSe procesaron {num_procesadas} respuestas.")

    elif opcion == '2':
        print("\nIngrese las respuestas una por una.")
        print("Formato: A, B, C, D o E")
        print("Deje en blanco para saltar una pregunta.\n")

        for i in range(1, total_preguntas + 1):
            while True:
                respuesta = input(f"Pregunta {i}: ").strip().upper()

                if not respuesta:
                    break

                if respuesta in ['A', 'B', 'C', 'D', 'E']:
                    generador.establecer_respuesta(i, respuesta)
                    break
                else:
                    print("  Respuesta inválida. Use A, B, C, D o E.")
    else:
        print("Operación cancelada.")
        return None

    return generador


if __name__ == "__main__":
    # Ejemplo de uso
    print("Ejemplo de uso del GeneradorClave")

    # Crear generador
    gen = GeneradorClave(total_preguntas=30)

    # Ejemplo con texto
    texto_ejemplo = """
    * B
    * B
    * A
    * A
    * C
    * A
    * A
    * A
    * C
    * D
    * C
    * D
    * D
    * B
    * A
    * C
    * A
    * A
    * B
    * C
    * C
    * C
    * C
    * A
    * A
    * B
    * C
    * B
    * D
    * B
    """

    num = gen.establecer_respuestas_desde_texto(texto_ejemplo)
    print(f"Respuestas procesadas: {num}")
    print(gen.generar_reporte_texto())

    # Generar CSV
    ruta = "clave_ejemplo.csv"
    if gen.generar_csv(ruta):
        print(f"\nArchivo generado: {ruta}")