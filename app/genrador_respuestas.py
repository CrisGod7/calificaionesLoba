import random
import csv


def generar_respuestas_aleatorias(num_preguntas=120):
    """
    Genera respuestas aleatorias para todas las preguntas
    Crea 2 archivos:
    1. clave_respuestas.csv - Para ver las respuestas correctas de forma clara
    2. examen.csv - Para generar el Google formulario
    """

    respuestas_aleatorias = [random.choice(['A', 'B', 'C', 'D']) for _ in range(num_preguntas)]

    # 1. Crear clave_respuestas.csv (para ver las respuestas de forma clara)
    print("📝 Generando clave de respuestas...")
    with open('clave_respuestas.csv', 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Pregunta', 'Respuesta Correcta'])
        for i, respuesta in enumerate(respuestas_aleatorias, 1):
            writer.writerow([i, respuesta])

    print(f"✓ clave_respuestas.csv creado ({num_preguntas} preguntas)")

    # 2. Crear examen.csv (para generar el formulario de Google)
    print("📋 Generando CSV para formulario...")
    with open('examen.csv', 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['numero', 'opcion_a', 'opcion_b', 'opcion_c', 'opcion_d', 'respuesta_correcta'])
        for i, respuesta in enumerate(respuestas_aleatorias, 1):
            writer.writerow([i, 'A', 'B', 'C', 'D', respuesta])

    print(f"✓ examen.csv creado ({num_preguntas} preguntas)")

    # Mostrar resumen
    print("\n" + "=" * 50)
    print("📊 RESUMEN DE RESPUESTAS")
    print("=" * 50)

    conteo = {'A': respuestas_aleatorias.count('A'),
              'B': respuestas_aleatorias.count('B'),
              'C': respuestas_aleatorias.count('C'),
              'D': respuestas_aleatorias.count('D')}

    for opcion, cantidad in sorted(conteo.items()):
        porcentaje = (cantidad / num_preguntas) * 100
        print(f"Opción {opcion}: {cantidad:3d} preguntas ({porcentaje:5.1f}%)")

    print("=" * 50)
    print("\n✅ Archivos generados:")
    print("   📄 clave_respuestas.csv - Para ver las respuestas correctas")
    print("   📄 examen.csv - Para generar el Google Forms")
    print("\nPróximo paso: python generar_formulario_google.py")


if __name__ == '__main__':
    generar_respuestas_aleatorias(120)