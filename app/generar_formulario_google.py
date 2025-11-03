import csv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import os

SCOPES = [
    'https://www.googleapis.com/auth/forms',
    'https://www.googleapis.com/auth/drive'
]


def autenticar():
    """Autentica con Google y retorna las credenciales"""
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    return creds


def crear_hoja_respuestas(csv_file, titulo_formulario):
    """Crea un formulario de Google con preguntas desde CSV"""

    # Autenticar primero
    creds = autenticar()

    # Crear el servicio con las credenciales
    service = build('forms', 'v1', credentials=creds)

    # Crear formulario
    form = {
        'info': {
            'title': titulo_formulario,
        }
    }

    result = service.forms().create(body=form).execute()
    form_id = result['formId']
    print(f"✓ Formulario creado: {form_id}")

    # Leer CSV y crear preguntas
    requests = []

    with open(csv_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for idx, row in enumerate(reader):
            requests.append({
                'createItem': {
                    'item': {
                        'title': f"Pregunta {row['numero']}",
                        'questionItem': {
                            'question': {
                                'required': True,
                                'choiceQuestion': {
                                    'type': 'RADIO',
                                    'options': [
                                        {'value': row['opcion_a']},
                                        {'value': row['opcion_b']},
                                        {'value': row['opcion_c']},
                                        {'value': row['opcion_d']}
                                    ]
                                }
                            }
                        }
                    },
                    'location': {'index': idx}
                }
            })

    # Agregar todas las preguntas en lotes (de 50 en 50)
    batch_size = 50
    total_preguntas = len(requests)

    for i in range(0, total_preguntas, batch_size):
        batch = requests[i:i + batch_size]
        try:
            service.forms().batchUpdate(
                formId=form_id,
                body={'requests': batch}
            ).execute()
            print(f"✓ {len(batch)} preguntas agregadas ({min(i + len(batch), total_preguntas)}/{total_preguntas})")
        except Exception as e:
            print(f"✗ Error en lote {i // batch_size + 1}: {e}")
            raise

    return form_id, f"https://docs.google.com/forms/d/{form_id}/edit"


if __name__ == '__main__':
    form_id, url = crear_hoja_respuestas('examen.csv', 'Examen - 120 Preguntas')
    print(f"\n✓ Formulario listo en:\n{url}")