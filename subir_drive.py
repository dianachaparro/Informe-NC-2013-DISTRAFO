"""Sube informes locales a una carpeta de Drive mediante OAuth de escritorio."""
import argparse
import json
import os
from pathlib import Path
import re
import sys
from urllib.parse import urlparse

# Permite gestionar una carpeta existente del usuario. drive.file requeriría
# que se seleccionara previamente esa carpeta mediante Google Picker.
SCOPES = ['https://www.googleapis.com/auth/drive']
FOLDER_MIME = 'application/vnd.google-apps.folder'
EXCEL_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def folder_id(value):
    if value.startswith('https://'):
        url = urlparse(value)
        if url.hostname != 'drive.google.com':
            raise ValueError('La carpeta debe pertenecer a drive.google.com.')
        match = re.search(r'/folders/([A-Za-z0-9_-]+)(?:/|$)', url.path)
        if not match:
            raise ValueError('Se requiere un enlace de carpeta de Drive.')
        return match[1]
    if not re.fullmatch(r'[A-Za-z0-9_-]+', value):
        raise ValueError('ID de carpeta inválido.')
    return value


def validate_inputs(client_file, reports):
    if not client_file.is_file():
        raise ValueError('No se encontró credentials.json en la ruta indicada.')
    try:
        config = json.loads(client_file.read_text(encoding='utf-8'))
    except (ValueError, OSError):
        raise ValueError('No se pudo leer la configuración OAuth.') from None
    if not isinstance(config, dict) or not isinstance(config.get('installed'), dict):
        raise ValueError('Se requiere el JSON de un cliente OAuth de App de escritorio.')
    if not reports:
        raise ValueError('Selecciona al menos un Excel para subir.')
    for report in reports:
        if not report.is_file() or report.suffix.lower() != '.xlsx':
            raise ValueError(f'No se encontró un informe Excel válido: {report.name}')
    if len({p.name for p in reports}) != len(reports):
        raise ValueError('Los informes deben tener nombres de archivo distintos.')


def authenticate(client_file, token_file):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.exceptions import RefreshError

    credentials = None
    if token_file.exists():
        try:
            credentials = Credentials.from_authorized_user_file(str(token_file), SCOPES)
        except ValueError:
            raise ValueError('La autorización local no es válida; mueve el token y vuelve a conectar.') from None
    if credentials and credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(Request())
        except RefreshError:
            credentials = None
    if not credentials or not credentials.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(client_file), SCOPES)
        credentials = flow.run_local_server(host='localhost', port=0, open_browser=True,
                                           authorization_prompt_message='Se abrirá Google en tu navegador para autorizar el programa.',
                                           success_message='Conexión completada. Puedes cerrar esta pestaña.')
    token_file.parent.mkdir(parents=True, exist_ok=True)
    descriptor = os.open(token_file, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(descriptor, 'w', encoding='utf-8') as token:
        token.write(credentials.to_json())
    return credentials


def upload_reports(service, parent, reports, media_factory, subfolder='Informes de NC'):
    """Comprueba capacidad de escritura; crea archivos nuevos sin reemplazar otros."""
    folder = service.files().get(fileId=parent,
                                 fields='id,name,mimeType,capabilities(canAddChildren)',
                                 supportsAllDrives=True).execute()
    if folder.get('mimeType') != FOLDER_MIME:
        raise ValueError('El destino no es una carpeta.')
    if not folder.get('capabilities', {}).get('canAddChildren'):
        raise ValueError('La cuenta autorizada no tiene permiso para añadir archivos a esta carpeta.')
    safe_name = subfolder.replace('\\', '\\\\').replace("'", "\\'")
    existing = service.files().list(
        q=f"'{parent}' in parents and trashed = false and mimeType = '{FOLDER_MIME}' and name = '{safe_name}'",
        fields='files(id,name),nextPageToken', supportsAllDrives=True,
        includeItemsFromAllDrives=True, pageSize=100).execute()
    matches = existing.get('files', [])
    if len(matches) > 1 or existing.get('nextPageToken'):
        raise ValueError('Hay varias subcarpetas del mismo nombre; usa otro nombre para evitar ambigüedad.')
    if matches:
        child = matches[0]['id']
    else:
        child = service.files().create(body={'name': subfolder, 'mimeType': FOLDER_MIME,
                                            'parents': [parent]}, fields='id',
                                       supportsAllDrives=True).execute()['id']
    results = []
    for report in reports:
        result = service.files().create(body={'name': report.name, 'parents': [child]},
                                        media_body=media_factory(str(report), mimetype=EXCEL_MIME, resumable=True),
                                        fields='id,name,webViewLink', supportsAllDrives=True).execute()
        # Imprimir cada resultado permite recuperar cargas parciales si una posterior falla.
        link = result.get('webViewLink') or f'https://drive.google.com/file/d/{result["id"]}/view'
        print(f'{report.name}: {link}')
        results.append(result)
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--credenciales', type=Path, default=Path('credentials.json'))
    parser.add_argument('--carpeta', required=True, help='ID o URL de la carpeta de revisión')
    parser.add_argument('--subcarpeta', default='Informes de NC')
    parser.add_argument('--token', type=Path,
                        default=Path.home() / '.informes-nc-retie' / 'token.json')
    parser.add_argument('informes', nargs='+', type=Path)
    args = parser.parse_args()
    try:
        parent = folder_id(args.carpeta)
        validate_inputs(args.credenciales, args.informes)
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        from googleapiclient.errors import HttpError
        credentials = authenticate(args.credenciales, args.token)
        service = build('drive', 'v3', credentials=credentials, cache_discovery=False)
        upload_reports(service, parent, args.informes, MediaFileUpload, args.subcarpeta)
    except (ValueError, OSError) as error:
        print(f'No se completó la carga: {error}', file=sys.stderr)
        return 1
    except ImportError:
        print('Faltan dependencias. Ejecuta: py -m pip install -r requirements.txt', file=sys.stderr)
        return 1
    except HttpError as error:
        print(f'Google Drive rechazó la operación (HTTP {error.resp.status}). '
              'Comprueba que la API está habilitada y la cuenta puede editar la carpeta. '
              'Los enlaces ya mostrados corresponden a archivos cargados.', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
