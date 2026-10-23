"""Reemplaza el contenido de un archivo existente de Google Drive por un notebook local.

El archivo conserva su ID, por lo que el link de Colab sigue siendo el mismo.

Uso:
    python scripts/subir_notebook_drive.py <file_id> <notebook.ipynb> [modifiedTime esperado]

Si se indica el modifiedTime esperado y el archivo de Drive cambió desde entonces, no se sobrescribe.

Requiere la llave JSON de una service account de Google Cloud con la Drive API habilitada, en la variable de
entorno GDRIVE_SA_KEY (el JSON o el JSON en base64) o en un archivo indicado por GDRIVE_SA_KEY_FILE, y el
archivo de Drive compartido con el correo de esa cuenta como Editor. Cierra la pestaña de Colab del notebook
antes de subirlo, para que su autoguardado no lo pise.
"""
import base64
import json
import os
import sys

import requests
from google.auth.transport.requests import Request
from google.oauth2 import service_account

API = 'https://www.googleapis.com/drive/v3/files'
UPLOAD_API = 'https://www.googleapis.com/upload/drive/v3/files'


def load_key():
    key_file = os.environ.get('GDRIVE_SA_KEY_FILE')
    if key_file:
        return json.load(open(key_file))
    raw = os.environ.get('GDRIVE_SA_KEY', '').strip()
    if not raw:
        sys.exit('Falta la variable de entorno GDRIVE_SA_KEY o GDRIVE_SA_KEY_FILE.')
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return json.loads(base64.b64decode(raw))


def main(file_id, path, expected_modified=None):
    content = open(path, 'rb').read()
    json.loads(content)  # el notebook debe ser JSON válido

    info = load_key()
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive'])
    creds.refresh(Request())
    headers = {'Authorization': f'Bearer {creds.token}'}

    meta = requests.get(f'{API}/{file_id}', headers=headers, timeout=60, params={
        'fields': 'id,name,mimeType,modifiedTime,capabilities/canEdit', 'supportsAllDrives': 'true'})
    if meta.status_code == 404:
        sys.exit(f'La service account no ve el archivo: compártelo con {info["client_email"]} como Editor.')
    meta.raise_for_status()
    meta = meta.json()
    print(f'Archivo: {meta["name"]} ({meta["mimeType"]}), modificado {meta["modifiedTime"]}')
    if not meta['capabilities']['canEdit']:
        sys.exit(f'Sin permiso de edición: comparte el archivo con {info["client_email"]} como Editor.')
    if expected_modified and meta['modifiedTime'] != expected_modified:
        sys.exit(f'El archivo cambió desde {expected_modified}; revísalo antes de sobrescribirlo.')

    response = requests.patch(
        f'{UPLOAD_API}/{file_id}', data=content, timeout=300,
        headers={**headers, 'Content-Type': meta['mimeType']},
        params={'uploadType': 'media', 'supportsAllDrives': 'true', 'fields': 'id,name,modifiedTime,size'})
    response.raise_for_status()
    print(f'Subido: {response.json()}')


if __name__ == '__main__':
    main(*sys.argv[1:4])
