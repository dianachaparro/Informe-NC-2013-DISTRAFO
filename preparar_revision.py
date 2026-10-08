"""Descarga una carpeta Drive y prepara paquetes de menos de 32 MiB."""
import json
import re
import zipfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

LIMIT = 25 * 1024 * 1024
EXPORTS = {
    'application/vnd.google-apps.document': ('application/pdf', '.pdf'),
    'application/vnd.google-apps.spreadsheet': ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xlsx'),
    'application/vnd.google-apps.presentation': ('application/pdf', '.pdf'),
    'application/vnd.google-apps.drawing': ('application/pdf', '.pdf'),
}


def safe_name(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .')[:120] or 'archivo'


def package_files(folder, manifest):
    """Partir archivos grandes sin alterar bytes; manifest permite reconstruirlos."""
    paths = []
    archive = None
    used = 0
    try:
        for source in sorted((folder / 'archivos').rglob('*')):
            if not source.is_file():
                continue
            relative = source.relative_to(folder).as_posix()
            size = source.stat().st_size
            with source.open('rb') as stream:
                part = 1
                while True:
                    data = stream.read(LIMIT)
                    if not data and part > 1:
                        break
                    entry = relative if size <= LIMIT else relative + f'.parte{part:04d}'
                    budget = len(data) + len(entry.encode('utf-8')) * 2 + 1024
                    if archive is None or used + budget > LIMIT + 4096:
                        if archive is not None:
                            archive.close()
                        target = folder / f'PARA_ADJUNTAR_{len(paths)+1:03d}.zip'
                        paths.append(target)
                        archive = zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED)
                        used = 0
                    archive.writestr(entry, data)
                    used += budget
                    if size > LIMIT:
                        manifest['partes'].append({'archivo': relative, 'parte': entry, 'orden': part})
                    part += 1
                    if not data or stream.tell() == size:
                        break
    finally:
        if archive is not None:
            archive.close()
    if any(p.stat().st_size >= 32 * 1024 * 1024 for p in paths):
        raise ValueError('Un paquete excedió el tamaño permitido.')
    (folder / 'inventario_revision.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    return paths


def download_folder(service, folder_id, destination, manifest, visited=None):
    from googleapiclient.http import MediaIoBaseDownload
    visited = set() if visited is None else visited
    if folder_id in visited:
        raise ValueError('Se encontró una carpeta repetida en el recorrido.')
    visited.add(folder_id)
    token = None
    while True:
        response = service.files().list(q=f"'{folder_id}' in parents and trashed = false", fields='nextPageToken,files(id,name,mimeType)', pageSize=1000, pageToken=token, supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        for item in response.get('files', []):
            target = destination / (safe_name(item['name']) + '__' + item['id'][-8:])
            mime = item['mimeType']
            if mime == 'application/vnd.google-apps.folder':
                target.mkdir(parents=True, exist_ok=True)
                download_folder(service, item['id'], target, manifest, visited)
                continue
            if mime.startswith('application/vnd.google-apps.'):
                if mime not in EXPORTS:
                    manifest['omitidos'].append(item)
                    continue
                export_mime, suffix = EXPORTS[mime]
                target = target.with_name(target.name + suffix)
                request = service.files().export_media(fileId=item['id'], mimeType=export_mime)
            else:
                # Mantener extensión para abrir PDF, fotos, planos y archivos comprimidos.
                suffix = Path(item['name']).suffix
                if suffix:
                    target = target.with_name(target.name + suffix)
                request = service.files().get_media(fileId=item['id'], supportsAllDrives=True)
            print('Descargando:', item['name'], flush=True)
            with target.open('wb') as stream:
                downloader = MediaIoBaseDownload(stream, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            manifest['archivos'].append(dict(item, local=str(target.relative_to(manifest['_root']))))
        token = response.get('nextPageToken')
        if not token:
            break


def main():
    from subir_drive import authenticate, folder_id
    from googleapiclient.discovery import build
    root = tk.Tk()
    root.withdraw()
    output = None
    try:
        url = simpledialog.askstring('Preparar revisión', 'Pega el enlace de la carpeta de la obra:', parent=root)
        if not url:
            return
        identifier = folder_id(url)
        base = filedialog.askdirectory(title='Elige dónde guardar la descarga', parent=root)
        if not base:
            return
        output = Path(base) / ('Revision_' + datetime.now().strftime('%Y%m%d_%H%M%S'))
        output.mkdir(exist_ok=False)
        client = Path(__file__).parent / 'credentials.json'
        token = Path.home() / '.informes-nc-retie' / 'token.json'
        if not token.exists() and not client.exists():
            selected = filedialog.askopenfilename(title='Selecciona el JSON de credenciales existente', filetypes=[('JSON', '*.json')], parent=root)
            if not selected:
                return
            client = Path(selected)
        credentials = authenticate(client, token)
        service = build('drive', 'v3', credentials=credentials)
        manifest = dict(carpeta=url, archivos=[], omitidos=[], partes=[], _root=str(output))
        destination = output / 'archivos'
        destination.mkdir()
        download_folder(service, identifier, destination, manifest)
        del manifest['_root']
        packages = package_files(output, manifest)
        messagebox.showinfo('Preparación terminada', f'Se prepararon {len(packages)} ZIP.\nAdjunta aquí inventario_revision.json y los ZIP PARA_ADJUNTAR.\n\nCarpeta: {output}\nOmitidos: {len(manifest["omitidos"])}', parent=root)
        import os
        if hasattr(os, 'startfile'):
            os.startfile(output)
    except Exception as error:
        messagebox.showerror('No se completó la preparación', f'{error}\n\nLas descargas parciales se conservan en: {output or "sin carpeta"}. No se modificó Drive.', parent=root)
    finally:
        root.destroy()


if __name__ == '__main__':
    main()
