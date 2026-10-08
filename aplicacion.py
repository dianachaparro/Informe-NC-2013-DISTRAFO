"""Aplicación de escritorio para proyectos de revisión RETIE."""
import json
from pathlib import Path
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from proyectos import SCOPES, RESULTS, new_project, import_checklist, save_project, load_project, validate_review, export_project
from subir_drive import folder_id, authenticate, upload_reports


class Application:
    def __init__(self, root):
        self.root = root
        root.title('Informes de no conformidades RETIE')
        root.geometry('1150x750')
        self.project = new_project()
        self.filename = None
        self.outputs = []
        self.dirty = False
        self.busy = False
        self.values = {}
        toolbar = ttk.Frame(root, padding=8)
        toolbar.pack(fill='x')
        for label, callback in [('Nuevo proyecto', self.new), ('Abrir proyecto', self.open), ('Guardar proyecto', self.save),
                                ('Generar Excel', self.export), ('Subir a Drive', self.upload)]:
            ttk.Button(toolbar, text=label, command=lambda f=callback: self.guard(f)).pack(side='left', padx=4)
        details = ttk.LabelFrame(root, text='Datos del proyecto', padding=8)
        details.pack(fill='x', padx=10)
        labels = [('project', 'Nombre de obra'), ('inspection', 'Código de inspección'), ('address', 'Dirección'),
                  ('regulation', 'Reglamento'), ('transition', 'Soportes de transición / observaciones'),
                  ('drive_folder', 'Enlace de carpeta Drive'), ('template', 'Formato F-GI-51')]
        for i, (key, label) in enumerate(labels):
            ttk.Label(details, text=label).grid(row=i, column=0, sticky='w', padx=4, pady=2)
            variable = tk.StringVar()
            self.values[key] = variable
            if key == 'regulation':
                widget = ttk.Combobox(details, textvariable=variable, values=['RETIE 2013', 'RETIE 2024', 'Por confirmar'], state='readonly')
            else:
                widget = ttk.Entry(details, textvariable=variable)
            widget.grid(row=i, column=1, sticky='ew', padx=4)
            variable.trace_add('write', lambda *_: self.changed())
        ttk.Button(details, text='Elegir formato', command=lambda: self.guard(self.template)).grid(row=6, column=2)
        details.columnconfigure(1, weight=1)
        scope_bar = ttk.LabelFrame(root, text='Alcances y listas de verificación', padding=8)
        scope_bar.pack(fill='x', padx=10, pady=8)
        self.scopes = {}
        for i, scope in enumerate(SCOPES):
            variable = tk.BooleanVar()
            self.scopes[scope] = variable
            ttk.Checkbutton(scope_bar, text=scope, variable=variable, command=self.scope_changed).grid(row=0, column=i, padx=12)
            ttk.Button(scope_bar, text='Importar lista', command=lambda s=scope: self.guard(lambda: self.import_list(s))).grid(row=1, column=i, pady=5)
        ttk.Label(scope_bar, text='Confirma que la lista y sus numerales corresponden a la versión RETIE elegida.').grid(row=2, column=0, columnspan=3, sticky='w')
        table_frame = ttk.Frame(root)
        table_frame.pack(fill='both', expand=True, padx=10)
        self.table = ttk.Treeview(table_frame, columns=('scope', 'item', 'activity', 'result'), show='headings', selectmode='browse')
        for col, title, width in [('scope', 'Alcance', 160), ('item', 'Ítem', 55), ('activity', 'Requisito', 630), ('result', 'Resultado', 170)]:
            self.table.heading(col, text=title)
            self.table.column(col, width=width, stretch=col == 'activity')
        scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.table.yview)
        self.table.configure(yscrollcommand=scrollbar.set)
        self.table.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        self.table.bind('<Double-1>', lambda _: self.guard(self.edit))
        actions = ttk.Frame(root, padding=8)
        actions.pack(fill='x')
        ttk.Button(actions, text='Revisar requisito / registrar NC', command=lambda: self.guard(self.edit)).pack(side='left', padx=5)
        ttk.Button(actions, text='Añadir requisito manual', command=lambda: self.guard(self.manual)).pack(side='left', padx=5)
        self.status = tk.StringVar()
        ttk.Label(root, textvariable=self.status, padding=8).pack(fill='x')
        root.protocol('WM_DELETE_WINDOW', lambda: self.guard(self.close))
        self.populate()

    def guard(self, callback):
        if self.busy:
            return
        try:
            callback()
        except Exception as error:
            messagebox.showerror('No se completó', str(error), parent=self.root)

    def changed(self):
        if 'regulation' in self.values and self.values['regulation'].get() != self.project['regulation']:
            self.project['regulation'] = self.values['regulation'].get()
            for row in self.project['requirements']:
                row['reference_validated'] = False
        self.dirty = True
        self.outputs = []

    def sync(self):
        self.project.update({key: value.get().strip() for key, value in self.values.items()})
        self.project['scopes'] = [s for s, value in self.scopes.items() if value.get()]

    def populate(self):
        for key, variable in self.values.items():
            variable.set(self.project[key])
        for scope, variable in self.scopes.items():
            variable.set(scope in self.project['scopes'])
        self.dirty = False
        self.outputs = []
        self.refresh()

    def refresh(self):
        self.table.delete(*self.table.get_children())
        selected = [s for s, value in self.scopes.items() if value.get()]
        for index, row in enumerate(self.project['requirements']):
            if row['scope'] in selected:
                self.table.insert('', 'end', iid=str(index), values=(row['scope'], row['item'], row['activity'], row['result']))
        count = sum(r['result'] == 'No cumple' and r['scope'] in selected for r in self.project['requirements'])
        self.status.set(f'{len(self.table.get_children())} requisitos visibles · {count} NC registradas · Informes preliminares pendientes de validación final')

    def scope_changed(self):
        self.changed()
        self.refresh()

    def discard(self):
        return not self.dirty or messagebox.askyesno('Cambios sin guardar', '¿Descartar los cambios sin guardar?', parent=self.root)

    def new(self):
        if self.discard():
            self.project = new_project()
            self.filename = None
            self.populate()

    def open(self):
        if not self.discard():
            return
        filename = filedialog.askopenfilename(title='Abrir proyecto', filetypes=[('Proyecto RETIE', '*.json')])
        if filename:
            project = load_project(filename)
            self.project, self.filename = project, Path(filename)
            self.populate()

    def save(self):
        self.sync()
        if self.filename is None:
            path = filedialog.asksaveasfilename(title='Guardar proyecto', defaultextension='.retie.json', initialfile='proyecto.retie.json', filetypes=[('Proyecto RETIE', '*.retie.json')])
            if not path:
                return False
            target = Path(path)
        else:
            target = self.filename
        save_project(self.project, target)
        self.filename = target
        self.project = load_project(self.filename)
        self.dirty = False
        self.status.set('Proyecto guardado. Las evidencias se conservaron en su carpeta propia.')
        return True

    def template(self):
        path = filedialog.askopenfilename(title='Seleccionar formato F-GI-51 con hoja Constructor', filetypes=[('Excel', '*.xlsx')])
        if path:
            self.values['template'].set(path)

    def import_list(self, scope):
        path = filedialog.askopenfilename(title='Lista aplicable a ' + scope, filetypes=[('Excel', '*.xlsx')])
        if path:
            count = import_checklist(self.project, scope, path)
            self.scopes[scope].set(True)
            self.changed()
            self.refresh()
            self.status.set(f'{count} requisitos importados. Revisa cada resultado y numeral.')

    def manual(self):
        chosen = [s for s, var in self.scopes.items() if var.get()]
        if not chosen:
            raise ValueError('Selecciona primero el alcance al que pertenece el requisito.')
        row = dict(key='manual-' + str(len(self.project['requirements'])), scope=chosen[0], item='Manual',
                   activity='Requisito documental o técnico adicional', reference='Pendiente de validar',
                   list_name='Registro manual', result='Pendiente', observation='', documents='', photos=[], reference_validated=False)
        self.edit(row)

    def edit(self, manual=None):
        if manual is None:
            selection = self.table.selection()
            if not selection:
                raise ValueError('Selecciona un requisito en la tabla.')
            row = self.project['requirements'][int(selection[0])]
        else:
            row = manual
        dialog = tk.Toplevel(self.root)
        dialog.title('Revisión del requisito')
        dialog.geometry('850x650')
        dialog.transient(self.root)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text=row['activity'], wraplength=790).pack(fill='x', pady=5)
        scope = tk.StringVar(value=row['scope'])
        ttk.Combobox(frame, textvariable=scope, values=[s for s, v in self.scopes.items() if v.get()], state='readonly').pack(fill='x')
        outcome = tk.StringVar(value=row['result'])
        ttk.Label(frame, text='Resultado de la revisión').pack(anchor='w')
        ttk.Combobox(frame, textvariable=outcome, values=RESULTS, state='readonly').pack(fill='x')
        ttk.Label(frame, text='Numeral aplicable (verificar versión RETIE)').pack(anchor='w', pady=(8, 0))
        reference = tk.StringVar(value=row['reference'])
        ttk.Entry(frame, textvariable=reference).pack(fill='x')
        confirmed = tk.BooleanVar(value=row['reference_validated'])
        ttk.Checkbutton(frame, text='Verifiqué el numeral y su aplicabilidad', variable=confirmed).pack(anchor='w')
        reference.trace_add('write', lambda *_: confirmed.set(False))
        ttk.Label(frame, text='Descripción NC: comenzar por «No se evidencia…». Otros resultados: comentario de revisión.').pack(anchor='w', pady=(8, 0))
        observation = tk.Text(frame, height=6, wrap='word')
        observation.pack(fill='x')
        observation.insert('1.0', row['observation'] or ('No se evidencia ' if row['result'] == 'No cumple' else ''))
        ttk.Label(frame, text='Documento, página y enlace de evidencia (si corresponde)').pack(anchor='w', pady=(8, 0))
        documents = tk.StringVar(value=row['documents'])
        ttk.Entry(frame, textvariable=documents).pack(fill='x')
        photos = list(row['photos'])
        photo_list = tk.Listbox(frame, height=4)
        photo_list.pack(fill='x', pady=5)
        def refresh_photos():
            photo_list.delete(0, 'end')
            for path in photos:
                photo_list.insert('end', Path(path).name)
        def add_photo():
            path = filedialog.askopenfilename(parent=dialog, title='Relacionar foto', filetypes=[('Fotografías', '*.jpg *.jpeg *.png *.webp')])
            if path and path not in photos:
                photos.append(path)
                refresh_photos()
        def remove_photo():
            if photo_list.curselection():
                photos.pop(photo_list.curselection()[0])
                refresh_photos()
        def commit():
            try:
                proposed = dict(row, scope=scope.get(), result=outcome.get(), reference=reference.get().strip(),
                                reference_validated=confirmed.get(), observation=observation.get('1.0', 'end').strip(), documents=documents.get().strip(), photos=photos)
                validate_review(proposed)
                row.update(proposed)
                if manual is not None:
                    self.project['requirements'].append(row)
                self.changed()
                self.refresh()
                dialog.destroy()
            except Exception as error:
                messagebox.showerror('Revisar datos', str(error), parent=dialog)
        refresh_photos()
        buttons = ttk.Frame(frame)
        buttons.pack(fill='x', pady=6)
        for label, action in [('Añadir foto', add_photo), ('Quitar foto', remove_photo), ('Guardar revisión', commit), ('Cancelar', dialog.destroy)]:
            ttk.Button(buttons, text=label, command=action).pack(side='left', padx=4)

    def export(self):
        if not self.save():
            return
        outputs = export_project(self.project, self.filename.parent / (self.filename.stem + '_resultados'))
        self.outputs = outputs
        self.status.set('Excel generados en ' + str(outputs[0].parent))
        messagebox.showinfo('Borradores generados', '\n'.join(str(p) for p in outputs), parent=self.root)

    def upload(self):
        self.sync()
        if not self.outputs:
            raise ValueError('Genera primero los Excel. Los cambios posteriores requieren volver a generarlos.')
        parent = folder_id(self.project['drive_folder'])
        token = Path.home() / '.informes-nc-retie' / 'token.json'
        credentials = None
        if not token.exists():
            path = filedialog.askopenfilename(title='Seleccionar credentials.json (cliente OAuth de escritorio)', filetypes=[('JSON', '*.json')])
            if not path:
                return
            credentials = Path(path)
        elif (Path(__file__).parent / 'credentials.json').exists():
            credentials = Path(__file__).parent / 'credentials.json'
        else:
            # Si el token caducó y no se puede renovar se pedirá el cliente antes de conectar.
            from google.oauth2.credentials import Credentials
            from subir_drive import SCOPES as DRIVE_SCOPES
            current = Credentials.from_authorized_user_file(str(token), DRIVE_SCOPES)
            if not current.valid and not current.refresh_token:
                path = filedialog.askopenfilename(title='Seleccionar credentials.json', filetypes=[('JSON', '*.json')])
                if not path:
                    return
                credentials = Path(path)
        if not messagebox.askyesno('Guardar en Drive', '¿Subir estos dos borradores a Informes de NC dentro de la carpeta indicada? No se reemplazarán archivos anteriores.', parent=self.root):
            return
        self.busy = True
        self.status.set('Conectando y subiendo a Drive…')
        outputs = list(self.outputs)
        # Cola para que solo el hilo principal interactúe con Tk.
        def worker():
            try:
                from googleapiclient.discovery import build
                from googleapiclient.http import MediaFileUpload
                auth = authenticate(credentials or Path('__cliente_no_seleccionado__'), token)
                service = build('drive', 'v3', credentials=auth, cache_discovery=False)
                results = upload_reports(service, parent, outputs, MediaFileUpload)
                self.queue.put((True, '\n'.join(r.get('webViewLink', '') for r in results)))
            except Exception:
                self.queue.put((False, 'No se completó la carga. Revisa cuenta, autorización, API y permiso de edición. Los archivos locales se conservaron; consulta Drive antes de repetir una carga parcial.'))
        import queue
        self.queue = queue.Queue()
        threading.Thread(target=worker, daemon=True).start()
        self.root.after(150, self.poll)

    def poll(self):
        import queue
        try:
            success, message = self.queue.get_nowait()
        except queue.Empty:
            self.root.after(150, self.poll)
            return
        self.busy = False
        self.status.set('Carga completada.' if success else 'Carga pendiente.')
        (messagebox.showinfo if success else messagebox.showerror)('Google Drive', message, parent=self.root)

    def close(self):
        if self.discard():
            self.root.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    Application(root)
    root.mainloop()
