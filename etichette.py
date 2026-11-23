"""Finestra Python per creare elenchi ed etichette a partire da dati locali."""
import copy
import json
import os
import queue
import re
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont

import etichette_core as core


def _atomic_write(path, write):
    descriptor, temporary = tempfile.mkstemp(prefix='.etichette_', suffix=Path(path).suffix, dir=str(Path(path).resolve().parent))
    os.close(descriptor)
    try:
        write(temporary)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


class FinestraEtichette:
    def __init__(self, parent, header, data, filename):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.withdraw()
        self.window.title('Elenco ed etichette')
        self.window.transient(parent)
        self.window.geometry('1160x760')
        self.window.minsize(900, 630)
        self.window.protocol('WM_DELETE_WINDOW', self.close)
        self.header = []
        self.data = []
        self.fields = []
        self.tables = []
        self.page = 0
        self.visible = []
        self.last_pdf = None
        self.pending = None
        self.updating = True
        self.font_cache = {}
        self.search = tk.StringVar()
        self.sort = tk.StringVar(value='Ordine originale')
        self.reverse = tk.BooleanVar(value=False)
        self.town = tk.StringVar(value='Verolanuova')
        self.include_town = tk.BooleanVar(value=True)
        self.title = tk.StringVar(value='Elenco anagrafico')
        self.preset = tk.StringVar(value=core.POSTAL)
        self.paper = tk.StringVar(value=next(iter(core.PAPER_PRESETS)))
        self.sheet = tk.StringVar()
        self.header_row = tk.StringVar(value='1')
        self.source = tk.StringVar()
        self.info = tk.StringVar()
        self.preview_info = tk.StringVar()
        self.page_info = tk.StringVar()
        self.token_choice = tk.StringVar()
        self.guides = tk.BooleanVar(value=True)
        defaults = dict(next(iter(core.PAPER_PRESETS.values())))
        defaults.update(font=9, padding=4, firstPad=10, skip=0, copies=1)
        self.geometry = {key: tk.StringVar(value=str(value)) for key, value in defaults.items()}
        self._build()
        self.set_initial(header, data, filename)
        self.updating = False
        for variable in (self.search, self.town, self.include_town, self.reverse, self.guides, *self.geometry.values()):
            variable.trace_add('write', self._changed)
        self.window.update_idletasks()
        x = max(0, parent.winfo_rootx() + (parent.winfo_width() - self.window.winfo_width()) // 2)
        y = max(0, parent.winfo_rooty() + (parent.winfo_height() - self.window.winfo_height()) // 2)
        self.window.geometry(f'+{x}+{y}')
        self.window.deiconify()
        self.window.grab_set()
        self.window.focus_set()
        self.render()

    def close(self):
        if self.pending:
            self.window.after_cancel(self.pending)
        self.window.destroy()

    def _build(self):
        outer = ttk.Frame(self.window, padding=12)
        outer.pack(fill='both', expand=True)
        top = ttk.Frame(outer)
        top.pack(fill='x')
        ttk.Label(top, text='Elenco ed etichette', font=('Segoe UI', 17, 'bold')).pack(side='left')
        ttk.Button(top, text='Chiudi', command=self.close).pack(side='right')
        ttk.Label(outer, text='Le etichette usano i filtri di mese e anno di nascita dei biglietti. La ricerca della finestra principale non esclude persone.', wraplength=1100).pack(anchor='w', pady=(2, 6))
        toolbar = ttk.Frame(outer)
        toolbar.pack(fill='x', pady=(0, 6))
        for label, command in (('PDF etichette', self.save_labels), ('PDF elenco', self.save_list),
                               ('Esporta CSV', self.save_csv), ('Salva progetto', self.save_project),
                               ('Apri ultimo PDF', self.open_pdf)):
            ttk.Button(toolbar, text=label, command=command).pack(side='left', padx=(0, 6))
        ttk.Label(outer, textvariable=self.source, wraplength=1100).pack(anchor='w')

        panes = ttk.Panedwindow(outer, orient='horizontal')
        panes.pack(fill='both', expand=True, pady=(5, 0))
        self.tabs = ttk.Notebook(panes, width=340)
        panes.add(self.tabs, weight=0)
        right = ttk.Frame(panes, padding=(10, 0, 0, 0))
        panes.add(right, weight=1)
        self._build_fields()
        self._build_template()
        self._build_paper()
        self._build_input()

        controls = ttk.Frame(right)
        controls.pack(fill='x')
        ttk.Label(controls, text='Cerca in tutti i campi:').grid(row=0, column=0, sticky='w')
        ttk.Entry(controls, textvariable=self.search, width=22).grid(row=0, column=1, sticky='ew', padx=6)
        ttk.Button(controls, text='Cancella', command=lambda: self.search.set('')).grid(row=0, column=2, padx=(0, 8))
        ttk.Label(controls, text='Ordina per:').grid(row=0, column=3, sticky='w')
        self.sort_box = ttk.Combobox(controls, textvariable=self.sort, state='readonly', width=19)
        self.sort_box.grid(row=0, column=4, sticky='ew', padx=6)
        self.sort_box.bind('<<ComboboxSelected>>', self._changed)
        ttk.Checkbutton(controls, text='Decrescente', variable=self.reverse).grid(row=0, column=5)
        controls.columnconfigure(1, weight=1)
        controls.columnconfigure(4, weight=1)
        ttk.Label(right, textvariable=self.info, wraplength=750).pack(anchor='w', pady=(5, 3))
        ttk.Label(right, text='Doppio clic su una cella per modificarla nella finestra delle etichette.').pack(anchor='w')
        table_area = ttk.Frame(right)
        table_area.pack(fill='both', expand=True, pady=(3, 5))
        self.table = ttk.Treeview(table_area, show='headings', height=6)
        table_y = ttk.Scrollbar(table_area, orient='vertical', command=self.table.yview)
        table_x = ttk.Scrollbar(table_area, orient='horizontal', command=self.table.xview)
        self.table.configure(yscrollcommand=table_y.set, xscrollcommand=table_x.set)
        table_y.pack(side='right', fill='y')
        table_x.pack(side='bottom', fill='x')
        self.table.pack(side='left', fill='both', expand=True)
        self.table.bind('<Double-1>', self.edit_cell)
        preview_head = ttk.Frame(right)
        preview_head.pack(fill='x')
        ttk.Label(preview_head, text='Anteprima etichette', font=('Segoe UI', 11, 'bold')).pack(side='left')
        ttk.Button(preview_head, text='◀', width=3, command=lambda: self.change_page(-1)).pack(side='left', padx=(10, 2))
        ttk.Label(preview_head, textvariable=self.page_info).pack(side='left', padx=3)
        ttk.Button(preview_head, text='▶', width=3, command=lambda: self.change_page(1)).pack(side='left')
        self.preview_info_label = ttk.Label(right, textvariable=self.preview_info, foreground='#a23616', wraplength=740)
        self.preview_info_label.pack(anchor='w')
        canvas_area = ttk.Frame(right)
        canvas_area.pack(fill='both', expand=True)
        self.preview = tk.Canvas(canvas_area, background='#e7ebef', highlightthickness=0, height=235)
        preview_y = ttk.Scrollbar(canvas_area, orient='vertical', command=self.preview.yview)
        preview_x = ttk.Scrollbar(canvas_area, orient='horizontal', command=self.preview.xview)
        self.preview.configure(yscrollcommand=preview_y.set, xscrollcommand=preview_x.set)
        preview_y.pack(side='right', fill='y')
        preview_x.pack(side='bottom', fill='x')
        self.preview.pack(fill='both', expand=True)
        self.preview.bind('<MouseWheel>', lambda event: self.preview.yview_scroll(-int(event.delta / 120), 'units'))
        self.preview.bind('<Button-4>', lambda event: self.preview.yview_scroll(-3, 'units'))
        self.preview.bind('<Button-5>', lambda event: self.preview.yview_scroll(3, 'units'))
        ttk.Label(outer, text='Stampa il PDF su A4 al 100% e verifica l’allineamento su carta comune.').pack(anchor='w', pady=(5, 0))

    def _build_fields(self):
        tab = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(tab, text='Campi')
        ttk.Label(tab, text='Campi per elenco ed etichette', font=('Segoe UI', 10, 'bold')).pack(anchor='w')
        ttk.Label(tab, text='Seleziona un campo e usa i pulsanti per includerlo e ordinarlo.', wraplength=300).pack(anchor='w', pady=5)
        fieldbox = ttk.Frame(tab)
        fieldbox.pack(fill='both', expand=True)
        self.field_list = tk.Listbox(fieldbox, exportselection=False, height=10)
        scroll = ttk.Scrollbar(fieldbox, orient='vertical', command=self.field_list.yview)
        self.field_list.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.field_list.pack(fill='both', expand=True)
        self.field_list.bind('<Double-Button-1>', lambda event: self.toggle_field())
        buttons = ttk.Frame(tab)
        buttons.pack(fill='x', pady=5)
        for label, command, width in (('Includi / escludi', self.toggle_field, 15),
                                      ('Rinomina', self.rename_field, 9),
                                      ('↑', lambda: self.move_field(-1), 3),
                                      ('↓', lambda: self.move_field(1), 3)):
            ttk.Button(buttons, text=label, command=command, width=width).pack(side='left', padx=(0, 4))
        ttk.Label(tab, text='Comune fisso di residenza').pack(anchor='w', pady=(8, 0))
        ttk.Entry(tab, textvariable=self.town).pack(fill='x')
        ttk.Checkbutton(tab, text='Aggiungi il Comune fisso all’elenco', variable=self.include_town).pack(anchor='w', pady=5)

    def _build_template(self):
        tab = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(tab, text='Testo')
        ttk.Label(tab, text='Modello per le etichette', font=('Segoe UI', 10, 'bold')).pack(anchor='w')
        choice = ttk.Combobox(tab, textvariable=self.preset, state='readonly', values=(core.POSTAL, core.PERSONAL, core.SELECTED))
        choice.pack(fill='x', pady=5)
        choice.bind('<<ComboboxSelected>>', self.use_preset)
        ttk.Label(tab, text='Testo personalizzabile. La prima riga è in grassetto.', wraplength=300).pack(anchor='w')
        self.template = tk.Text(tab, height=9, wrap='word', undo=True)
        self.template.pack(fill='both', expand=True, pady=5)
        self.template.bind('<<Modified>>', self.template_changed)
        ttk.Label(tab, text='Inserisci un campo nel punto del cursore').pack(anchor='w')
        self.token_box = ttk.Combobox(tab, textvariable=self.token_choice, state='readonly')
        self.token_box.pack(fill='x', pady=4)
        ttk.Button(tab, text='Inserisci campo', command=self.insert_token).pack(anchor='w')
        ttk.Label(tab, text='Puoi anche digitare direttamente i segnaposto, ad esempio {residenza}.', wraplength=300).pack(anchor='w', pady=8)

    def _build_paper(self):
        tab = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(tab, text='Carta')
        ttk.Label(tab, text='Formato del foglio', font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w')
        self.paper_box = ttk.Combobox(tab, textvariable=self.paper, state='readonly', values=(*core.PAPER_PRESETS, core.CUSTOM_PAPER), width=27)
        self.paper_box.grid(row=1, column=0, columnspan=2, sticky='ew', pady=4)
        self.paper_box.bind('<<ComboboxSelected>>', self.paper_changed)
        labels = [('pw','Larghezza pagina (mm)'), ('ph','Altezza pagina (mm)'), ('w','Larghezza etichetta (mm)'),
                  ('h','Altezza etichetta (mm)'), ('left','Margine sinistro (mm)'), ('top','Margine superiore (mm)'),
                  ('dx','Passo orizzontale (mm)'), ('dy','Passo verticale (mm)'), ('cols','Colonne'), ('rows','Righe'),
                  ('font','Dimensione testo (pt)'), ('padding','Spazio interno (mm)'),
                  ('firstPad','Sinistra prima colonna (mm)'), ('skip','Posizioni già usate'), ('copies','Copie per etichetta')]
        for i, (key, label) in enumerate(labels, start=2):
            ttk.Label(tab, text=label).grid(row=i, column=0, sticky='w', pady=1)
            entry = ttk.Entry(tab, textvariable=self.geometry[key], width=8)
            entry.grid(row=i, column=1, sticky='e', pady=1)
            if key in ('pw','ph','w','h','left','top','dx','dy','cols','rows'):
                entry.bind('<KeyRelease>', lambda event: self.paper.set(core.CUSTOM_PAPER))
        ttk.Checkbutton(tab, text='Mostra guide nell’anteprima', variable=self.guides).grid(row=17, column=0, columnspan=2, sticky='w', pady=5)
        tab.columnconfigure(0, weight=1)

    def _build_input(self):
        tab = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(tab, text='File')
        ttk.Label(tab, text='Dati di partenza', font=('Segoe UI', 10, 'bold')).pack(anchor='w')
        ttk.Label(tab, text='I dati arrivano dai biglietti. Puoi anche aprire un altro file o un progetto salvato.', wraplength=290).pack(anchor='w', pady=8)
        ttk.Button(tab, text='Apri tabella o progetto…', command=self.open_file).pack(anchor='w')
        ttk.Label(tab, text='Foglio o tabella').pack(anchor='w', pady=(10, 0))
        self.sheet_box = ttk.Combobox(tab, textvariable=self.sheet, state='readonly')
        self.sheet_box.pack(fill='x')
        self.sheet_box.bind('<<ComboboxSelected>>', self.sheet_changed)
        ttk.Label(tab, text='Riga delle intestazioni').pack(anchor='w', pady=(10, 0))
        ttk.Entry(tab, textvariable=self.header_row, width=8).pack(anchor='w')
        ttk.Button(tab, text='Applica foglio e riga', command=self.apply_sheet).pack(anchor='w', pady=6)
        ttk.Label(tab, text='Titolo del PDF elenco').pack(anchor='w', pady=(10, 0))
        ttk.Entry(tab, textvariable=self.title).pack(fill='x')
        ttk.Label(tab, text='Le modifiche alle celle in questa finestra valgono per le etichette e non cambiano il file Excel né l’elenco dei biglietti.', wraplength=290).pack(anchor='w', pady=12)

    def set_initial(self, header, data, filename):
        rows = [[core.cell_text(value) for value in row] for row in data]
        self.tables = [('Selezione dai biglietti', [list(header)] + copy.deepcopy(rows))]
        self.sheet_box.configure(values=['Selezione dai biglietti'])
        self.sheet.set('Selezione dai biglietti')
        self.header_row.set('1')
        self.adopt_data(header, rows)
        self.source.set(f'{filename or "Elenco"} · {len(rows)} persone dai filtri dei biglietti')

    def adopt_data(self, header, data):
        self.header = [core.cell_text(value) or f'Campo {i+1}' for i, value in enumerate(header)]
        self.data = [[core.cell_text(row[i]) if i < len(row) else '' for i in range(len(self.header))] for row in data]
        self.fields, is_person = core.initial_fields(self.header)
        self.include_town.set(is_person)
        self.preset.set(core.POSTAL if is_person else core.SELECTED)
        self.replace_template(core.default_template(self.header, self.fields, self.preset.get()))
        self.refresh_fields()
        self.page = 0

    def replace_template(self, text):
        self.template.delete('1.0', 'end')
        self.template.insert('1.0', text)
        self.template.edit_modified(False)

    def use_preset(self, event=None):
        self.replace_template(core.default_template(self.header, self.fields, self.preset.get()))
        self._changed()

    def template_changed(self, event=None):
        if self.template.edit_modified():
            self.template.edit_modified(False)
            self._changed()

    def insert_token(self):
        index = self.token_box.current()
        if 0 <= index < len(self.tokens):
            self.template.insert('insert', '{' + self.tokens[index] + '}')
            self.template.focus_set()
            self._changed()

    def refresh_fields(self, selected=None):
        prior_sort = self.sort.get()
        self.field_list.delete(0, 'end')
        for field in self.fields:
            mark = '☑' if field['enabled'] else '☐'
            self.field_list.insert('end', f'{mark}  {field["name"]}')
        if selected is not None and 0 <= selected < len(self.fields):
            self.field_list.selection_set(selected)
            self.field_list.see(selected)
        self.sort_values = [('Ordine originale', None)] + [(f'{field["name"]} [{field["i"]+1}]', field['i']) for field in self.fields]
        self.sort_box.configure(values=[name for name, _ in self.sort_values])
        if prior_sort not in [name for name, _ in self.sort_values]:
            self.sort.set('Ordine originale')
        self.tokens = [f'c{i}' for i in range(len(self.header))] + ['residenza', 'nato']
        token_labels = [f'{name} [{i+1}]' for i, name in enumerate(self.header)] + ['Comune fisso', 'Nato/a']
        self.token_box.configure(values=token_labels)
        if token_labels:
            self.token_box.current(0)

    def sync_fields_to_template(self):
        lines = self.template.get('1.0', 'end-1c').split('\n')
        selected = [field for field in self.fields if field['enabled']]
        enabled = {field['i'] for field in selected}
        fixed = [line for line in lines if not re.search(r'\{c\d+\}', line) and line.strip()]
        ordered = []
        for field in selected:
            token = '{c' + str(field['i']) + '}'
            existing = next((line for line in lines if token in line and all(int(n) in enabled for n in re.findall(r'\{c(\d+)\}', line))), None)
            if existing is None:
                existing = token
                name = self.header[field['i']]
                if self.preset.get() == core.PERSONAL and re.search(r'comune.*nasc|luogo.*nasc', name, re.I):
                    existing = '{nato} a ' + token
                elif self.preset.get() == core.PERSONAL and re.search(r'data.*nasc', name, re.I):
                    existing = 'il ' + token
            if existing not in ordered:
                ordered.append(existing)
        if self.preset.get() == core.POSTAL and '{residenza}' in fixed:
            address = next((i for i, field in enumerate(selected) if re.search(r'indirizzo|^via$', self.header[field['i']], re.I)), len(ordered))
            ordered.insert(min(address, len(ordered)), '{residenza}')
            fixed = [line for line in fixed if line != '{residenza}']
        self.replace_template('\n'.join(ordered + fixed))

    def selected_field_index(self):
        selection = self.field_list.curselection()
        return int(selection[0]) if selection else None

    def toggle_field(self):
        i = self.selected_field_index()
        if i is None:
            return
        self.fields[i]['enabled'] = not self.fields[i]['enabled']
        self.sync_fields_to_template()
        self.refresh_fields(i)
        self._changed()

    def rename_field(self):
        i = self.selected_field_index()
        if i is None:
            return
        name = simpledialog.askstring('Rinomina campo', 'Nome del campo:', initialvalue=self.fields[i]['name'], parent=self.window)
        if name is not None and name.strip():
            self.fields[i]['name'] = name.strip()
            self.refresh_fields(i)
            self._changed()

    def move_field(self, delta):
        i = self.selected_field_index()
        if i is None or not 0 <= i + delta < len(self.fields):
            return
        self.fields[i], self.fields[i + delta] = self.fields[i + delta], self.fields[i]
        self.sync_fields_to_template()
        self.refresh_fields(i + delta)
        self._changed()

    def paper_changed(self, event=None):
        values = core.PAPER_PRESETS.get(self.paper.get())
        if values:
            for key, value in values.items():
                self.geometry[key].set(str(value))
        self.page = 0
        self._changed()

    def current_geometry(self):
        result = {}
        for key, variable in self.geometry.items():
            text = variable.get().strip().replace(',', '.')
            try:
                number = float(text)
            except ValueError as error:
                raise ValueError('Inserisci valori numerici in tutti i campi del formato carta.') from error
            result[key] = number
        for key in ('cols', 'rows', 'skip', 'copies'):
            if not result[key].is_integer():
                raise ValueError('Righe, colonne, posizioni già usate e copie devono essere numeri interi.')
            result[key] = int(result[key])
        return result

    def _changed(self, *args):
        if self.updating or not self.window.winfo_exists():
            return
        if self.pending:
            self.window.after_cancel(self.pending)
        self.pending = self.window.after(120, self.render)

    def filtered_indexes(self):
        query = self.search.get().casefold().strip()
        indexes = [i for i, row in enumerate(self.data) if not query or any(query in value.casefold() for value in row)]
        selected = next((index for label, index in self.sort_values if label == self.sort.get()), None)
        if selected is not None:
            def natural(i):
                value = self.data[i][selected].casefold()
                return [(1, int(part)) if part.isdigit() else (0, part) for part in re.split(r'(\d+)', value)]
            indexes.sort(key=natural, reverse=self.reverse.get())
        return indexes

    def render(self):
        self.pending = None
        if not self.window.winfo_exists():
            return
        self.visible = self.filtered_indexes()
        self.info.set(f'{len(self.visible)} voci selezionate su {len(self.data)}' + (' · tabella limitata alle prime 300, i PDF includono tutte le voci selezionate.' if len(self.visible) > 300 else '.'))
        active = [field for field in self.fields if field['enabled']]
        columns = [f'c{field["i"]}' for field in active]
        if self.include_town.get():
            columns.append('town')
        self.table.delete(*self.table.get_children())
        self.table.configure(columns=columns)
        for field in active:
            column = f'c{field["i"]}'
            self.table.heading(column, text=field['name'])
            self.table.column(column, width=140, minwidth=75)
        if self.include_town.get():
            self.table.heading('town', text='Comune di residenza')
            self.table.column('town', width=150, minwidth=75)
        for index in self.visible[:300]:
            row = self.data[index]
            values = [row[field['i']] for field in active]
            if self.include_town.get():
                values.append(self.town.get())
            self.table.insert('', 'end', iid=str(index), values=values)
        self.draw_preview()

    def draw_preview(self):
        self.preview.delete('all')
        try:
            g = self.current_geometry()
            items = [self.data[i] for i in self.visible]
            core.validate_geometry(g, len(items))
            if not items:
                self.preview_info.set('Nessuna voce da mostrare.')
                self.page_info.set('0 / 0')
                return
            template = self.template.get('1.0', 'end-1c')
            if not template.strip():
                raise ValueError('Il testo delle etichette è vuoto.')
            count = core.page_count(g, len(items))
            self.page = min(max(0, self.page), count - 1)
            labels = list(core.page_layout(g, items, self.header, template, self.town.get(), self.page))
            self.page_info.set(f'{self.page+1} / {count}')
            self.preview_info.set('')
            scale = 510 / g['pw']
            left, top = 12, 12
            width, height = g['pw'] * scale, g['ph'] * scale
            self.preview.configure(scrollregion=(0, 0, width + 24, height + 24))
            self.preview.create_rectangle(left, top, left + width, top + height, fill='white', outline='#a7afb6')
            font_pixels = max(5, round(g['font'] * scale * 25.4 / 72))
            for label in labels:
                x, y = left + label['x'] * scale, top + label['y'] * scale
                if self.guides.get():
                    self.preview.create_rectangle(x, y, x + g['w']*scale, y + g['h']*scale, outline='#aebdcb', dash=(3, 3))
                height_pixels = label['content_height'] * scale * 25.4 / 72
                text_y = y + (g['h']*scale - height_pixels) / 2
                for line, font_name in label['lines']:
                    weight = 'bold' if font_name == 'Etichette-Bold' else 'normal'
                    key = (font_pixels, weight)
                    if key not in self.font_cache:
                        self.font_cache[key] = tkfont.Font(family='DejaVu Sans', size=-font_pixels, weight=weight)
                    self.preview.create_text(x + label['pad_left']*scale, text_y, text=line, anchor='nw', font=self.font_cache[key], fill='#111')
                    text_y += g['font'] * 1.18 * scale * 25.4 / 72
        except (ValueError, OSError) as error:
            self.preview_info.set(str(error))
            self.page_info.set('—')

    def change_page(self, delta):
        self.page = max(0, self.page + delta)
        self.draw_preview()
        self.preview.yview_moveto(0)

    def edit_cell(self, event):
        row_id = self.table.identify_row(event.y)
        column_id = self.table.identify_column(event.x)
        if not row_id or not column_id:
            return
        active = [field for field in self.fields if field['enabled']]
        column = int(column_id[1:]) - 1
        if not 0 <= column < len(active):
            return
        index = int(row_id)
        field = active[column]
        value = simpledialog.askstring('Modifica dato', field['name'] + ':', initialvalue=self.data[index][field['i']], parent=self.window)
        if value is not None:
            self.data[index][field['i']] = value
            self._changed()

    def open_file(self):
        path = filedialog.askopenfilename(parent=self.window, title='Apri tabella o progetto', filetypes=[('Tabelle e progetti', '*.xlsx *.xlsm *.xls *.csv *.tsv *.html *.htm *.json'), ('Tutti i file', '*.*')])
        if not path:
            return
        try:
            if Path(path).suffix.lower() == '.json':
                self.open_project(path)
                return
            tables = core.read_tables(path)
            if not tables:
                raise ValueError('Nessuna tabella trovata.')
            self.tables = tables
            self.sheet_box.configure(values=[name for name, _ in tables])
            index = max(range(len(tables)), key=lambda i: len(tables[i][1]))
            self.sheet_box.current(index)
            self.header_row.set(str(core.detect_header(tables[index][1]) + 1))
            self.apply_sheet()
            self.source.set(f'{Path(path).name} · {len(self.data)} righe')
        except Exception as error:
            messagebox.showerror('Impossibile aprire il file', str(error), parent=self.window)

    def sheet_changed(self, event=None):
        index = self.sheet_box.current()
        if 0 <= index < len(self.tables):
            self.header_row.set(str(core.detect_header(self.tables[index][1]) + 1))
            self.apply_sheet()

    def apply_sheet(self):
        try:
            index = self.sheet_box.current()
            if not 0 <= index < len(self.tables):
                raise ValueError('Scegli un foglio o una tabella.')
            header_index = int(self.header_row.get()) - 1
            header, rows = core.apply_header(self.tables[index][1], header_index)
            self.updating = True
            self.adopt_data(header, rows)
            self.search.set('')
            self.sort.set('Ordine originale')
            self.updating = False
            self.render()
        except (ValueError, IndexError) as error:
            self.updating = False
            messagebox.showerror('Importazione', str(error), parent=self.window)

    def selected_data(self):
        return [self.data[index][:] for index in self.filtered_indexes()]

    def run_pdf_job(self, path, writer, description):
        dialog = tk.Toplevel(self.window)
        dialog.title('Creazione PDF')
        dialog.transient(self.window)
        dialog.resizable(False, False)
        dialog.protocol('WM_DELETE_WINDOW', lambda: None)
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text=description, font=('Segoe UI', 11, 'bold')).pack(anchor='w')
        ttk.Label(frame, text='Attendi il completamento.').pack(anchor='w', pady=6)
        bar = ttk.Progressbar(frame, mode='indeterminate', length=290)
        bar.pack(fill='x')
        bar.start(12)
        dialog.update_idletasks()
        x = max(0, self.window.winfo_rootx() + (self.window.winfo_width() - dialog.winfo_reqwidth()) // 2)
        y = max(0, self.window.winfo_rooty() + (self.window.winfo_height() - dialog.winfo_reqheight()) // 2)
        dialog.geometry(f'+{x}+{y}')
        dialog.grab_set()
        result = queue.Queue()
        def work():
            try:
                _atomic_write(path, writer)
                result.put((True, ''))
            except Exception as error:
                result.put((False, str(error)))
        def poll():
            try:
                ok, message = result.get_nowait()
            except queue.Empty:
                dialog.after(75, poll)
                return
            bar.stop()
            dialog.destroy()
            self.window.grab_set()
            if ok:
                self.last_pdf = path
                messagebox.showinfo('PDF creato', f'File salvato:\n{path}\n\nUsa «Apri ultimo PDF» per aprirlo e stamparlo.', parent=self.window)
            else:
                messagebox.showerror('Impossibile creare il PDF', message, parent=self.window)
        def start():
            threading.Thread(target=work, daemon=True).start()
            poll()
        dialog.after(80, start)

    def save_labels(self):
        try:
            g = self.current_geometry()
            items = self.selected_data()
            core.validate_geometry(g, len(items))
            if not items:
                raise ValueError('Nessuna persona selezionata.')
            template = self.template.get('1.0', 'end-1c')
            if not template.strip():
                raise ValueError('Il testo delle etichette è vuoto.')
            header, town = self.header[:], self.town.get()
            path = filedialog.asksaveasfilename(parent=self.window, defaultextension='.pdf', initialfile='Etichette.pdf', filetypes=[('PDF', '*.pdf')])
            if path:
                self.run_pdf_job(path, lambda temp: core.create_labels_pdf(temp, items, header, template, town, g), 'Creazione PDF etichette')
        except ValueError as error:
            messagebox.showerror('Controlla le etichette', str(error), parent=self.window)

    def save_list(self):
        items = self.selected_data()
        fields = copy.deepcopy(self.fields)
        town, include_town, title = self.town.get(), self.include_town.get(), self.title.get()
        if not items or not (any(field['enabled'] for field in fields) or include_town):
            messagebox.showerror('Controlla l’elenco', 'Seleziona almeno una persona e un campo.', parent=self.window)
            return
        path = filedialog.asksaveasfilename(parent=self.window, defaultextension='.pdf', initialfile='Elenco.pdf', filetypes=[('PDF', '*.pdf')])
        if path:
            self.run_pdf_job(path, lambda temp: core.create_list_pdf(temp, items, fields, town, include_town, title), 'Creazione PDF elenco')

    def save_csv(self):
        items, fields = self.selected_data(), copy.deepcopy(self.fields)
        if not items or not (any(field['enabled'] for field in fields) or self.include_town.get()):
            messagebox.showerror('Controlla l’elenco', 'Seleziona almeno una persona e un campo.', parent=self.window)
            return
        path = filedialog.asksaveasfilename(parent=self.window, defaultextension='.csv', initialfile='Elenco.csv', filetypes=[('CSV', '*.csv')])
        if not path:
            return
        try:
            _atomic_write(path, lambda temp: core.export_csv(temp, items, fields, self.town.get(), self.include_town.get()))
            messagebox.showinfo('CSV creato', f'File salvato:\n{path}', parent=self.window)
        except Exception as error:
            messagebox.showerror('Impossibile creare il CSV', str(error), parent=self.window)

    def project_settings(self):
        settings = {key: variable.get() for key, variable in self.geometry.items()}
        sort_index = next((index for label, index in self.sort_values if label == self.sort.get()), None)
        settings.update(includeTown=self.include_town.get(), town=self.town.get(), template=self.template.get('1.0', 'end-1c'),
                        guides=self.guides.get(), title=self.title.get(), filter=self.search.get(), sort='' if sort_index is None else str(sort_index),
                        reverse=self.reverse.get(), preset={core.POSTAL:'post', core.PERSONAL:'anag', core.SELECTED:'all'}[self.preset.get()],
                        paperPreset=next(('as' if i == 0 else 'u99' for i, name in enumerate(core.PAPER_PRESETS) if name == self.paper.get()), 'custom'))
        return settings

    def save_project(self):
        if not self.data:
            messagebox.showerror('Progetto', 'Non ci sono dati da salvare.', parent=self.window)
            return
        path = filedialog.asksaveasfilename(parent=self.window, defaultextension='.json', initialfile='Progetto_etichette.json', filetypes=[('Progetto JSON', '*.json')])
        if not path:
            return
        project = {'version': 1, 'header': self.header[:], 'data': copy.deepcopy(self.data), 'fields': copy.deepcopy(self.fields), 'settings': self.project_settings()}
        def write(temp):
            with open(temp, 'w', encoding='utf-8') as output:
                json.dump(project, output, ensure_ascii=False, indent=2)
        try:
            _atomic_write(path, write)
            messagebox.showinfo('Progetto salvato', f'File salvato:\n{path}', parent=self.window)
        except Exception as error:
            messagebox.showerror('Impossibile salvare il progetto', str(error), parent=self.window)

    def open_project(self, path):
        if Path(path).stat().st_size > 30_000_000:
            raise ValueError('Il progetto supera 30 MB.')
        with open(path, encoding='utf-8-sig') as source:
            project = json.load(source)
        header, data, fields, settings = (project.get(key) for key in ('header', 'data', 'fields', 'settings'))
        if project.get('version') != 1 or not isinstance(header, list) or not 0 < len(header) <= 200 or not isinstance(data, list) or not isinstance(fields, list) or not isinstance(settings, dict):
            raise ValueError('Progetto non riconosciuto o danneggiato.')
        if any(not isinstance(row, list) for row in data) or any(not isinstance(field, dict) or not isinstance(field.get('i'), int) or not 0 <= field['i'] < len(header) for field in fields):
            raise ValueError('Dati o campi del progetto non validi.')
        if len({field['i'] for field in fields}) != len(fields):
            raise ValueError('Campi duplicati nel progetto.')
        self.updating = True
        try:
            self.tables = []
            self.sheet_box.configure(values=[])
            self.sheet.set('')
            self.header_row.set('1')
            self.adopt_data(header, data)
            self.fields = [{'i': field['i'], 'name': str(field.get('name', self.header[field['i']])), 'enabled': bool(field.get('enabled', True))} for field in fields]
            self.refresh_fields()
            for key, variable in self.geometry.items():
                if key in settings:
                    variable.set(str(settings[key]))
            self.town.set(str(settings.get('town', 'Verolanuova')))
            self.include_town.set(bool(settings.get('includeTown', True)))
            self.title.set(str(settings.get('title', 'Elenco anagrafico')))
            self.search.set(str(settings.get('filter', '')))
            self.reverse.set(bool(settings.get('reverse', False)))
            self.guides.set(bool(settings.get('guides', True)))
            self.preset.set({'post':core.POSTAL, 'anag':core.PERSONAL, 'all':core.SELECTED}.get(settings.get('preset'), core.SELECTED))
            self.paper.set({'as':list(core.PAPER_PRESETS)[0], 'u99':list(core.PAPER_PRESETS)[1]}.get(settings.get('paperPreset'), core.CUSTOM_PAPER))
            self.replace_template(str(settings.get('template', '')))
            sort_index = str(settings.get('sort', ''))
            self.sort.set(next((label for label, index in self.sort_values if index is not None and str(index) == sort_index), 'Ordine originale'))
            self.source.set(f'{Path(path).name} · progetto con {len(self.data)} righe')
            self.page = 0
        finally:
            self.updating = False
        self.render()

    def open_pdf(self):
        if not self.last_pdf or not os.path.isfile(self.last_pdf):
            messagebox.showinfo('Apri PDF', 'Crea prima un PDF in questa finestra.', parent=self.window)
            return
        try:
            if hasattr(os, 'startfile'):
                os.startfile(self.last_pdf)
            else:
                import webbrowser
                webbrowser.open(Path(self.last_pdf).resolve().as_uri())
        except OSError as error:
            messagebox.showerror('Impossibile aprire il PDF', str(error), parent=self.window)


def apri_etichette(parent, header, data, filename):
    FinestraEtichette(parent, header[:], copy.deepcopy(data), filename)
