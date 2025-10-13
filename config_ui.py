import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pyautogui
import os

class ToolTip:
    """
    Create a tooltip for a given widget.
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")

        label = tk.Label(self.tooltip_window, text=self.text, justify='left',
                         background="#ffffe0", relief='solid', borderwidth=1,
                         wraplength=400, font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

class CaptureHelper:
    """A helper class to create a temporary, fullscreen, transparent window
    for capturing mouse coordinates with a simple click."""
    def __init__(self, parent, on_capture_callback, capture_mode='point'):
        self.parent = parent
        self.on_capture_callback = on_capture_callback
        self.capture_mode = capture_mode

        self.capture_window = tk.Toplevel(parent)
        self.capture_window.attributes('-alpha', 0.3)
        self.capture_window.attributes('-fullscreen', True)
        self.capture_window.configure(cursor="crosshair")

        self.canvas = tk.Canvas(self.capture_window, bg='grey', highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None

        self.parent.winfo_toplevel().withdraw()

        if self.capture_mode == 'region':
            self.setup_region_capture()
        else:
            self.setup_point_capture()

        self.capture_window.focus_force()

    def setup_point_capture(self):
        label = tk.Label(self.canvas, text="FAI CLICK per catturare la posizione",
                         font=("Arial", 22, "bold"), bg="white", fg="blue", relief="solid", borderwidth=2)
        self.canvas.create_window(self.capture_window.winfo_screenwidth() / 2, self.capture_window.winfo_screenheight() / 2, window=label)
        self.capture_window.bind("<Button-1>", self.capture_point_coords)

    def setup_region_capture(self):
        label = tk.Label(self.canvas, text="TIENI PREMUTO e TRASCINA per selezionare una regione",
                         font=("Arial", 22, "bold"), bg="white", fg="blue", relief="solid", borderwidth=2)
        self.canvas.create_window(self.capture_window.winfo_screenwidth() / 2, self.capture_window.winfo_screenheight() / 2, window=label)
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def capture_point_coords(self, event=None):
        x, y = pyautogui.position()
        self.close_window()
        self.on_capture_callback((x, y))

    def on_button_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

    def on_mouse_drag(self, event):
        cur_x, cur_y = (event.x, event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_button_release(self, event):
        end_x, end_y = (event.x, event.y)

        left = min(self.start_x, end_x)
        top = min(self.start_y, end_y)
        width = abs(self.start_x - end_x)
        height = abs(self.start_y - end_y)

        self.close_window()
        self.on_capture_callback([left, top, width, height])

    def close_window(self):
        self.capture_window.destroy()
        self.parent.winfo_toplevel().deiconify()

class ConfigFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        self.config_data = self.load_config('config.json')
        self.tooltips_data = self.load_config('tooltips.json', is_tooltip=True)

        if not self.config_data:
            self.create_error_label()
            return

        self.vars = {}
        self.create_widgets()

    def load_config(self, filename, is_tooltip=False):
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            if not is_tooltip:
                messagebox.showerror("Errore di Caricamento", f"Impossibile caricare '{filename}':\n{e}\n\nL'applicazione potrebbe non funzionare correttamente.")
            else:
                print(f"Avviso: file '{filename}' non trovato. I tooltip non saranno disponibili.")
            return None

    def create_error_label(self):
        ttk.Label(self, text="Errore nel caricamento di config.json. Impossibile mostrare le impostazioni.",
                  font=("Arial", 14, "bold"), foreground="red").pack(pady=50, padx=20)

    def create_widgets(self):
        # Il frame principale ora usa grid per un miglior controllo
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_notebook = ttk.Notebook(self)
        main_notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 0))

        # --- Main Tab 1: Profili ---
        profiles_tab = ttk.Frame(main_notebook)
        main_notebook.add(profiles_tab, text="Profili")
        self.create_profile_management_tab(profiles_tab)

        # --- Main Tab 2: Mappature (con sub-tabs) ---
        file_mappature_tab = ttk.Frame(main_notebook)
        main_notebook.add(file_mappature_tab, text="Mappature e ODC")
        file_mappature_notebook = ttk.Notebook(file_mappature_tab)
        file_mappature_notebook.pack(pady=5, padx=5, expand=True, fill="both")

        self.mapping_tab_frame = ttk.Frame(file_mappature_notebook)
        self.odc_tab_frame = ttk.Frame(file_mappature_notebook)
        self.mapping_vars = {}

        file_mappature_notebook.add(self.mapping_tab_frame, text="Mappatura Colonne per Profilo")
        self.draw_mapping_entries()

        file_mappature_notebook.add(self.odc_tab_frame, text="Impostazioni ODC per Profilo")
        self.draw_odc_settings()

        # --- Main Tab 3: Impostazioni Globali (con sub-tabs) ---
        global_settings_tab = ttk.Frame(main_notebook)
        main_notebook.add(global_settings_tab, text="Impostazioni Globali")
        global_settings_notebook = ttk.Notebook(global_settings_tab)
        global_settings_notebook.pack(pady=5, padx=5, expand=True, fill="both")

        self.create_generic_tab(global_settings_notebook, ("file_e_fogli_excel", "impostazioni_file"), "Percorsi File")
        self.create_generic_tab(global_settings_notebook, ("file_e_fogli_excel", "mappature_colonne_foglio_avanzamento"), "Colonne Avanzamento")
        self.create_generic_tab(global_settings_notebook, ("coordinate_e_dati", "gui"), "Coordinate GUI")
        self.create_generic_tab(global_settings_notebook, ("timing_e_ritardi",), "Timing")
        self.create_generic_tab(global_settings_notebook, ("pulizia_appunti",), "Pulizia Appunti")
        self.create_generic_tab(global_settings_notebook, ("tasti_rapidi",), "Tasti Rapidi")

        # Frame per il pulsante Salva, posizionato sotto il notebook
        button_frame = ttk.Frame(self)
        button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        # Centra il pulsante nel suo frame
        button_frame.grid_columnconfigure(0, weight=1)
        btn_save = ttk.Button(button_frame, text="Salva Configurazione", command=self.save_config, style="Accent.TButton")
        btn_save.grid(row=0, column=0, pady=5)

        # Chiamata finale per popolare i dati dipendenti dal profilo
        self.on_profile_change()

    def get_nested_data(self, keys):
        data = self.config_data
        for key in keys:
            data = data[key]
        return data

    def set_nested_data(self, keys, new_value):
        data = self.config_data
        for key in keys[:-1]:
            data = data[key]
        data[keys[-1]] = new_value

    def create_generic_tab(self, notebook, keys, title):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=title)

        canvas = tk.Canvas(tab)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.populate_generic_tab(scrollable_frame, keys)

    def populate_generic_tab(self, parent, keys):
        frame = ttk.LabelFrame(parent, text="Impostazioni")
        frame.pack(padx=10, pady=10, fill="x", expand=True)

        # 1. Carica la sezione di configurazione principale (obbligatoria)
        try:
            data_section = self.get_nested_data(keys)
        except KeyError:
            # Se la sezione di configurazione non esiste in config.json, la tab non può essere creata.
            ttk.Label(frame, text=f"Sezione di configurazione '{keys}' non trovata.").pack()
            return

        # 2. Carica la sezione dei tooltip corrispondente (opzionale)
        tooltip_section = {}
        if self.tooltips_data:
            tooltip_keys = keys
            # Se stiamo popolando la tab ODC, il percorso dei tooltip è fisso e non dipende dal nome del profilo
            if keys and keys[-1] == "impostazioni_odc":
                tooltip_keys = ("file_e_fogli_excel", "mappatura_colonne_profili", "profili", "impostazioni_odc")

            try:
                tooltip_section = self.get_nested_data(tooltip_keys, data_source=self.tooltips_data)
            except KeyError:
                # Non è un errore critico. Significa solo che non ci sono tooltip per questa sezione.
                print(f"Avviso: Nessun tooltip trovato per la sezione '{keys}'.")
                pass

        critical_params = ["ritardo_click_singolo", "ritardo_prima_incolla", "ritardo_dopo_tab"]
        important_params = ["ritardo_dopo_copia_excel", "ritardo_dopo_select_all", "pausa_1", "pausa_2", "pausa_3", "riconferma_copia_pausa"]
        profile_dependant_params = ["regione_popup_odc_gia_registrato", "regione_popup_matricola_disabilitata"]


        self.vars[keys] = {}
        for i, (field, value) in enumerate(data_section.items()):
            label_frame = ttk.Frame(frame)
            label_frame.grid(row=i, column=0, sticky="w", padx=5, pady=5)

            label_text = f"{field}:"
            label_color = "black"
            label_font = ("Arial", 9)
            label_bg = None

            if field in critical_params:
                label_text = f"* [CRITICO] {field}:"
                label_color = "darkred"
                label_font = ("Arial", 9, "bold")
            elif field in important_params:
                label_text = f"* [IMPORTANTE] {field}:"
                label_color = "#E65100"
                label_font = ("Arial", 9, "bold")
            # Applica lo sfondo giallo solo se siamo nella tab ODC e il campo è uno di quelli variabili
            is_odc_tab = keys and keys[-1] == "impostazioni_odc"
            if is_odc_tab and field in profile_dependant_params:
                label_bg = "yellow"

            label_widget = ttk.Label(label_frame, text=label_text, foreground=label_color, font=label_font)
            if label_bg:
                label_widget.configure(background=label_bg)
            label_widget.pack(side="left")


            tooltip_text = tooltip_section.get(field, "Nessun aiuto disponibile.")
            if self.tooltips_data:
                help_icon = ttk.Label(label_frame, text=" (?)", cursor="hand2", foreground="blue")
                help_icon.pack(side="left", padx=(2,0))
                ToolTip(help_icon, tooltip_text)

            var = tk.StringVar(value=str(value))
            self.vars[keys][field] = var

            if isinstance(value, bool):
                widget = ttk.Combobox(frame, textvariable=var, values=["True", "False"], state="readonly")
            elif "percorso" in field or "path" in field:
                widget_frame = ttk.Frame(frame)
                entry = ttk.Entry(widget_frame, textvariable=var, width=60)
                entry.pack(side="left", fill="x", expand=True)

                browse_btn = ttk.Button(widget_frame, text="Sfoglia...", command=lambda v=var: self.browse_file(v))
                browse_btn.pack(side="left", padx=(5, 2))

                if field == "path_tesseract_cmd":
                    detect_btn = ttk.Button(widget_frame, text="Rileva", command=lambda v=var: self.autodetect_tesseract(v))
                    detect_btn.pack(side="left", padx=(0, 5))

                widget = widget_frame
            else:
                widget = ttk.Entry(frame, textvariable=var, width=50)

            widget.grid(row=i, column=1, sticky="ew", padx=5, pady=5)

            if "coordinate" in field or "regione" in field:
                capture_mode = 'region' if 'regione' in field else 'point'
                btn_capture = ttk.Button(frame, text="Cattura", command=lambda v=var, m=capture_mode: self.start_capture(v, m))
                btn_capture.grid(row=i, column=2, padx=5)

        frame.columnconfigure(1, weight=1)

    def get_nested_data(self, keys, data_source=None):
        if data_source is None:
            data_source = self.config_data

        data = data_source
        for key in keys:
            data = data[key]
        return data

    def create_profile_management_tab(self, parent_tab):
        profile_frame = ttk.LabelFrame(parent_tab, text="Gestione Profili")
        profile_frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(profile_frame, text="Profilo Attivo:").grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.profile_data = self.get_nested_data(("file_e_fogli_excel", "mappatura_colonne_profili"))
        self.active_profile_var = tk.StringVar(value=self.profile_data['profilo_attivo'])

        profiles = list(self.profile_data['profili'].keys())
        self.profile_menu = ttk.Combobox(profile_frame, textvariable=self.active_profile_var, values=profiles, state="readonly", width=40)
        self.profile_menu.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.profile_menu.bind("<<ComboboxSelected>>", self.on_profile_change)

        btn_add = ttk.Button(profile_frame, text="Aggiungi Profilo", command=self.add_profile)
        btn_add.grid(row=1, column=0, padx=5, pady=10)
        btn_rename = ttk.Button(profile_frame, text="Rinomina Selezionato", command=self.rename_profile)
        btn_rename.grid(row=1, column=1, padx=5, pady=10, sticky="w")
        btn_delete = ttk.Button(profile_frame, text="Elimina Selezionato", command=self.delete_profile)
        btn_delete.grid(row=1, column=2, padx=5, pady=10, sticky="w")

        profile_frame.columnconfigure(1, weight=1)

    def draw_mapping_entries(self):
        # This check is crucial for the very first run before the profile var is set
        if not hasattr(self, 'mapping_tab_frame'):
            return

        for widget in self.mapping_tab_frame.winfo_children():
            widget.destroy()

        active_profile_name = self.active_profile_var.get()
        if not active_profile_name: return

        canvas = tk.Canvas(self.mapping_tab_frame)
        scrollbar = ttk.Scrollbar(self.mapping_tab_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Label(scrollable_frame, text="Nome Campo", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Label(scrollable_frame, text="Colonna Excel", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(scrollable_frame, text="Coordinata X", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=5)

        current_profile_data = self.profile_data['profili'][active_profile_name]['mappature']
        self.mapping_vars[active_profile_name] = []

        for i, item in enumerate(current_profile_data):
            row_vars = {}
            ttk.Label(scrollable_frame, text=item['nome']).grid(row=i + 1, column=0, sticky="w", padx=5, pady=5)

            var_col = tk.StringVar(value=item['colonna_excel'])
            row_vars['colonna_excel'] = var_col
            ttk.Entry(scrollable_frame, textvariable=var_col, width=15).grid(row=i + 1, column=1, padx=5, pady=5)

            var_x = tk.StringVar(value=item['target_x_remoto'])
            row_vars['target_x_remoto'] = var_x

            x_entry_frame = ttk.Frame(scrollable_frame)
            x_entry_frame.grid(row=i + 1, column=2, padx=5, pady=5, sticky="ew")
            ttk.Entry(x_entry_frame, textvariable=var_x, width=10).pack(side="left")
            btn_capture = ttk.Button(x_entry_frame, text="Cattura", command=lambda v=var_x: self.start_capture_x_only(v))
            btn_capture.pack(side="left", padx=5)

            row_vars['nome'] = item['nome']
            self.mapping_vars[active_profile_name].append(row_vars)

    def draw_odc_settings(self):
        for widget in self.odc_tab_frame.winfo_children():
            widget.destroy()

        active_profile_name = self.active_profile_var.get()
        if not active_profile_name: return

        # Usiamo populate_generic_tab per creare dinamicamente i campi
        # Definiamo un percorso "virtuale" per le chiavi che punti ai dati del profilo attivo
        keys = ("file_e_fogli_excel", "mappatura_colonne_profili", "profili", active_profile_name, "impostazioni_odc")
        self.populate_generic_tab(self.odc_tab_frame, keys)


    def on_profile_change(self, event=None):
        # Ridisegna le tab che dipendono dal profilo
        self.draw_mapping_entries()
        self.draw_odc_settings()

    def update_profile_list(self):
        profiles = list(self.profile_data['profili'].keys())
        self.profile_menu['values'] = profiles
        if self.active_profile_var.get() not in profiles:
            self.active_profile_var.set(profiles[0] if profiles else "")
        self.on_profile_change()

    def add_profile(self):
        from tkinter.simpledialog import askstring
        new_name = askstring("Nuovo Profilo", "Inserisci il nome del nuovo profilo:", parent=self.parent)
        if new_name and new_name not in self.profile_data['profili']:
            current_profile_name = self.active_profile_var.get()
            # Copia profonda del profilo selezionato come base
            import copy
            self.profile_data['profili'][new_name] = copy.deepcopy(self.profile_data['profili'][current_profile_name])
            self.active_profile_var.set(new_name)
            self.update_profile_list()
        elif new_name:
            messagebox.showwarning("Errore", "Un profilo con questo nome esiste già.", parent=self.parent)

    def rename_profile(self):
        from tkinter.simpledialog import askstring
        old_name = self.active_profile_var.get()
        if not old_name: return
        new_name = askstring("Rinomina Profilo", f"Inserisci il nuovo nome per '{old_name}':", parent=self.parent)
        if new_name and new_name not in self.profile_data['profili']:
            self.profile_data['profili'][new_name] = self.profile_data['profili'].pop(old_name)
            self.active_profile_var.set(new_name)
            self.update_profile_list()
        elif new_name:
            messagebox.showwarning("Errore", "Un profilo con questo nome esiste già.", parent=self.parent)

    def delete_profile(self):
        if len(self.profile_data['profili']) <= 1:
            messagebox.showwarning("Impossibile Eliminare", "Non puoi eliminare l'ultimo profilo rimasto.", parent=self.parent)
            return

        profile_to_delete = self.active_profile_var.get()
        if messagebox.askyesno("Conferma Eliminazione", f"Sei sicuro di voler eliminare il profilo '{profile_to_delete}'?", parent=self.parent):
            del self.profile_data['profili'][profile_to_delete]
            self.update_profile_list()

    def browse_file(self, var_to_update):
        filepath = filedialog.askopenfilename(title="Seleziona un file", filetypes=(("Tutti i file", "*.*"),))
        if filepath:
            var_to_update.set(filepath)

    def autodetect_tesseract(self, var_to_update):
        """Cerca tesseract.exe in percorsi comuni e aggiorna la variabile."""
        print("Avvio rilevamento automatico di Tesseract...")
        potential_paths = [
            "C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
            "C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe"
        ]

        found_path = None
        for path in potential_paths:
            if os.path.exists(path):
                found_path = path
                break

        if found_path:
            var_to_update.set(found_path)
            messagebox.showinfo("Successo", f"Tesseract trovato e impostato a:\n{found_path}", parent=self)
        else:
            messagebox.showwarning("Fallito", "Impossibile trovare Tesseract nei percorsi predefiniti.\n\nPer favore, selezionalo manualmente usando 'Sfoglia...'.", parent=self)
            # Apri la finestra di dialogo come fallback
            self.browse_file(var_to_update)

    def start_capture(self, var_to_update, mode='point'):
        def on_capture(coords):
            var_to_update.set(str(list(coords)))
            # Assicurati che la finestra principale torni in primo piano
            self.winfo_toplevel().focus_force()
            self.winfo_toplevel().lift()
        CaptureHelper(self, on_capture, capture_mode=mode)

    def start_capture_x_only(self, var_to_update):
        def on_capture(coords):
            var_to_update.set(str(coords[0])) # Salva solo la coordinata X
            # Assicurati che la finestra principale torni in primo piano
            self.winfo_toplevel().focus_force()
            self.winfo_toplevel().lift()
        CaptureHelper(self, on_capture)

    def update_config_data(self):
        # Gestione delle tab generiche e dei dati ODC per profilo
        for keys, section_vars in self.vars.items():
            # Questo try-except gestisce il caso in cui una sezione (es. un profilo appena eliminato)
            # esista ancora in self.vars ma non più in self.config_data
            try:
                original_section = self.get_nested_data(keys)
                if isinstance(original_section, dict):
                    for field, var in section_vars.items():
                        # Controlla che il campo esista ancora prima di aggiornarlo
                        if field in original_section:
                            original_value = original_section[field]
                            new_value_str = var.get()
                            try:
                                if isinstance(original_value, bool):
                                    original_section[field] = (new_value_str.lower() == 'true')
                                elif isinstance(original_value, int):
                                    original_section[field] = int(new_value_str)
                                elif isinstance(original_value, float):
                                    original_section[field] = float(new_value_str)
                                elif isinstance(original_value, list):
                                    # Gestisce sia le liste JSON valide che le stringhe semplici
                                    try:
                                        original_section[field] = json.loads(new_value_str.replace("'", '"'))
                                    except json.JSONDecodeError:
                                        original_section[field] = new_value_str # Fallback per stringhe non-JSON
                                else:
                                    original_section[field] = new_value_str
                            except (ValueError, json.JSONDecodeError):
                                # In caso di errore di conversione, salva come stringa grezza
                                original_section[field] = new_value_str
            except KeyError:
                print(f"Avviso: Sezione di configurazione per le chiavi '{keys}' non trovata durante il salvataggio. Potrebbe essere stata rimossa.")
                continue


        # Gestione della tab di mappatura profili
        self.profile_data['profilo_attivo'] = self.active_profile_var.get()
        for profile_name, profile_vars_list in self.mapping_vars.items():
            if profile_name in self.profile_data['profili']:
                new_mapping_data = []
                for row_vars in profile_vars_list:
                    new_item = {'nome': row_vars['nome']}
                    try:
                        new_item['colonna_excel'] = row_vars['colonna_excel'].get()
                        new_item['target_x_remoto'] = int(row_vars['target_x_remoto'].get())
                        new_mapping_data.append(new_item)
                    except (ValueError, KeyError):
                        # Ignora righe con dati non validi
                        print(f"Avviso: Riga di mappatura non valida per il profilo '{profile_name}' saltata.")
                        pass
                self.profile_data['profili'][profile_name]['mappature'] = new_mapping_data


    def save_config(self):
        self.update_config_data()
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config_data, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Successo", "Configurazione salvata con successo!", parent=self)
        except Exception as e:
            messagebox.showerror("Errore di Salvataggio", f"Impossibile salvare 'config.json':\n{e}", parent=self)