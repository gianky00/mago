import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pyautogui

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
        # Frame principale per contenere tutto, inclusi i pulsanti
        main_frame = ttk.Frame(self)
        main_frame.pack(expand=True, fill="both")

        main_notebook = ttk.Notebook(main_frame)
        main_notebook.pack(pady=10, padx=10, expand=True, fill="both")

        # --- Main Tab 1: File e Fogli Excel ---
        excel_tab = ttk.Frame(main_notebook)
        main_notebook.add(excel_tab, text="File e Fogli Excel")
        excel_notebook = ttk.Notebook(excel_tab)
        excel_notebook.pack(pady=5, padx=5, expand=True, fill="both")

        self.create_generic_tab(excel_notebook, ("file_e_fogli_excel", "impostazioni_file"), "Impostazioni File")
        self.create_generic_tab(excel_notebook, ("file_e_fogli_excel", "mappature_colonne_foglio_avanzamento"), "Mappature Colonne Foglio Avanzamento")
        self.create_mapping_tab(excel_notebook, ("file_e_fogli_excel", "mappatura_colonne_foglio_dati"), "Mappatura Colonne Foglio Dati")

        # --- Main Tab 2: Coordinate e Dati ---
        coords_tab = ttk.Frame(main_notebook)
        main_notebook.add(coords_tab, text="Coordinate e Dati")
        coords_notebook = ttk.Notebook(coords_tab)
        coords_notebook.pack(pady=5, padx=5, expand=True, fill="both")

        self.create_generic_tab(coords_notebook, ("coordinate_e_dati", "gui"), "GUI")
        self.create_generic_tab(coords_notebook, ("coordinate_e_dati", "odc"), "Registrazione ODC")

        # --- Main Tab 3: Altre Impostazioni ---
        other_tab = ttk.Frame(main_notebook)
        main_notebook.add(other_tab, text="Altre Impostazioni")
        other_notebook = ttk.Notebook(other_tab)
        other_notebook.pack(pady=5, padx=5, expand=True, fill="both")

        self.create_generic_tab(other_notebook, ("generali",), "Generali")
        self.create_generic_tab(other_notebook, ("timing_e_ritardi",), "Timing")
        self.create_generic_tab(other_notebook, ("pulizia_appunti",), "Pulizia Appunti")
        self.create_generic_tab(other_notebook, ("tasti_rapidi",), "Tasti Rapidi")

        # --- Pulsante Salva ---
        btn_save = ttk.Button(main_frame, text="Salva Configurazione", command=self.save_config, style="Accent.TButton")
        btn_save.pack(pady=15, padx=10)

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

        try:
            data_section = self.get_nested_data(keys)
            tooltip_section = self.get_nested_data(keys, data_source=self.tooltips_data) if self.tooltips_data else {}
        except KeyError:
            ttk.Label(frame, text=f"Sezione non trovata.").pack()
            return

        critical_params = ["ritardo_click_singolo", "ritardo_prima_incolla", "ritardo_dopo_select_all", "ritardo_dopo_copia_excel"]

        self.vars[keys] = {}
        for i, (field, value) in enumerate(data_section.items()):
            label_frame = ttk.Frame(frame)
            label_frame.grid(row=i, column=0, sticky="w", padx=5, pady=5)

            is_critical = field in critical_params
            label_text = f" * {field}:" if is_critical else f"{field}:"
            label_color = "darkred" if is_critical else "black"
            label_font = ("Arial", 9, "bold") if is_critical else ("Arial", 9)

            ttk.Label(label_frame, text=label_text, foreground=label_color, font=label_font).pack(side="left")

            tooltip_text = tooltip_section.get(field, "Nessun aiuto disponibile.")
            if self.tooltips_data:
                help_icon = ttk.Label(label_frame, text=" (?)", cursor="hand2", foreground="blue")
                help_icon.pack(side="left", padx=(2,0))
                ToolTip(help_icon, tooltip_text)

            var = tk.StringVar(value=str(value))
            self.vars[keys][field] = var

            if isinstance(value, bool):
                widget = ttk.Combobox(frame, textvariable=var, values=["True", "False"], state="readonly")
            elif "percorso" in field:
                widget_frame = ttk.Frame(frame)
                entry = ttk.Entry(widget_frame, textvariable=var, width=60)
                entry.pack(side="left", fill="x", expand=True)
                btn = ttk.Button(widget_frame, text="Sfoglia...", command=lambda v=var: self.browse_file(v))
                btn.pack(side="left", padx=5)
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

    def create_mapping_tab(self, notebook, keys, title):
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

        self.populate_mapping_tab(scrollable_frame, keys)

    def populate_mapping_tab(self, parent, keys):
        frame = ttk.LabelFrame(parent, text="Mappatura")
        frame.pack(padx=10, pady=10, fill="x", expand=True)

        try:
            data_section = self.get_nested_data(keys)
        except KeyError:
            ttk.Label(frame, text=f"Sezione non trovata.").pack()
            return

        ttk.Label(frame, text="Nome Campo", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Label(frame, text="Colonna Excel", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(frame, text="Coordinata X", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=5)

        self.vars[keys] = []
        for i, item in enumerate(data_section):
            row_vars = {}
            ttk.Label(frame, text=item['nome']).grid(row=i + 1, column=0, sticky="w", padx=5, pady=5)

            var_col = tk.StringVar(value=item['colonna_excel'])
            row_vars['colonna_excel'] = var_col
            ttk.Entry(frame, textvariable=var_col, width=15).grid(row=i + 1, column=1, padx=5, pady=5)

            var_x = tk.StringVar(value=item['target_x_remoto'])
            row_vars['target_x_remoto'] = var_x

            x_entry_frame = ttk.Frame(frame)
            x_entry_frame.grid(row=i + 1, column=2, padx=5, pady=5, sticky="ew")

            ttk.Entry(x_entry_frame, textvariable=var_x, width=10).pack(side="left")
            btn_capture = ttk.Button(x_entry_frame, text="Cattura", command=lambda v=var_x: self.start_capture_x_only(v))
            btn_capture.pack(side="left", padx=5)

            row_vars['nome'] = item['nome']
            self.vars[keys].append(row_vars)

    def browse_file(self, var_to_update):
        filepath = filedialog.askopenfilename(title="Seleziona un file", filetypes=(("Tutti i file", "*.*"),))
        if filepath:
            var_to_update.set(filepath)

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
        for keys, section_vars in self.vars.items():
            original_section = self.get_nested_data(keys)

            if isinstance(original_section, list):
                new_data = []
                for row_vars in section_vars:
                    new_item = {'nome': row_vars['nome']}
                    try:
                        new_item['colonna_excel'] = row_vars['colonna_excel'].get()
                        new_item['target_x_remoto'] = int(row_vars['target_x_remoto'].get())
                    except (ValueError, KeyError):
                        pass
                    new_data.append(new_item)
                self.set_nested_data(keys, new_data)

            elif isinstance(original_section, dict):
                for field, var in section_vars.items():
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
                            original_section[field] = json.loads(new_value_str.replace("'", '"'))
                        else:
                            original_section[field] = new_value_str
                    except (ValueError, json.JSONDecodeError):
                        original_section[field] = new_value_str

    def save_config(self):
        self.update_config_data()
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config_data, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Successo", "Configurazione salvata con successo!", parent=self)
        except Exception as e:
            messagebox.showerror("Errore di Salvataggio", f"Impossibile salvare 'config.json':\n{e}", parent=self)