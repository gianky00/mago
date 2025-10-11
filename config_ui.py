import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pyautogui
import keyboard

class CaptureHelper:
    """A helper class to create a temporary, fullscreen, transparent window
    for capturing mouse coordinates with a simple click."""
    def __init__(self, parent, on_capture_callback):
        self.parent = parent
        self.on_capture_callback = on_capture_callback

        self.capture_window = tk.Toplevel(parent)
        self.capture_window.attributes('-alpha', 0.4)  # Slightly more transparent
        self.capture_window.attributes('-fullscreen', True)
        self.capture_window.configure(bg='grey', cursor="crosshair")  # Set a crosshair cursor

        # Center the label
        label_frame = tk.Frame(self.capture_window, bg="grey")
        label_frame.pack(expand=True)
        label = tk.Label(label_frame, text="Muovi il mouse sulla posizione desiderata e FAI CLICK per catturare.",
                         font=("Arial", 22, "bold"), bg="white", fg="blue", relief="solid", borderwidth=2)
        label.pack(pady=20, padx=20)

        self.parent.withdraw()  # Hide the config window to prevent it from being captured
        self.capture_window.bind("<Button-1>", self.capture_coords) # Bind left-click
        self.capture_window.focus_force() # Ensure the window has focus

    def capture_coords(self, event=None):
        """Callback for when the user clicks."""
        x, y = pyautogui.position()
        self.capture_window.destroy()
        self.parent.deiconify()  # Show the config window again
        self.on_capture_callback((x, y))

class ConfigWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Configurazione")
        self.geometry("750x800")
        self.parent = parent

        self.config_data = self.load_config()
        if not self.config_data:
            self.destroy()
            return

        self.vars = {}
        self.create_widgets()

    def load_config(self):
        try:
            with open('config.json', 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            messagebox.showerror("Errore", f"Impossibile caricare config.json: {e}", parent=self)
            return None

    def create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(pady=10, padx=10, expand=True, fill="both")

        # Create tabs
        self.create_tab(notebook, 'generali', "Impostazioni Generali")
        self.create_tab(notebook, 'file_excel', "File e Fogli Excel")
        self.create_tab(notebook, 'automazione_gui', "Coordinate GUI")
        self.create_tab(notebook, 'procedura_odc', "Coordinate ODC")
        self.create_tab(notebook, 'timing_e_ritardi', "Timing")
        self.create_tab(notebook, 'parametri_ricerca', "Parametri Ricerca")

        btn_save = ttk.Button(self, text="Salva e Chiudi", command=self.save_config)
        btn_save.pack(pady=10)

    def create_tab(self, notebook, key, title):
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

        self.populate_tab(scrollable_frame, key)

    def populate_tab(self, parent, key):
        frame = ttk.LabelFrame(parent, text=key.replace('_', ' ').title())
        frame.pack(padx=10, pady=10, fill="x", expand=True)

        if key not in self.config_data:
            ttk.Label(frame, text=f"Sezione '{key}' non trovata.").pack()
            return

        self.vars[key] = {}
        for i, (field, value) in enumerate(self.config_data[key].items()):
            ttk.Label(frame, text=f"{field}:").grid(row=i, column=0, sticky="w", padx=5, pady=5)

            var = tk.StringVar(value=str(value))
            self.vars[key][field] = var

            # --- WIDGET DINAMICI ---
            if isinstance(value, bool):
                widget = ttk.Combobox(frame, textvariable=var, values=["True", "False"], state="readonly")
            elif "path" in field or "file" in field:
                widget_frame = ttk.Frame(frame)
                entry = ttk.Entry(widget_frame, textvariable=var, width=60)
                entry.pack(side="left", fill="x", expand=True)
                btn = ttk.Button(widget_frame, text="Sfoglia...", command=lambda v=var: self.browse_file(v))
                btn.pack(side="left", padx=5)
                widget = widget_frame
            else:
                widget = ttk.Entry(frame, textvariable=var, width=50)
                # Make specific fields in 'parametri_ricerca' read-only as per user request
                if key == 'parametri_ricerca' and field in ['riga_inizio', 'riga_fine', 'stato_da_cercare']:
                    widget.config(state='readonly')

            widget.grid(row=i, column=1, sticky="ew", padx=5, pady=5)

            if "coordinate" in field or "regione" in field:
                btn_capture = ttk.Button(frame, text="Cattura", command=lambda v=var: self.start_capture(v))
                btn_capture.grid(row=i, column=2, padx=5)

        frame.columnconfigure(1, weight=1)

    def browse_file(self, var_to_update):
        filepath = filedialog.askopenfilename(
            title="Seleziona un file",
            filetypes=(("Tutti i file", "*.*"),)
        )
        if filepath:
            var_to_update.set(filepath)

    def start_capture(self, var_to_update):
        def on_capture(coords):
            var_to_update.set(str(list(coords)))
            self.parent.focus_force()
            self.lift()

        CaptureHelper(self, on_capture)

    def update_config_data(self):
        for section, fields in self.vars.items():
            for field, var in fields.items():
                original_value = self.config_data[section][field]
                new_value_str = var.get()

                try:
                    if isinstance(original_value, bool):
                        self.config_data[section][field] = (new_value_str.lower() == 'true')
                    elif isinstance(original_value, int):
                        self.config_data[section][field] = int(new_value_str)
                    elif isinstance(original_value, float):
                        self.config_data[section][field] = float(new_value_str)
                    elif isinstance(original_value, list):
                        self.config_data[section][field] = json.loads(new_value_str.replace("'", '"'))
                    else:
                        self.config_data[section][field] = new_value_str
                except (ValueError, json.JSONDecodeError):
                    self.config_data[section][field] = new_value_str

    def save_config(self):
        self.update_config_data()
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config_data, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Successo", "Configurazione salvata!", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile salvare config.json: {e}", parent=self)

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    app = ConfigWindow(root)
    root.mainloop()