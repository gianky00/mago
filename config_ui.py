import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pyautogui
import keyboard

class CaptureHelper:
    def __init__(self, parent, on_capture_callback):
        self.parent = parent
        self.on_capture_callback = on_capture_callback

        self.capture_window = tk.Toplevel(parent)
        self.capture_window.attributes('-alpha', 0.5)
        self.capture_window.attributes('-fullscreen', True)
        self.capture_window.configure(bg='grey')

        label = tk.Label(self.capture_window, text="Muovi il mouse sulla posizione desiderata e premi 'Ctrl' per catturare.",
                         font=("Arial", 18), bg="white", fg="black")
        label.pack(pady=100)

        self.parent.withdraw()
        keyboard.add_hotkey('ctrl', self.capture_coords, suppress=True)

    def capture_coords(self):
        x, y = pyautogui.position()
        keyboard.remove_hotkey('ctrl')
        self.capture_window.destroy()
        self.parent.deiconify()
        self.on_capture_callback((x, y))

class ConfigWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Configurazione")
        self.geometry("700x750")
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

        # Create tabs dynamically
        self.create_tab(notebook, 'generali', "Impostazioni Generali")
        self.create_tab(notebook, 'file_excel', "File e Fogli Excel")
        self.create_tab(notebook, 'automazione_gui', "Coordinate GUI")
        self.create_tab(notebook, 'procedura_odc', "Coordinate Procedura ODC")
        self.create_tab(notebook, 'timing_e_ritardi', "Timing e Ritardi")

        btn_save = ttk.Button(self, text="Salva e Chiudi", command=self.save_config)
        btn_save.pack(pady=10)

    def create_tab(self, notebook, key, title):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=title)

        # Use a canvas and scrollbar for tabs that might overflow
        canvas = tk.Canvas(tab)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Populate the frame
        self.populate_tab(scrollable_frame, key)

    def populate_tab(self, parent, key):
        frame = ttk.LabelFrame(parent, text=key.replace('_', ' ').title())
        frame.pack(padx=10, pady=10, fill="x")

        if key not in self.config_data:
            ttk.Label(frame, text=f"Sezione '{key}' non trovata in config.json").pack()
            return

        self.vars[key] = {}
        for i, (field, value) in enumerate(self.config_data[key].items()):
            ttk.Label(frame, text=f"{field}:").grid(row=i, column=0, sticky="w", padx=5, pady=5)

            var = tk.StringVar(value=str(value))
            self.vars[key][field] = var

            entry = ttk.Entry(frame, textvariable=var, width=50)
            entry.grid(row=i, column=1, sticky="ew", padx=5, pady=5)

            if "coordinate" in field or "regione" in field:
                # Pass a lambda to capture the current var
                btn = ttk.Button(frame, text="Cattura", command=lambda v=var: self.start_capture(v))
                btn.grid(row=i, column=2, padx=5)

        frame.columnconfigure(1, weight=1)

    def start_capture(self, var_to_update):
        def on_capture(coords):
            # For regions, we need to handle it differently, but for now, just set coords
            var_to_update.set(str(list(coords)))
            self.parent.focus_force() # Bring focus back to the main app
            self.lift() # Bring config window to front

        CaptureHelper(self, on_capture)

    def update_config_data(self):
        for section, fields in self.vars.items():
            for field, var in fields.items():
                original_value = self.config_data[section][field]
                new_value_str = var.get()

                # Try to convert back to the original type (int, float, list, bool)
                try:
                    if isinstance(original_value, bool):
                        self.config_data[section][field] = new_value_str.lower() in ['true', '1', 'yes']
                    elif isinstance(original_value, int):
                        self.config_data[section][field] = int(new_value_str)
                    elif isinstance(original_value, float):
                        self.config_data[section][field] = float(new_value_str)
                    elif isinstance(original_value, list):
                        # Safely evaluate string to list
                        self.config_data[section][field] = json.loads(new_value_str.replace("'", '"'))
                    else: # string
                        self.config_data[section][field] = new_value_str
                except (ValueError, json.JSONDecodeError):
                    # If conversion fails, keep it as a string
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
    root.protocol("WM_DELETE_WINDOW", lambda: root.destroy())
    root.mainloop()