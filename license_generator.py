import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import os
import json
import shutil
import sys

SAVED_LICENSES_FILE = "saved_licenses.json"

class LicenseGenerator(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Generatore di Licenze e Pacchetti")
        self.geometry("800x650")

        # --- Stile ---
        style = ttk.Style(self)
        style.configure("TLabel", padding=5, font=('Helvetica', 10))
        style.configure("TEntry", padding=5, font=('Helvetica', 10))
        style.configure("TButton", padding=5, font=('Helvetica', 10, 'bold'))
        style.configure("Bold.TButton", font=('Helvetica', 11, 'bold'))
        style.configure("TFrame", padding=10)
        style.configure("TLabelframe.Label", font=('Helvetica', 11, 'bold'))

        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Frame FASE 1: Utility Info Hardware ---
        hw_utility_frame = ttk.LabelFrame(main_frame, text="FASE 1: Crea Utility per Info Hardware")
        hw_utility_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(hw_utility_frame, text="Crea un eseguibile da inviare al collega per ottenere il suo ID Hardware.").pack(anchor='w', padx=5, pady=5)
        self.create_hw_utility_button = ttk.Button(hw_utility_frame, text="Crea Utility Info Hardware...", command=self.create_hw_utility)
        self.create_hw_utility_button.pack(fill=tk.X, padx=5, pady=5)

        # --- Frame FASE 2: Gestione Licenze ---
        manage_licenses_frame = ttk.LabelFrame(main_frame, text="FASE 2: Salva e Gestisci Licenze")
        manage_licenses_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(manage_licenses_frame, text="ID Licenza (es. 'PC Ufficio Mario Rossi'):").grid(row=0, column=0, sticky='w', padx=5)
        self.license_id_entry = ttk.Entry(manage_licenses_frame, width=40)
        self.license_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        ttk.Label(manage_licenses_frame, text="ID Hardware (ricevuto dal collega):").grid(row=1, column=0, sticky='w', padx=5)
        self.hw_id_entry = ttk.Entry(manage_licenses_frame, width=60)
        self.hw_id_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

        self.save_license_button = ttk.Button(manage_licenses_frame, text="Salva Licenza", command=self.save_license)
        self.save_license_button.grid(row=0, column=2, rowspan=2, padx=10, pady=5, sticky='ns')
        manage_licenses_frame.columnconfigure(1, weight=1)

        # --- Frame FASE 3: Crea Pacchetto Finale ---
        final_package_frame = ttk.LabelFrame(main_frame, text="FASE 3: Genera Pacchetto Licenziato Finale")
        final_package_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(final_package_frame, text="Seleziona Licenza Salvata:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.saved_licenses_combo = ttk.Combobox(final_package_frame, state="readonly", width=38)
        self.saved_licenses_combo.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.saved_licenses_combo.bind("<<ComboboxSelected>>", self.on_license_selected)

        ttk.Label(final_package_frame, text="Data Scadenza (YYYY-MM-DD):").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.expire_date_entry = ttk.Entry(final_package_frame, width=20)
        self.expire_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.expire_date_entry.insert(0, "2025-12-31")

        self.generate_button = ttk.Button(final_package_frame, text="Genera Pacchetto Finale...", style="Bold.TButton", command=self.generate_final_package)
        self.generate_button.grid(row=0, column=2, rowspan=2, padx=10, pady=10, sticky='ns')
        final_package_frame.columnconfigure(1, weight=1)

        # --- Output Log ---
        log_frame = ttk.LabelFrame(main_frame, text="Log di Output")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, bg="#f0f0f0")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.licenses = {}
        self.load_licenses()

    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.update_idletasks()

    def run_command(self, command):
        self.log(f"Eseguo: {' '.join(command)}")
        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None: break
                if output: self.log(output.strip())
            return process.poll() == 0
        except Exception as e:
            self.log(f"ERRORE CRITICO: {e}")
            messagebox.showerror("Errore Critico", f"Impossibile eseguire il comando PyArmor.\n{e}")
            return False

    def load_licenses(self):
        try:
            if os.path.exists(SAVED_LICENSES_FILE):
                with open(SAVED_LICENSES_FILE, 'r') as f:
                    self.licenses = json.load(f)
            else:
                self.licenses = {}

            self.update_licenses_dropdown()
            self.log(f"Caricate {len(self.licenses)} licenze salvate da '{SAVED_LICENSES_FILE}'.")
        except Exception as e:
            self.log(f"Errore nel caricamento di {SAVED_LICENSES_FILE}: {e}")
            self.licenses = {}

    def save_license(self):
        license_id = self.license_id_entry.get().strip()
        hw_id = self.hw_id_entry.get().strip()
        if not license_id or not hw_id:
            messagebox.showerror("Errore", "ID Licenza e ID Hardware sono obbligatori.")
            return

        self.licenses[license_id] = hw_id
        try:
            with open(SAVED_LICENSES_FILE, 'w') as f:
                json.dump(self.licenses, f, indent=4)
            self.log(f"Licenza '{license_id}' salvata con successo.")
            self.update_licenses_dropdown()
            self.license_id_entry.delete(0, tk.END)
            self.hw_id_entry.delete(0, tk.END)
        except Exception as e:
            self.log(f"Errore nel salvataggio della licenza: {e}")
            messagebox.showerror("Errore", f"Impossibile salvare il file delle licenze:\n{e}")

    def update_licenses_dropdown(self):
        self.saved_licenses_combo['values'] = sorted(list(self.licenses.keys()))
        if self.saved_licenses_combo['values']:
            self.saved_licenses_combo.current(0)
            self.on_license_selected(None)

    def on_license_selected(self, event):
        selected_id = self.saved_licenses_combo.get()
        hw_id = self.licenses.get(selected_id, "")
        self.hw_id_entry.delete(0, tk.END)
        self.hw_id_entry.insert(0, hw_id)

    def create_hw_utility(self):
        output_zip = filedialog.asksaveasfilename(
            title="Salva l'utility per l'info hardware come ZIP",
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip")]
        )
        if not output_zip:
            self.log("Creazione utility annullata.")
            return

        temp_dir = "temp_hw_utility"
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
        os.makedirs(temp_dir)

        self.log("Inizio creazione utility info hardware...")

        # Creazione script di avvio per Windows (più robusto)
        bat_content = """\
@echo off
:: Imposta la directory di lavoro e il percorso di ricerca di Python
cd /d "%~dp0"
set PYTHONPATH=%~dp0

:: Controlla se i file necessari esistono
IF NOT EXIST "get_hw_info.py" (
    echo ERRORE: File 'get_hw_info.py' non trovato.
    echo Assicurati di aver scompattato tutti i file dallo ZIP nella stessa cartella.
    pause
    exit /b
)

IF NOT EXIST "pyarmor_runtime_*" (
    echo ERRORE: Cartella di runtime 'pyarmor_runtime_*' non trovata.
    echo Assicurati di aver scompattato tutti i file dallo ZIP nella stessa cartella.
    pause
    exit /b
)

:: Esegui lo script python
pythonw.exe get_hw_info.py
"""
        with open(os.path.join(temp_dir, "AVVIA_PER_ID_HARDWARE.bat"), "w") as f:
            f.write(bat_content)

        command = ["pyarmor", "gen", "-O", temp_dir, "get_hw_info.py"]

        if self.run_command(command):
            self.log("Offuscamento completato. Creo l'archivio ZIP...")
            shutil.make_archive(output_zip.replace('.zip', ''), 'zip', temp_dir)
            self.log(f"Utility creata con successo: {output_zip}")
            messagebox.showinfo("Successo", f"Utility per info hardware creata con successo!\nInvia il file '{os.path.basename(output_zip)}' al tuo collega.")
        else:
            self.log("ERRORE durante la creazione dell'utility.")
            messagebox.showerror("Fallito", "Creazione utility fallita. Controlla il log.")

        shutil.rmtree(temp_dir)

    def generate_final_package(self):
        selected_license_id = self.saved_licenses_combo.get()
        expire_date = self.expire_date_entry.get().strip()

        if not selected_license_id or not expire_date:
            messagebox.showerror("Errore", "Seleziona una licenza e inserisci una data di scadenza.")
            return

        hw_id = self.licenses[selected_license_id]

        output_dir = filedialog.askdirectory(title="Seleziona la cartella di destinazione per il pacchetto finale")
        if not output_dir:
            self.log("Generazione pacchetto finale annullata.")
            return

        self.log(f"Inizio generazione pacchetto per '{selected_license_id}'...")
        scripts = ["gui.py", "magoPyton.py", "config_ui.py", "catturaPosizioneMouse.py"]
        for script in scripts:
            if not os.path.exists(script):
                messagebox.showerror("Errore", f"Script non trovato: {script}")
                return

        command = ["pyarmor", "gen", "--expired", expire_date, "--bind-hd", hw_id, "-O", output_dir] + scripts

        if self.run_command(command):
            self.log("Offuscamento completato. Copio i file ausiliari...")
            files_to_copy = ["config.json", "tooltips.json", "avvio.bat", "Dataease_ALLEGRETTI_02-2025.xlsm"]
            for f in files_to_copy:
                if os.path.exists(f):
                    shutil.copy(f, output_dir)
                    self.log(f"Copiato: {f}")
                else:
                    self.log(f"ATTENZIONE: File ausiliario non trovato: {f}")

            self.log("\nGenerazione completata con successo!")
            messagebox.showinfo("Successo", f"Pacchetto finale generato con successo in:\n{output_dir}")
        else:
            self.log("ERRORE durante la generazione del pacchetto finale.")
            messagebox.showerror("Fallito", "Generazione pacchetto finale fallita. Controlla il log.")

if __name__ == "__main__":
    app = LicenseGenerator()
    app.mainloop()