import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import os

class LicenseGenerator(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Generatore di Licenze PyArmor")
        self.geometry("600x450")

        # --- Stile ---
        style = ttk.Style(self)
        style.configure("TLabel", padding=5, font=('Helvetica', 10))
        style.configure("TEntry", padding=5, font=('Helvetica', 10))
        style.configure("TButton", padding=5, font=('Helvetica', 10, 'bold'))
        style.configure("TFrame", padding=10)

        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Input Fields ---
        input_frame = ttk.LabelFrame(main_frame, text="Dati per la Licenza", padding=10)
        input_frame.pack(fill=tk.X, pady=5)

        # Hardware ID
        ttk.Label(input_frame, text="ID Hardware (da 'pyarmor hdinfo'):").grid(row=0, column=0, sticky=tk.W)
        self.hw_id_entry = ttk.Entry(input_frame, width=50)
        self.hw_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)

        # Expiration Date
        ttk.Label(input_frame, text="Data Scadenza (YYYY-MM-DD):").grid(row=1, column=0, sticky=tk.W)
        self.expire_date_entry = ttk.Entry(input_frame, width=20)
        self.expire_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.expire_date_entry.insert(0, "2025-12-31")

        input_frame.columnconfigure(1, weight=1)

        # --- Generate Button ---
        self.generate_button = ttk.Button(main_frame, text="Genera Pacchetto Licenziato", command=self.generate_license)
        self.generate_button.pack(fill=tk.X, pady=10)

        # --- Output Log ---
        log_frame = ttk.LabelFrame(main_frame, text="Log di Output", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, bg="#f0f0f0")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def log(self, message):
        """Aggiunge un messaggio al log."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.update_idletasks()

    def generate_license(self):
        """Costruisce ed esegue il comando PyArmor."""
        hw_id = self.hw_id_entry.get().strip()
        expire_date = self.expire_date_entry.get().strip()

        if not hw_id or not expire_date:
            messagebox.showerror("Errore", "ID Hardware e Data di Scadenza sono obbligatori.")
            return

        output_dir = filedialog.askdirectory(title="Seleziona la cartella di destinazione per il pacchetto")
        if not output_dir:
            self.log("Generazione annullata dall'utente.")
            return

        self.log("Inizio la generazione del pacchetto...")
        self.generate_button.config(state=tk.DISABLED)

        # Lista degli script da offuscare
        scripts_to_obfuscate = ["gui.py", "magoPyton.py", "config_ui.py", "catturaPosizioneMouse.py"]

        # Verifica che gli script esistano prima di lanciare il comando
        for script in scripts_to_obfuscate:
            if not os.path.exists(script):
                messagebox.showerror("Errore", f"Script non trovato: {script}\nAssicurati di eseguire questo generatore dalla cartella principale del progetto.")
                self.log(f"ERRORE: Script non trovato: {script}")
                self.generate_button.config(state=tk.NORMAL)
                return

        command = [
            "pyarmor", "gen",
            "--expired", expire_date,
            "--bind-hd", hw_id,
            "-O", output_dir,
        ] + scripts_to_obfuscate

        self.log(f"Eseguo il comando:\n{' '.join(command)}")

        try:
            # Esecuzione del comando in un processo separato
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8')

            # Legge l'output in tempo reale
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    self.log(output.strip())

            rc = process.poll()
            if rc == 0:
                self.log("\nGenerazione completata con successo!")
                self.log(f"Il pacchetto licenziato si trova in: {output_dir}")
                messagebox.showinfo("Successo", f"Pacchetto generato con successo in:\n{output_dir}")
            else:
                self.log(f"\nERRORE: Il processo è terminato con codice d'errore {rc}.")
                messagebox.showerror("Errore", f"La generazione è fallita. Controlla il log per i dettagli.")

        except Exception as e:
            self.log(f"Eccezione durante la generazione: {e}")
            messagebox.showerror("Errore Critico", f"Si è verificato un errore imprevisto:\n{e}")
        finally:
            self.generate_button.config(state=tk.NORMAL)


if __name__ == "__main__":
    app = LicenseGenerator()
    app.mainloop()