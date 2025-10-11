import tkinter as tk
from tkinter import messagebox
from config_ui import ConfigWindow
import json
import threading
import magoPyton as mpy

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automazione RPA")
        self.root.geometry("350x200")

        self.btn_configure = tk.Button(root, text="Configura", command=self.open_configuration, font=("Arial", 12))
        self.btn_configure.pack(pady=15, padx=20, fill=tk.X)

        self.btn_run = tk.Button(root, text="Avvia Automazione", command=self.run_automation_thread, font=("Arial", 12), bg="lightblue")
        self.btn_run.pack(pady=15, padx=20, fill=tk.X)

    def open_configuration(self):
        config_window = ConfigWindow(self.root)
        config_window.grab_set()

    def run_automation_thread(self):
        """Avvia l'automazione in un thread separato per non bloccare la GUI."""
        self.btn_run.config(state="disabled", text="Automazione in corso...")

        automation_thread = threading.Thread(target=self.execute_automation)
        automation_thread.daemon = True
        automation_thread.start()

        self.root.after(100, self.check_thread, automation_thread)

    def execute_automation(self):
        """Carica la config ed esegue lo script di automazione."""
        try:
            config = mpy.load_config()
            if not config:
                messagebox.showerror("Errore", "Impossibile caricare la configurazione. Eseguire prima la configurazione.")
                return

            mpy.run_automation(config)

            messagebox.showinfo("Completato", "Automazione terminata con successo!")

        except OSError as e:
            # Questo errore si verifica tipicamente su Windows quando si tenta di aprire un file
            # (es. con os.startfile) ma non c'è un'applicazione predefinita per quel tipo di file.
            # Esempio: tentare di aprire un .pdf senza un lettore PDF installato.
            error_message = (
                "Errore di associazione file.\n\n"
                "Windows non ha trovato un'applicazione associata per eseguire un'azione.\n\n"
                "Causa comune: Manca un programma predefinito (es. un lettore PDF per i file .pdf, "
                "o un browser per i link web).\n\n"
                "Soluzione: Assicurati che le applicazioni necessarie siano installate e impostate come predefinite.\n\n"
                f"Dettagli tecnici: {e}"
            )
            messagebox.showerror("Errore di Sistema", error_message)
        except Exception as e:
            messagebox.showerror("Errore Critico", f"Si è verificato un errore imprevisto durante l'automazione:\n{e}")

    def check_thread(self, thread):
        """Controlla lo stato del thread e riabilita il pulsante se ha finito."""
        if not thread.is_alive():
            self.btn_run.config(state="normal", text="Avvia Automazione")
        else:
            self.root.after(100, self.check_thread, thread)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()