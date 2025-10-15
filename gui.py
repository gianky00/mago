import tkinter as tk
from tkinter import ttk, messagebox
from config_ui import ConfigFrame  # Modificato da ConfigWindow a ConfigFrame
import json
import threading
import magoPyton as mpy
import sys
import io

class TextRedirector(io.StringIO):
    def __init__(self, widget):
        self.widget = widget

    def write(self, s):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, s)
        self.widget.see(tk.END)
        self.widget.configure(state='disabled')

    def flush(self):
        pass

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automazione RPA Mago")
        self.root.geometry("800x600")

        # Creazione del Notebook per le schede
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")

        # --- Scheda Avvio Automazione ---
        self.automation_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.automation_tab, text="Avvio Automazione")
        self.create_automation_tab()

        # --- Scheda Configurazione ---
        self.config_tab = ConfigFrame(self.notebook)
        self.notebook.add(self.config_tab, text="Configurazione")

    def create_automation_tab(self):
        # Frame per i controlli
        control_frame = ttk.Frame(self.automation_tab)
        control_frame.pack(pady=10, padx=10, fill='x')

        self.btn_run = ttk.Button(control_frame, text="Avvia Automazione", command=self.run_automation_thread, style="Accent.TButton")
        self.btn_run.pack(pady=10, padx=10, fill=tk.X)

        # Stile per il pulsante
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 12), padding=10)

        # Frame per il log
        log_frame = ttk.LabelFrame(self.automation_tab, text="Log in tempo reale")
        log_frame.pack(pady=10, padx=10, expand=True, fill="both")

        self.log_widget = tk.Text(log_frame, wrap='word', state='disabled', font=("Courier New", 9))
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_widget.yview)
        self.log_widget.config(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_widget.pack(expand=True, fill="both", padx=5, pady=5)

        # Reindirizza stdout e stderr al widget di log
        sys.stdout = TextRedirector(self.log_widget)
        sys.stderr = TextRedirector(self.log_widget)

        print("--- Log di Automazione RPA Mago ---")
        print("Benvenuto! Premi 'Avvia Automazione' per iniziare.")
        print("Vai alla scheda 'Configurazione' per modificare le impostazioni.")


    def run_automation_thread(self):
        """Avvia l'automazione in un thread separato per non bloccare la GUI."""
        self.btn_run.config(state="disabled")
        print("\n" + "="*50)
        print("Avvio dell'automazione in corso...")
        print("="*50 + "\n")

        automation_thread = threading.Thread(target=self.execute_automation)
        automation_thread.daemon = True
        automation_thread.start()

    def execute_automation(self):
        """Carica la config ed esegue lo script di automazione."""
        try:
            config = mpy.load_config()
            if not config:
                messagebox.showerror("Errore", "Impossibile caricare la configurazione. Eseguire prima la configurazione.")
                self.root.after(0, self.on_automation_complete, False)
                return

            mpy.run_automation(config)

            self.root.after(0, self.on_automation_complete, True)

        except OSError as e:
            error_message = (
                f"ERRORE DI SISTEMA: {e}\n\n"
                "Questo errore si verifica quando Windows non trova un'applicazione predefinita per un tipo di file.\n"
                "Causa comune: Manca un lettore PDF, un browser web o un'altra applicazione necessaria.\n"
                "Soluzione: Installa il software necessario e impostalo come predefinito nel sistema."
            )
            print(f"\nCRITICAL ERROR:\n{error_message}\n")
            messagebox.showerror("Errore di Sistema", error_message)
            self.root.after(0, self.on_automation_complete, False)
        except Exception as e:
            print(f"\nCRITICAL ERROR: Si è verificato un errore imprevisto durante l'automazione:\n{e}\n")
            messagebox.showerror("Errore Critico", f"Si è verificato un errore imprevisto:\n{e}\n\nControlla il log per i dettagli.")
            self.root.after(0, self.on_automation_complete, False)

    def on_automation_complete(self, success):
        """Funzione chiamata alla fine dell'automazione per aggiornare la GUI."""
        self.btn_run.config(state="normal")
        if success:
            print("\n" + "="*50)
            print("Automazione terminata con successo!")
            print("="*50)
            messagebox.showinfo("Completato", "Automazione terminata con successo!")
        else:
            print("\n" + "="*50)
            print("Automazione terminata con errori.")
            print("="*50)

if __name__ == "__main__":
    try:
        # This is the PyArmor runtime check
        from pytransform import pyarmor_runtime
        pyarmor_runtime()

        # Salva stdout e stderr originali per ripristinarli se necessario
        original_stdout = sys.stdout
        original_stderr = sys.stderr

        try:
            root = tk.Tk()
            app = App(root)
            root.mainloop()
        finally:
            # Ripristina stdout e stderr quando l'applicazione si chiude
            sys.stdout = original_stdout
            sys.stderr = original_stderr

    except ImportError:
        # This occurs if the pytransform module isn't found (e.g., dev environment)
        # You can add a simple Tkinter window here to show the error if needed.
        # For now, we just print and exit.
        print("Errore: Impossibile trovare i file di runtime di PyArmor.")
        # Optional: create a simple tkinter error message box
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Errore di Avvio", "Impossibile trovare i file di runtime necessari. L'applicazione non può partire.")
        sys.exit(1)

    except Exception as e:
        # This will catch PyArmor's license-related exceptions
        error_title = "Errore di Licenza"
        error_message = f"Si è verificato un errore critico relativo alla licenza e l'applicazione non può continuare.\n\nDettagli: {e}"

        # We need a separate, simple Tkinter instance to show the error,
        # as the main app loop may not have started.
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        messagebox.showerror(error_title, error_message)
        sys.exit(1)