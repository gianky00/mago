import customtkinter as ctk
from database import Database
from tkinter import filedialog
import subprocess
import datetime
import re
import os
import sys
import time # Aggiunto per la pausa

class LicenseManagerApp(ctk.CTk):
    def __init__(self, db_connection):
        super().__init__()

        self.db = db_connection
        self.title("Gestionale Licenze PyArmor")
        self.geometry("800x600")
        self.user_data_map = {}
        self.selected_user_id = None
        self.output_folder = ""

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.tab_view = ctk.CTkTabview(self, width=780, height=580)
        self.tab_view.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.tab_view.add("Genera Licenza")
        self.tab_view.add("Gestione Utenze")
        self.tab_view.add("Storico Licenze")

        self.create_generate_license_tab(self.tab_view.tab("Genera Licenza"))
        self.create_manage_users_tab(self.tab_view.tab("Gestione Utenze"))
        self.create_license_history_tab(self.tab_view.tab("Storico Licenze"))
        
        self._refresh_user_dropdown()

    def create_generate_license_tab(self, tab):
        tab.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(tab, text="Seleziona Utenza:").grid(row=0, column=0, padx=20, pady=10, sticky="w")
        self.user_dropdown_var = ctk.StringVar(value="Nessun utente selezionato")
        self.user_dropdown = ctk.CTkOptionMenu(tab, variable=self.user_dropdown_var, values=[], command=self._on_user_selected)
        self.user_dropdown.grid(row=0, column=1, padx=20, pady=10, sticky="ew")

        ctk.CTkLabel(tab, text="ID Hardware (Sola Lettura):").grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.hwid_entry = ctk.CTkEntry(tab, state="readonly")
        self.hwid_entry.grid(row=1, column=1, padx=20, pady=10, sticky="ew")

        ctk.CTkLabel(tab, text="Data di Scadenza (GG/MM/AAAA):").grid(row=2, column=0, padx=20, pady=10, sticky="w")
        self.expiry_date_entry = ctk.CTkEntry(tab, placeholder_text="Es. 31/12/2024")
        self.expiry_date_entry.grid(row=2, column=1, padx=20, pady=10, sticky="ew")

        ctk.CTkLabel(tab, text="Cartella di Destinazione:").grid(row=3, column=0, padx=20, pady=10, sticky="w")
        self.folder_path_label = ctk.CTkLabel(tab, text="Nessuna cartella selezionata", anchor="w")
        self.folder_path_label.grid(row=3, column=1, padx=20, pady=10, sticky="ew")
        ctk.CTkButton(tab, text="Seleziona Cartella...", command=self._select_output_folder).grid(row=4, column=1, padx=20, pady=10, sticky="e")

        self.generate_button = ctk.CTkButton(tab, text="Genera e Registra Licenza", command=self._generate_license)
        self.generate_button.grid(row=5, column=0, columnspan=2, padx=20, pady=20)

        self.status_label = ctk.CTkLabel(tab, text="", text_color="green", anchor="w")
        self.status_label.grid(row=6, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

    def _refresh_user_dropdown(self):
        self.user_data_map.clear()
        users = self.db.get_all_users()
        user_names = ["Nessun utente selezionato"]
        
        if users:
            for user_id, name, hwid in users:
                self.user_data_map[name] = (user_id, hwid)
                user_names.append(name)
        
        self.user_dropdown.configure(values=user_names)
        self.user_dropdown_var.set("Nessun utente selezionato")
        self._clear_license_fields()

    def _on_user_selected(self, selected_name):
        if selected_name in self.user_data_map:
            user_id, hwid = self.user_data_map[selected_name]
            self.selected_user_id = user_id
            
            self.hwid_entry.configure(state="normal")
            self.hwid_entry.delete(0, "end")
            self.hwid_entry.insert(0, hwid)
            self.hwid_entry.configure(state="readonly")
        else:
            self._clear_license_fields()
    
    def _clear_license_fields(self):
        self.selected_user_id = None
        self.hwid_entry.configure(state="normal")
        self.hwid_entry.delete(0, "end")
        self.hwid_entry.configure(state="readonly")

    def _select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder
            self.folder_path_label.configure(text=folder)

    def _generate_license(self):
        scripts_dir = os.path.dirname(sys.executable)
        pyarmor_cmd = os.path.join(scripts_dir, 'pyarmor-7.exe')

        if not os.path.exists(pyarmor_cmd):
            pyarmor_cmd = os.path.join(scripts_dir, 'pyarmor-7')
            if not os.path.exists(pyarmor_cmd):
                self.status_label.configure(text="Errore: pyarmor-7 non trovato in Scripts del venv.", text_color="red")
                return

        if self.selected_user_id is None:
            self.status_label.configure(text="Errore: Selezionare un utente.", text_color="red")
            return

        expiry_date_str = self.expiry_date_entry.get()
        if not re.match(r"^\d{2}/\d{2}/\d{4}$", expiry_date_str):
            self.status_label.configure(text="Errore: Formato data non valido.", text_color="red")
            return
            
        if not self.output_folder:
            self.status_label.configure(text="Errore: Selezionare una cartella di destinazione.", text_color="red")
            return

        try:
            date_obj = datetime.datetime.strptime(expiry_date_str, "%d/%m/%Y")
            formatted_date = date_obj.strftime("%Y-%m-%d")
        except ValueError:
            self.status_label.configure(text="Errore: Data non valida.", text_color="red")
            return

        hwid = self.hwid_entry.get()
        regfile_path = "C:\\Users\\Coemi\\Desktop\\SCRIPT\\pyarmor-regfile-9329.zip"

        try:
            self.status_label.configure(text="Step 1/2: Registrazione file licenza...", text_color="orange")
            self.update_idletasks()
            
            reg_command = [pyarmor_cmd, "register", regfile_path]
            subprocess.run(reg_command, capture_output=True, text=True, check=True, encoding='utf-8')

            self.status_label.configure(text="Step 2/2: Generazione del file license.lic...", text_color="orange")
            self.update_idletasks()
            
            gen_command = [
                pyarmor_cmd, "licenses",
                "--expired", formatted_date,
                "--bind-disk", hwid,
                "--output", self.output_folder
            ]

            result = subprocess.run(gen_command, capture_output=True, text=True, check=True, encoding='utf-8')
            
            # --- MESSA A PUNTO FINALE ---
            # Aggiunge una piccola pausa per dare tempo al file system di aggiornarsi
            time.sleep(0.5)

            expected_file_path = os.path.join(self.output_folder, "license.lic")
            if os.path.exists(expected_file_path):
                success, msg = self.db.add_license_record(self.selected_user_id, expiry_date_str)
                if success:
                    self.status_label.configure(text=f"Successo! File license.lic generato in '{self.output_folder}' e registrato.", text_color="green")
                else:
                    self.status_label.configure(text=f"Licenza generata, ma fallita registrazione DB: {msg}", text_color="orange")
            else:
                self.status_label.configure(text=f"Errore: PyArmor non ha creato il file. Output: {result.stdout}", text_color="red")

        except subprocess.CalledProcessError as e:
            error_output = e.stderr if e.stderr else e.stdout
            self.status_label.configure(text=f"Errore da PyArmor: {error_output.strip()}", text_color="red")
        except Exception as e:
            self.status_label.configure(text=f"Errore sconosciuto: {e}", text_color="red")
    
    # ... il resto del codice rimane invariato ...

    def create_manage_users_tab(self, tab):
        self.selected_user_for_edit = ctk.StringVar()
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        button_frame = ctk.CTkFrame(tab)
        button_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(button_frame, text="Aggiungi Utente", command=self._open_add_edit_user_popup).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Modifica Utente Selezionato", command=lambda: self._open_add_edit_user_popup(edit_mode=True)).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Elimina Utente Selezionato", command=self._delete_selected_user).pack(side="left", padx=5)
        self.user_list_frame = ctk.CTkScrollableFrame(tab, label_text="Utenti Registrati")
        self.user_list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.user_list_frame.grid_columnconfigure(0, weight=1)
        self.user_status_label = ctk.CTkLabel(tab, text="")
        self.user_status_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self._refresh_user_list()

    def _refresh_user_list(self):
        for widget in self.user_list_frame.winfo_children():
            widget.destroy()
        all_users = self.db.get_all_users()
        header_frame = ctk.CTkFrame(self.user_list_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 5))
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(header_frame, text="Seleziona", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10)
        ctk.CTkLabel(header_frame, text="Nome Utente", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=5, sticky="w")
        ctk.CTkLabel(header_frame, text="ID Scheda Madre", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=5, sticky="w")
        if not all_users:
            ctk.CTkLabel(self.user_list_frame, text="Nessun utente registrato.").pack(pady=10)
            return
        for user_id, name, hwid in all_users:
            user_frame = ctk.CTkFrame(self.user_list_frame)
            user_frame.pack(fill="x", pady=2, padx=5)
            user_frame.grid_columnconfigure(1, weight=1)
            user_frame.grid_columnconfigure(2, weight=1)
            radio_button = ctk.CTkRadioButton(user_frame, text="", variable=self.selected_user_for_edit, value=str(user_id))
            radio_button.grid(row=0, column=0, padx=10)
            ctk.CTkLabel(user_frame, text=name, anchor="w").grid(row=0, column=1, padx=5, sticky="ew")
            ctk.CTkLabel(user_frame, text=hwid, anchor="w").grid(row=0, column=2, padx=5, sticky="ew")

    def _open_add_edit_user_popup(self, edit_mode=False):
        user_id_to_edit = self.selected_user_for_edit.get()
        if edit_mode and not user_id_to_edit:
            self.user_status_label.configure(text="Errore: Nessun utente selezionato per la modifica.", text_color="red")
            return
        popup = ctk.CTkToplevel(self)
        popup.transient(self)
        popup.grab_set()
        if edit_mode:
            popup.title("Modifica Utente")
            user_data = next((u for u in self.db.get_all_users() if str(u[0]) == user_id_to_edit), None)
            if not user_data:
                self.user_status_label.configure(text="Errore: Utente non trovato.", text_color="red")
                popup.destroy()
                return
        else:
            popup.title("Aggiungi Nuovo Utente")
            user_data = None
        ctk.CTkLabel(popup, text="Nome Utente:").pack(pady=(10, 0))
        name_entry = ctk.CTkEntry(popup, width=300)
        name_entry.pack(pady=5, padx=10)
        if user_data: name_entry.insert(0, user_data[1])
        ctk.CTkLabel(popup, text="ID Scheda Madre:").pack()
        hwid_entry = ctk.CTkEntry(popup, width=300)
        hwid_entry.pack(pady=5, padx=10)
        if user_data: hwid_entry.insert(0, user_data[2])
        def save_action():
            name = name_entry.get().strip()
            hwid = hwid_entry.get().strip()
            if not name or not hwid: return
            if edit_mode: success, msg = self.db.update_user(user_id_to_edit, name, hwid)
            else: success, msg = self.db.add_user(name, hwid)
            if success:
                self.user_status_label.configure(text=msg, text_color="green")
                self._refresh_all_user_views()
                popup.destroy()
            else:
                error_label = ctk.CTkLabel(popup, text=msg, text_color="red")
                error_label.pack(pady=5)
        ctk.CTkButton(popup, text="Salva", command=save_action).pack(pady=10)

    def _delete_selected_user(self):
        user_id_to_delete = self.selected_user_for_edit.get()
        if not user_id_to_delete:
            self.user_status_label.configure(text="Errore: Nessun utente selezionato.", text_color="red")
            return
        popup = ctk.CTkToplevel(self)
        popup.title("Conferma Eliminazione")
        popup.transient(self)
        popup.grab_set()
        user_name = next((u[1] for u in self.db.get_all_users() if str(u[0]) == user_id_to_delete), "N/A")
        ctk.CTkLabel(popup, text=f"Sei sicuro di voler eliminare l'utente '{user_name}'?").pack(padx=20, pady=10)
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=10)
        def confirm_action():
            success, msg = self.db.delete_user(user_id_to_delete)
            self.user_status_label.configure(text=msg, text_color="green" if success else "red")
            self._refresh_all_user_views()
            self.selected_user_for_edit.set("")
            popup.destroy()
        ctk.CTkButton(btn_frame, text="Sì, Elimina", command=confirm_action, fg_color="red").pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Annulla", command=popup.destroy).pack(side="left", padx=10)

    def create_license_history_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        controls_frame = ctk.CTkFrame(tab)
        controls_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(controls_frame, text="Aggiorna Storico", command=self._refresh_license_history).pack(side="left", padx=5)
        self.history_list_frame = ctk.CTkScrollableFrame(tab, label_text="Storico Licenze Generate")
        self.history_list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.history_list_frame.grid_columnconfigure(0, weight=1)
        self._refresh_license_history()

    def _refresh_license_history(self):
        for widget in self.history_list_frame.winfo_children():
            widget.destroy()
        history_data = self.db.get_license_history()
        header = ctk.CTkFrame(self.history_list_frame, fg_color=("gray85", "gray20"))
        header.pack(fill="x", pady=(0, 5), padx=5)
        header.grid_columnconfigure(0, weight=2)
        header.grid_columnconfigure(1, weight=1)
        header.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(header, text="Nome Utente", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(header, text="Data Scadenza", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(header, text="Data Generazione", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=5, pady=2, sticky="w")
        if not history_data:
            ctk.CTkLabel(self.history_list_frame, text="Nessuna licenza nello storico.").pack(pady=10)
            return
        for i, (user_name, expiry_date, generation_date) in enumerate(history_data):
            row_color = ("#f0f0f0", "#303030") if i % 2 == 0 else ("#e0e0e0", "#252525")
            row_frame = ctk.CTkFrame(self.history_list_frame, fg_color=row_color)
            row_frame.pack(fill="x", pady=1, padx=5)
            row_frame.grid_columnconfigure(0, weight=2)
            row_frame.grid_columnconfigure(1, weight=1)
            row_frame.grid_columnconfigure(2, weight=1)
            ctk.CTkLabel(row_frame, text=user_name, anchor="w").grid(row=0, column=0, padx=5, pady=2, sticky="w")
            ctk.CTkLabel(row_frame, text=expiry_date, anchor="w").grid(row=0, column=1, padx=5, pady=2, sticky="w")
            ctk.CTkLabel(row_frame, text=generation_date, anchor="w").grid(row=0, column=2, padx=5, pady=2, sticky="w")

    def _refresh_all_user_views(self):
        self._refresh_user_list()
        self._refresh_user_dropdown()
        self._refresh_license_history()

    def on_closing(self):
        self.db.close()
        self.destroy()

if __name__ == "__main__":
    db_conn = None
    try:
        db_conn = Database()
        app = LicenseManagerApp(db_conn)
        app.protocol("WM_DELETE_WINDOW", app.on_closing)
        app.mainloop()
    except Exception as e:
        print(f"Errore durante l'avvio dell'applicazione: {e}")
    finally:
        if db_conn:
            db_conn.close()
            print("Connessione al database chiusa.")