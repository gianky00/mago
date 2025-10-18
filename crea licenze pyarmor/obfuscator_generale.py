import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import subprocess
import shutil
import os
import glob
import queue
import sys
import tempfile
import re
import datetime
import time
import pathlib
import shlex
import urllib.request
import zipfile
import traceback # Importato per logging errori
import fnmatch # Importato per la copia degli asset
import customtkinter as ctk
from database import Database

class ObfuscatorApp(ctk.CTk):
    def __init__(self, db_connection):
        super().__init__()
        self.db = db_connection
        self.title("General Obfuscator and License Manager")
        self.geometry("800x600")

        # Variabili per i percorsi e dati
        self.source_path = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.license_path = tk.StringVar()

        self.obfuscation_queue = queue.Queue()
        self.license_queue = queue.Queue()

        self.user_data_map = {}
        self.selected_license_id = tk.StringVar()

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Creazione del Notebook per le schede
        self.notebook = ctk.CTkTabview(self, width=780, height=580)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)

        # Creazione dei frame per le schede
        self.notebook.add('Obfuscator')
        self.notebook.add('License Manager')
        self.notebook.add("Gestione Utenze")
        self.notebook.add("Storico Licenze")

        self.obfuscator_tab = self.notebook.tab('Obfuscator')
        self.license_tab = self.notebook.tab('License Manager')
        self.user_management_tab = self.notebook.tab("Gestione Utenze")
        self.license_history_tab = self.notebook.tab("Storico Licenze")


        self.create_obfuscator_tab()
        self.create_license_tab()
        self.create_user_management_tab()
        self.create_license_history_tab()
        self._refresh_all_user_views()

    def create_license_tab(self):
        # Variabili per la generazione licenza
        self.expiry_date = tk.StringVar()
        self.device_id = tk.StringVar()
        self.selected_user_id_for_license = None

        input_frame = ctk.CTkFrame(self.license_tab, fg_color="transparent")
        input_frame.pack(fill='x', padx=20, pady=(20,10))
        input_frame.grid_columnconfigure(1, weight=1)

        # User Selection Dropdown
        ctk.CTkLabel(input_frame, text="Seleziona Utenza:").grid(row=0, column=0, sticky='w', padx=5, pady=10)
        self.license_user_dropdown_var = ctk.StringVar(value="Nessun utente selezionato")
        self.license_user_dropdown = ctk.CTkOptionMenu(input_frame, variable=self.license_user_dropdown_var, values=[], command=self._on_license_user_selected)
        self.license_user_dropdown.grid(row=0, column=1, columnspan=2, sticky='ew', padx=5, pady=10)

        # Expiry Date
        ctk.CTkLabel(input_frame, text="Data di Scadenza (YYYY-MM-DD):").grid(row=1, column=0, sticky='w', padx=5, pady=10)
        self.expiry_entry = ctk.CTkEntry(input_frame, textvariable=self.expiry_date)
        self.expiry_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=10)
        self.two_months_button = ctk.CTkButton(input_frame, text="Scadenza 2 Mesi", command=self._set_two_months_expiry, width=120)
        self.two_months_button.grid(row=1, column=2, padx=(10, 5), pady=10)

        # Hardware ID
        ctk.CTkLabel(input_frame, text="Serial N. Disco rigido:").grid(row=2, column=0, sticky='w', padx=5, pady=10)
        self.device_id_entry = ctk.CTkEntry(input_frame, textvariable=self.device_id, state="readonly")
        self.device_id_entry.grid(row=2, column=1, columnspan=2, sticky='ew', padx=5, pady=10)

        self.generate_license_button = ctk.CTkButton(self.license_tab, text="Generate License Key", command=self.start_license_generation)
        self.generate_license_button.pack(pady=20, padx=20)

        license_status_frame = ctk.CTkFrame(self.license_tab, fg_color="transparent")
        license_status_frame.pack(expand=True, fill='both', padx=20, pady=10)
        ctk.CTkLabel(license_status_frame, text="Status:").pack(anchor='w')
        self.license_status_text = ctk.CTkTextbox(license_status_frame, state='disabled', fg_color="black", text_color="white")
        self.license_status_text.pack(expand=True, fill='both')

    def create_obfuscator_tab(self):
        source_frame = ctk.CTkFrame(self.obfuscator_tab, fg_color="transparent")
        source_frame.pack(fill='x', padx=20, pady=(20, 10))
        source_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(source_frame, text="Source Folder:").grid(row=0, column=0, sticky='w', padx=5)
        self.source_entry = ctk.CTkEntry(source_frame, textvariable=self.source_path, state='readonly')
        self.source_entry.grid(row=0, column=1, sticky='ew', padx=5)
        self.browse_source_button = ctk.CTkButton(source_frame, text="Browse...", command=self.select_source, width=80)
        self.browse_source_button.grid(row=0, column=2, padx=(5,0))

        dest_frame = ctk.CTkFrame(self.obfuscator_tab, fg_color="transparent")
        dest_frame.pack(fill='x', padx=20, pady=10)
        dest_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(dest_frame, text="Destination Folder:").grid(row=0, column=0, sticky='w', padx=5)
        self.dest_entry = ctk.CTkEntry(dest_frame, textvariable=self.destination_path, state='readonly')
        self.dest_entry.grid(row=0, column=1, sticky='ew', padx=5)
        self.browse_dest_button = ctk.CTkButton(dest_frame, text="Browse...", command=self.select_destination, width=80)
        self.browse_dest_button.grid(row=0, column=2, padx=(5,0))

        license_frame = ctk.CTkFrame(self.obfuscator_tab, fg_color="transparent")
        license_frame.pack(fill='x', padx=20, pady=10)
        license_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(license_frame, text="License File (optional):").grid(row=0, column=0, sticky='w', padx=5)
        self.license_entry = ctk.CTkEntry(license_frame, textvariable=self.license_path, state='readonly')
        self.license_entry.grid(row=0, column=1, sticky='ew', padx=5)
        self.browse_license_button = ctk.CTkButton(license_frame, text="Browse...", command=self.select_license, width=80)
        self.browse_license_button.grid(row=0, column=2, padx=(5,0))

        self.start_button = ctk.CTkButton(self.obfuscator_tab, text="Start Obfuscation", command=self.start_obfuscation, state='disabled')
        self.start_button.pack(pady=20, padx=20)

        status_frame = ctk.CTkFrame(self.obfuscator_tab, fg_color="transparent")
        status_frame.pack(expand=True, fill='both', padx=20, pady=10)
        ctk.CTkLabel(status_frame, text="Status:").pack(anchor='w')
        self.obfuscation_status_text = ctk.CTkTextbox(status_frame, state='disabled', fg_color="black", text_color="white")
        self.obfuscation_status_text.pack(expand=True, fill='both')

    def _update_status(self, message):
        self.obfuscation_status_text.configure(state='normal')
        self.obfuscation_status_text.insert(tk.END, message)
        self.obfuscation_status_text.see(tk.END)
        self.obfuscation_status_text.configure(state='disabled')

    def select_source(self):
        path = filedialog.askdirectory(title="Select Source Folder")
        if path:
            self.source_path.set(path)
            self._update_status(f"Source folder set to: {path}\n")
            self.source_entry.configure(state="normal")
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, path)
            self.source_entry.configure(state="readonly")
            self.check_paths()

    def select_destination(self):
        path = filedialog.askdirectory(title="Select Destination Folder")
        if path:
            self.destination_path.set(path)
            self._update_status(f"Destination folder set to: {path}\n")
            self.dest_entry.configure(state="normal")
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, path)
            self.dest_entry.configure(state="readonly")
            self.check_paths()

    def select_license(self):
        path = filedialog.askopenfilename(
            title="Select License File",
            filetypes=[("License Files", "*.lic *.rkey"), ("All files", "*.*")]
        )
        if path:
            self.license_path.set(path)
            self._update_status(f"License file set to: {path}\n")
            self.license_entry.configure(state="normal")
            self.license_entry.delete(0, tk.END)
            self.license_entry.insert(0, path)
            self.license_entry.configure(state="readonly")

    def check_paths(self):
        if self.source_path.get() and self.destination_path.get():
            self.start_button.configure(state='normal')
        else:
            self.start_button.configure(state='disabled')

    def start_obfuscation(self):
        source = self.source_path.get()
        dest = self.destination_path.get()
        license_f = self.license_path.get()

        if not source or not dest:
            messagebox.showerror("Error", "Please select both a source and destination folder.")
            return

        self.start_button.configure(state='disabled')
        self.browse_source_button.configure(state='disabled')
        self.browse_dest_button.configure(state='disabled')
        self.browse_license_button.configure(state='disabled')

        self.obfuscation_status_text.configure(state='normal')
        self.obfuscation_status_text.delete('1.0', tk.END)
        self.obfuscation_status_text.configure(state='disabled')

        thread = threading.Thread(target=obfuscation_process, args=(source, dest, license_f, self.obfuscation_queue))
        thread.daemon = True
        thread.start()
        self.process_obfuscation_queue()

    def process_obfuscation_queue(self):
        try:
            while True:
                message = self.obfuscation_queue.get_nowait()
                if isinstance(message, tuple) and message[0] == "PROCESS_COMPLETE":
                    self.start_button.configure(state='normal')
                    self.browse_source_button.configure(state='normal')
                    self.browse_dest_button.configure(state='normal')
                    self.browse_license_button.configure(state='normal')
                    self._update_status("\n--- Ready for next operation. ---\n")
                    break
                else:
                    self._update_status(message)
        except queue.Empty:
            self.after(100, self.process_obfuscation_queue)

    def on_closing(self):
        self.db.close()
        self.destroy()

    def start_license_generation(self):
        expiry = self.expiry_date.get()
        device_id = self.device_id.get()

        if not self.selected_user_id_for_license:
            messagebox.showerror("Error", "Please select a user.")
            return

        if not expiry or not device_id:
            messagebox.showerror("Error", "Please provide both an expiry date and a device ID.")
            return

        output_folder = filedialog.askdirectory(title="Select a folder to save the license key")
        if not output_folder:
            return

        self.generate_license_button.configure(state='disabled')
        self.license_status_text.configure(state='normal')
        self.license_status_text.delete('1.0', tk.END)
        self.license_status_text.configure(state='disabled')

        # Pass the user ID to the generation process
        thread = threading.Thread(target=self.license_generation_process, args=(expiry, device_id, output_folder, self.selected_user_id_for_license, self.license_queue))
        thread.daemon = True
        thread.start()
        self.process_license_queue()

    def process_license_queue(self):
        try:
            while True:
                message = self.license_queue.get_nowait()
                if isinstance(message, tuple):
                    if message[0] == "ADD_LICENSE_RECORD":
                        user_id, expiry_date = message[1], message[2]
                        db_success, db_msg = self.db.add_license_record(user_id, expiry_date)
                        if db_success:
                            self._update_license_status("Successfully recorded in history.\n")
                        else:
                            self._update_license_status(f"WARNING: Failed to record in history: {db_msg}\n")
                    elif message[0] == "LICENSE_PROCESS_COMPLETE":
                        self.generate_license_button.configure(state='normal')
                        self._update_license_status("\n--- Ready for next operation. ---\n")
                        self._refresh_license_history()  # Refresh history after generation
                        break
                else:
                    self._update_license_status(message)
        except queue.Empty:
            self.after(100, self.process_license_queue)

    def _update_license_status(self, message):
        self.license_status_text.configure(state='normal')
        self.license_status_text.insert(tk.END, message)
        self.license_status_text.see(tk.END)
        self.license_status_text.configure(state='disabled')

    def create_license_history_tab(self):
        self.license_history_tab.grid_columnconfigure(0, weight=1)
        self.license_history_tab.grid_rowconfigure(1, weight=1)

        controls_frame = ctk.CTkFrame(self.license_history_tab)
        controls_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.user_filter_var = ctk.StringVar(value="All Users")
        self.user_filter_menu = ctk.CTkOptionMenu(controls_frame, variable=self.user_filter_var, values=["All Users"], command=self._on_user_filter_selected)
        self.user_filter_menu.pack(side="left", padx=5, pady=5)

        ctk.CTkButton(controls_frame, text="Refresh History", command=self._refresh_license_history).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(controls_frame, text="Delete Selected", command=self._delete_selected_license, fg_color="red").pack(side="right", padx=5, pady=5)

        self.history_status_label = ctk.CTkLabel(self.license_history_tab, text="", anchor="w")
        self.history_status_label.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

        self.history_list_frame = ctk.CTkScrollableFrame(self.license_history_tab, label_text="Generated License History")
        self.history_list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.history_list_frame.grid_columnconfigure(0, weight=1)
        self._refresh_license_history()

    def _refresh_all_user_views(self):
        self._refresh_user_dropdowns()
        self._refresh_user_list()
        self._refresh_license_history()

    def _refresh_user_dropdowns(self):
        self.user_data_map.clear()
        users = self.db.get_all_users()
        user_names = ["Nessun utente selezionato"]
        if users:
            for user_id, name, hwid in users:
                self.user_data_map[name] = (user_id, hwid)
                user_names.append(name)

        # Update dropdowns in both tabs
        self.user_filter_menu.configure(values=["All Users"] + user_names[1:]) # History tab
        self.license_user_dropdown.configure(values=user_names) # License tab
        self.license_user_dropdown_var.set("Nessun utente selezionato")
        self._clear_license_fields()

    def _on_license_user_selected(self, selected_name):
        if selected_name in self.user_data_map:
            user_id, hwid = self.user_data_map[selected_name]
            self.selected_user_id_for_license = user_id
            self.device_id.set(hwid)
        else:
            self._clear_license_fields()

    def _clear_license_fields(self):
        self.selected_user_id_for_license = None
        self.device_id.set("")
        self.expiry_date.set("")

    def _set_two_months_expiry(self):
        from datetime import datetime, timedelta
        from dateutil.relativedelta import relativedelta

        future_date = datetime.now() + relativedelta(months=2)
        self.expiry_date.set(future_date.strftime("%Y-%m-%d"))

    def create_user_management_tab(self):
        self.selected_user_for_edit = tk.StringVar()
        self.user_management_tab.grid_columnconfigure(0, weight=1)
        self.user_management_tab.grid_rowconfigure(0, weight=1)

        button_frame = ctk.CTkFrame(self.user_management_tab)
        button_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        ctk.CTkButton(button_frame, text="Aggiungi Utente", command=self._open_add_edit_user_popup).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Modifica Utente Selezionato", command=lambda: self._open_add_edit_user_popup(edit_mode=True)).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Elimina Utente Selezionato", command=self._delete_selected_user, fg_color="red").pack(side="left", padx=5)

        self.user_list_frame = ctk.CTkScrollableFrame(self.user_management_tab, label_text="Utenti Registrati")
        self.user_list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.user_list_frame.grid_columnconfigure(0, weight=1)

        self.user_status_label = ctk.CTkLabel(self.user_management_tab, text="", anchor="w")
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
        ctk.CTkLabel(header_frame, text="Serial N. Disco rigido", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=5, sticky="w")

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

        ctk.CTkLabel(popup, text="Serial N. Disco rigido:").pack()
        hwid_entry = ctk.CTkEntry(popup, width=300)
        hwid_entry.pack(pady=5, padx=10)
        if user_data: hwid_entry.insert(0, user_data[2])

        def save_action():
            name = name_entry.get().strip()
            hwid = hwid_entry.get().strip()
            if not name or not hwid: return

            if edit_mode:
                success, msg = self.db.update_user(user_id_to_edit, name, hwid)
            else:
                success, msg = self.db.add_user(name, hwid)

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

        dialog = ctk.CTkInputDialog(text="Sei sicuro di voler eliminare questo utente?\nScrivi 'DELETE' per confermare.", title="Conferma Eliminazione")
        confirmation = dialog.get_input()

        if confirmation == "DELETE":
            success, msg = self.db.delete_user(user_id_to_delete)
            self.user_status_label.configure(text=msg, text_color="green" if success else "red")
            self._refresh_all_user_views()
            self.selected_user_for_edit.set("")
        else:
            self.user_status_label.configure(text="Eliminazione annullata.", text_color="orange")

    def _on_user_filter_selected(self, selected_name):
        self._refresh_license_history()

    def _refresh_license_history(self):
        for widget in self.history_list_frame.winfo_children():
            widget.destroy()

        selected_user_name = self.user_filter_var.get()
        history_data = []
        if selected_user_name == "All Users":
            history_data = self.db.get_license_history()
        elif selected_user_name in self.user_data_map:
            user_id, _ = self.user_data_map[selected_user_name]
            history_data = self.db.get_license_history_by_user(user_id)

        header = ctk.CTkFrame(self.history_list_frame, fg_color=("gray85", "gray20"))
        header.pack(fill="x", pady=(0, 5), padx=5)
        header.grid_columnconfigure(0, weight=0) # Radio button
        header.grid_columnconfigure(1, weight=2) # User
        header.grid_columnconfigure(2, weight=1) # Expiry
        header.grid_columnconfigure(3, weight=1) # Generated

        ctk.CTkLabel(header, text="", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, pady=2)
        ctk.CTkLabel(header, text="Username", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(header, text="Expiry Date", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(header, text="Generation Date", font=ctk.CTkFont(weight="bold")).grid(row=0, column=3, padx=5, pady=2, sticky="w")

        if not history_data:
            ctk.CTkLabel(self.history_list_frame, text="No licenses found in history.").pack(pady=10)
            return

        for i, (license_id, user_name, expiry_date, generation_date) in enumerate(history_data):
            row_color = ("#f0f0f0", "#303030") if i % 2 == 0 else ("#e0e0e0", "#252525")
            row_frame = ctk.CTkFrame(self.history_list_frame, fg_color=row_color)
            row_frame.pack(fill="x", pady=1, padx=5)
            row_frame.grid_columnconfigure(1, weight=2)
            row_frame.grid_columnconfigure(2, weight=1)
            row_frame.grid_columnconfigure(3, weight=1)

            radio = ctk.CTkRadioButton(row_frame, text="", variable=self.selected_license_id, value=str(license_id))
            radio.grid(row=0, column=0, padx=5, pady=2)
            ctk.CTkLabel(row_frame, text=user_name, anchor="w").grid(row=0, column=1, padx=5, pady=2, sticky="w")
            ctk.CTkLabel(row_frame, text=expiry_date, anchor="w").grid(row=0, column=2, padx=5, pady=2, sticky="w")
            ctk.CTkLabel(row_frame, text=generation_date, anchor="w").grid(row=0, column=3, padx=5, pady=2, sticky="w")

    def _delete_selected_license(self):
        license_id_to_delete = self.selected_license_id.get()
        if not license_id_to_delete:
            self.history_status_label.configure(text="Error: No license selected to delete.", text_color="red")
            return

        # Confirmation Dialog
        dialog = ctk.CTkInputDialog(text="Are you sure you want to delete this license record?\nType 'DELETE' to confirm.", title="Confirm Deletion")
        confirmation = dialog.get_input()

        if confirmation == "DELETE":
            success, msg = self.db.delete_license_record(license_id_to_delete)
            if success:
                self.history_status_label.configure(text=msg, text_color="green")
                self.selected_license_id.set("") # Clear selection
                self._refresh_license_history()
            else:
                self.history_status_label.configure(text=msg, text_color="red")
        else:
            self.history_status_label.configure(text="Deletion cancelled.", text_color="orange")

    def license_generation_process(self, expiry_date, device_id, output_folder, user_id, queue_obj):
        try:
            queue_obj.put("--- Starting License Generation ---\n")
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", expiry_date):
                raise ValueError("Invalid date format. Please use YYYY-MM-DD.")
            queue_obj.put(f"Expiry: {expiry_date}, Device ID: {device_id}\n")

            cmd1_list = ["pyarmor", "gen", "key", "-O", output_folder, "-e", expiry_date, "-b", device_id]
            queue_obj.put(f"Executing: {' '.join(shlex.quote(arg) for arg in cmd1_list)}\n")
            proc1 = subprocess.run(cmd1_list, capture_output=True, text=True, encoding='utf-8', errors='ignore', check=False)

            success = False
            p = pathlib.Path(output_folder)
            if list(p.glob("*.lic")) or list(p.glob("*.rkey")):
                success = True

            if not success:
                queue_obj.put(f"STDOUT:\n{proc1.stdout}\nSTDERR:\n{proc1.stderr}\n")
                queue_obj.put("First attempt failed. Retrying with 'disk:' prefix...\n")
                time.sleep(1)
                cmd2_list = ["pyarmor", "gen", "key", "-O", output_folder, "-e", expiry_date, "-b", f"disk:{device_id}"]
                queue_obj.put(f"Executing: {' '.join(shlex.quote(arg) for arg in cmd2_list)}\n")
                proc2 = subprocess.run(cmd2_list, capture_output=True, text=True, encoding='utf-8', errors='ignore', check=False)
                if list(p.glob("*.lic")) or list(p.glob("*.rkey")):
                    success = True
                final_proc = proc2
            else:
                final_proc = proc1

            queue_obj.put(f"Final STDOUT:\n{final_proc.stdout}\n")
            queue_obj.put(f"Final STDERR:\n{final_proc.stderr}\n")

            if success:
                # Pass data back to the main thread to be added to the database
                queue_obj.put(("ADD_LICENSE_RECORD", user_id, expiry_date))
                queue_obj.put("\n--- SUCCESS! ---\nLicense key created. Recording to history...\n")
            else:
                error_details = final_proc.stderr if final_proc.stderr else final_proc.stdout
                raise RuntimeError(f"License generation failed. Details: {error_details.strip()}")

        except Exception as e:
            queue_obj.put(f"\n--- AN ERROR OCCURRED ---\n{traceback.format_exc()}\n{str(e)}\n")
        finally:
            queue_obj.put(("LICENSE_PROCESS_COMPLETE",))


def obfuscation_process(source_dir, dest_dir, license_path, queue_obj):
    # Definizione della versione di Python da usare
    PYTHON_VERSION = "3.10.11"
    PYTHON_DOWNLOAD_URL = f"https://www.python.org/ftp/python/{PYTHON_VERSION}/python-{PYTHON_VERSION}-embed-amd64.zip"
    PYTHON_DIR_NAME = "python-embed"

    try:
        dest_dir = os.path.normpath(dest_dir)
        source_dir = os.path.normpath(source_dir)
        obfuscated_dir = os.path.join(dest_dir, "obfuscated")
        python_embed_dir = os.path.join(dest_dir, PYTHON_DIR_NAME)

        queue_obj.put(f"--- Starting Obfuscation Process for Python {PYTHON_VERSION} ---\n")

        # 1. Pulizia e creazione directory
        if os.path.exists(dest_dir):
            queue_obj.put(f"Removing existing destination directory: {dest_dir}\n")
            shutil.rmtree(dest_dir)
        queue_obj.put(f"Creating fresh directories...\n")
        os.makedirs(obfuscated_dir)
        os.makedirs(python_embed_dir)

        # 2. Download e estrazione di Python embeddable
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = os.path.join(temp_dir, "python.zip")
            queue_obj.put(f"Downloading Python from: {PYTHON_DOWNLOAD_URL}\n")
            urllib.request.urlretrieve(PYTHON_DOWNLOAD_URL, zip_path)

            queue_obj.put(f"Extracting Python to: {python_embed_dir}\n")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(python_embed_dir)

        python_exe_path = os.path.join(python_embed_dir, "python.exe")
        if not os.path.exists(python_exe_path):
            raise FileNotFoundError("python.exe not found after extraction.")
        queue_obj.put("--- Python embeddable setup complete ---\n\n")

        # 3. Offuscamento con PyArmor
        queue_obj.put("\n--- Running PyArmor ---\n")
        all_scripts = glob.glob(os.path.join(source_dir, '*.py'))
        if not all_scripts:
            raise FileNotFoundError("No Python files found in the source directory.")

        command = ["pyarmor", "gen", "--outer", "--output", obfuscated_dir] + all_scripts
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='ignore')
        for line in iter(process.stdout.readline, ''):
            queue_obj.put(line)
        return_code = process.wait()
        if return_code != 0:
            raise subprocess.CalledProcessError(return_code, command, "PyArmor obfuscation failed.")
        queue_obj.put("--- PyArmor Finished Successfully ---\n\n")

        # 4. Creazione dei file .bat launcher
        queue_obj.put("--- Creating .bat Launchers ---\n")
        for script_path in all_scripts:
            script_name = os.path.basename(script_path)
            base_name = os.path.splitext(script_name)[0]
            bat_path = os.path.join(dest_dir, f"{base_name}.bat")

            # Il path relativo a python.exe dalla cartella del .bat
            relative_python_path = f"%~dp0{PYTHON_DIR_NAME}\\python.exe"

            launcher_content = f'''@echo off
setlocal
REM Cambia la directory corrente alla cartella 'obfuscated'
cd /d "%~dp0obfuscated"

REM Esegui lo script usando l'interprete Python embeddable
echo Running: "{relative_python_path}" "{script_name}" %*
"{relative_python_path}" "{script_name}" %*

endlocal
pause
'''
            with open(bat_path, 'w', encoding='utf-8') as f:
                f.write(launcher_content)
            queue_obj.put(f"Created launcher: {os.path.basename(bat_path)}\n")
        queue_obj.put("--- Launchers Created Successfully ---\n\n")

        # 5. Copia degli asset (tutto tranne i .py)
        queue_obj.put("--- Copying Assets to Destination ---\n")
        asset_dest = dest_dir # Copia gli asset nella root di destinazione
        for item in os.listdir(source_dir):
            s = os.path.join(source_dir, item)
            d = os.path.join(asset_dest, item)
            if not item.endswith('.py') and not item == '__pycache__':
                if os.path.isdir(s):
                    queue_obj.put(f"Copying directory: {item}\n")
                    shutil.copytree(s, d, dirs_exist_ok=True)
                else:
                    queue_obj.put(f"Copying file: {item}\n")
                    shutil.copy2(s, d) # copy2 preserva i metadati

        # 6. Copia del file di licenza (se fornito)
        if license_path and os.path.exists(license_path):
            queue_obj.put(f"Copying license file to {dest_dir} and {obfuscated_dir}\n")
            shutil.copy(license_path, dest_dir)
            shutil.copy(license_path, obfuscated_dir)

        queue_obj.put("\n====== OBFUSCATION AND PACKAGING COMPLETE ======\n")
        queue_obj.put(f"Self-contained application is ready in: {dest_dir}\n")

    except Exception as e:
        queue_obj.put(f"\n--- AN ERROR OCCURRED ---\n")
        queue_obj.put(traceback.format_exc())
        queue_obj.put(f"\nERROR: {str(e)}\n")
    finally:
        queue_obj.put(("PROCESS_COMPLETE",))

if __name__ == "__main__":
    db_conn = None
    try:
        db_conn = Database()
        app = ObfuscatorApp(db_conn)
        app.protocol("WM_DELETE_WINDOW", app.on_closing)
        app.mainloop()
    except Exception as e:
        print(f"Error during application startup: {e}")
    finally:
        if db_conn:
            db_conn.close()
            print("Database connection closed.")