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
import urllib.request # Mantenuto per la generazione licenza, se necessario
# import zipfile # Non più necessario senza Python embeddable
import traceback # Importato per logging errori
import fnmatch # Importato per la copia degli asset

class ObfuscatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("General Obfuscator and License Manager")
        self.geometry("700x550")

        # Variabili per i percorsi e dati
        self.source_path = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.license_path = tk.StringVar()

        self.obfuscation_queue = queue.Queue()
        self.license_queue = queue.Queue()

        # Creazione del Notebook per le schede
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)

        # Creazione dei frame per le schede
        self.obfuscator_tab = ttk.Frame(self.notebook)
        self.license_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.obfuscator_tab, text='Obfuscator')
        self.notebook.add(self.license_tab, text='License Manager')

        self.create_obfuscator_tab()
        self.create_license_tab()

    def create_license_tab(self):
        # Variabili per la generazione licenza
        self.expiry_date = tk.StringVar()
        self.device_id = tk.StringVar()

        # Frame per i campi di input
        input_frame = tk.Frame(self.license_tab, pady=5)
        input_frame.pack(fill='x', padx=10, pady=(10,0))

        tk.Label(input_frame, text="Expiry Date (YYYY-MM-DD):").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.expiry_entry = tk.Entry(input_frame, textvariable=self.expiry_date, width=40)
        self.expiry_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)

        tk.Label(input_frame, text="Device ID:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.device_id_entry = tk.Entry(input_frame, textvariable=self.device_id, width=40)
        self.device_id_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        input_frame.columnconfigure(1, weight=1)

        # Pulsante di generazione
        self.generate_license_button = tk.Button(self.license_tab, text="Generate License Key", command=self.start_license_generation, font=('Helvetica', 12, 'bold'), pady=10)
        self.generate_license_button.pack(pady=10)

        # Area di stato per la licenza
        license_status_frame = tk.Frame(self.license_tab, pady=10)
        license_status_frame.pack(expand=True, fill='both', padx=10)
        tk.Label(license_status_frame, text="Status:").pack(anchor='w')
        self.license_status_text = tk.Text(license_status_frame, height=10, state='disabled', bg='black', fg='white', font=('Courier', 9))
        self.license_status_text.pack(expand=True, fill='both')

    def create_obfuscator_tab(self):
        # Frame per la selezione del percorso di origine
        source_frame = tk.Frame(self.obfuscator_tab, pady=5)
        source_frame.pack(fill='x', padx=10, pady=(10,0))
        tk.Label(source_frame, text="Source Folder:").pack(side='left')
        self.source_entry = tk.Entry(source_frame, textvariable=self.source_path, state='readonly', width=50)
        self.source_entry.pack(side='left', expand=True, fill='x', padx=5)
        self.browse_source_button = tk.Button(source_frame, text="Browse...", command=self.select_source)
        self.browse_source_button.pack(side='left')

        # Frame per la selezione del percorso di destinazione
        dest_frame = tk.Frame(self.obfuscator_tab, pady=5)
        dest_frame.pack(fill='x', padx=10)
        tk.Label(dest_frame, text="Destination Folder:").pack(side='left')
        self.dest_entry = tk.Entry(dest_frame, textvariable=self.destination_path, state='readonly', width=50)
        self.dest_entry.pack(side='left', expand=True, fill='x', padx=5)
        self.browse_dest_button = tk.Button(dest_frame, text="Browse...", command=self.select_destination)
        self.browse_dest_button.pack(side='left')

        # Frame per la selezione del file di licenza
        license_frame = tk.Frame(self.obfuscator_tab, pady=5)
        license_frame.pack(fill='x', padx=10)
        tk.Label(license_frame, text="License File (optional):").pack(side='left')
        self.license_entry = tk.Entry(license_frame, textvariable=self.license_path, state='readonly', width=50)
        self.license_entry.pack(side='left', expand=True, fill='x', padx=5)
        self.browse_license_button = tk.Button(license_frame, text="Browse...", command=self.select_license)
        self.browse_license_button.pack(side='left')

        # Pulsante di avvio
        self.start_button = tk.Button(self.obfuscator_tab, text="Start Obfuscation", command=self.start_obfuscation, state='disabled', font=('Helvetica', 12, 'bold'), pady=10)
        self.start_button.pack(pady=10)

        # Area di stato
        status_frame = tk.Frame(self.obfuscator_tab, pady=10)
        status_frame.pack(expand=True, fill='both', padx=10)
        tk.Label(status_frame, text="Status:").pack(anchor='w')
        self.obfuscation_status_text = tk.Text(status_frame, height=10, state='disabled', bg='black', fg='white', font=('Courier', 9))
        self.obfuscation_status_text.pack(expand=True, fill='both')

    def _update_status(self, message):
        self.obfuscation_status_text.config(state='normal')
        self.obfuscation_status_text.insert(tk.END, message)
        self.obfuscation_status_text.see(tk.END)
        self.obfuscation_status_text.config(state='disabled')

    def select_source(self):
        path = filedialog.askdirectory(title="Select Source Folder")
        if path:
            self.source_path.set(path)
            self._update_status(f"Source folder set to: {path}\n")
            self.check_paths()

    def select_destination(self):
        path = filedialog.askdirectory(title="Select Destination Folder")
        if path:
            self.destination_path.set(path)
            self._update_status(f"Destination folder set to: {path}\n")
            self.check_paths()

    def select_license(self):
        path = filedialog.askopenfilename(
            title="Select License File",
            filetypes=[("License Files", "*.lic *.rkey"), ("All files", "*.*")]
        )
        if path:
            self.license_path.set(path)
            self._update_status(f"License file set to: {path}\n")

    def check_paths(self):
        if self.source_path.get() and self.destination_path.get():
            self.start_button.config(state='normal')
        else:
            self.start_button.config(state='disabled')

    def start_obfuscation(self):
        source = self.source_path.get()
        dest = self.destination_path.get()
        license_f = self.license_path.get()

        if not source or not dest:
            messagebox.showerror("Error", "Please select both a source and destination folder.")
            return

        # Disable buttons
        self.start_button.config(state='disabled')
        self.browse_source_button.config(state='disabled')
        self.browse_dest_button.config(state='disabled')
        self.browse_license_button.config(state='disabled')

        # Clear status area
        self.obfuscation_status_text.config(state='normal')
        self.obfuscation_status_text.delete('1.0', tk.END)
        self.obfuscation_status_text.config(state='disabled')

        # Start background thread
        thread = threading.Thread(target=obfuscation_process, args=(source, dest, license_f, self.obfuscation_queue))
        thread.daemon = True
        thread.start()

        # Start processing queue
        self.process_obfuscation_queue()

    def process_obfuscation_queue(self):
        try:
            while True:
                message = self.obfuscation_queue.get_nowait()
                if isinstance(message, tuple) and message[0] == "PROCESS_COMPLETE":
                    # Re-enable buttons
                    self.start_button.config(state='normal')
                    self.browse_source_button.config(state='normal')
                    self.browse_dest_button.config(state='normal')
                    self.browse_license_button.config(state='normal')
                    self._update_status("\n--- Ready for next operation. ---\n")
                    break
                else:
                    self._update_status(message)
        except queue.Empty:
            self.after(100, self.process_obfuscation_queue)

    def run(self):
        self.mainloop()

    def start_license_generation(self):
        expiry = self.expiry_date.get()
        device_id = self.device_id.get()

        if not expiry or not device_id:
            messagebox.showerror("Error", "Please provide both an expiry date and a device ID.")
            return

        output_folder = filedialog.askdirectory(title="Select a folder to save the license key")
        if not output_folder:
            return # User cancelled

        self.generate_license_button.config(state='disabled')
        self.license_status_text.config(state='normal')
        self.license_status_text.delete('1.0', tk.END)
        self.license_status_text.config(state='disabled')

        thread = threading.Thread(target=license_generation_process, args=(expiry, device_id, output_folder, self.license_queue))
        thread.daemon = True
        thread.start()

        self.process_license_queue()

    def process_license_queue(self):
        try:
            while True:
                message = self.license_queue.get_nowait()
                if isinstance(message, tuple) and message[0] == "LICENSE_PROCESS_COMPLETE":
                    self.generate_license_button.config(state='normal')
                    self._update_license_status("\n--- Ready for next operation. ---\n")
                    break
                else:
                    self._update_license_status(message)
        except queue.Empty:
            self.after(100, self.process_license_queue)

    def _update_license_status(self, message):
        self.license_status_text.config(state='normal')
        self.license_status_text.insert(tk.END, message)
        self.license_status_text.see(tk.END)
        self.license_status_text.config(state='disabled')


def license_generation_process(expiry_date, device_id, output_folder, queue_obj):
    try:
        queue_obj.put("--- Starting License Generation ---\n")

        # Validate and format date
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", expiry_date):
                raise ValueError("Invalid date format. Please use YYYY-MM-DD.")

        queue_obj.put(f"Expiry: {expiry_date}, Device ID: {device_id}\n")

        # --- Attempt 1: Standard command ---
        cmd1_list = ["pyarmor", "gen", "key", "-O", output_folder, "-e", expiry_date, "-b", device_id]
        queue_obj.put(f"Executing: {' '.join(shlex.quote(arg) for arg in cmd1_list)}\n")
        proc1 = subprocess.run(cmd1_list, capture_output=True, text=True, encoding='utf-8', errors='ignore', check=False)

        success = False
        p = pathlib.Path(output_folder)
        license_files = list(p.glob("*.lic")) + list(p.glob("*.rkey"))
        if license_files:
            success = True
            queue_obj.put(f"Found license file: {license_files[0]}\n")

        if not success:
            queue_obj.put(f"Command output (stdout):\n{proc1.stdout}\n")
            queue_obj.put(f"Command output (stderr):\n{proc1.stderr}\n")
            queue_obj.put("First attempt failed. Retrying with 'disk:' prefix...\n")
            time.sleep(1)

            # --- Attempt 2 (Fallback): "disk:" prefix ---
            cmd2_list = ["pyarmor", "gen", "key", "-O", output_folder, "-e", expiry_date, "-b", f"disk:{device_id}"]
            queue_obj.put(f"Executing: {' '.join(shlex.quote(arg) for arg in cmd2_list)}\n")
            proc2 = subprocess.run(cmd2_list, capture_output=True, text=True, encoding='utf-8', errors='ignore', check=False)

            license_files = list(p.glob("*.lic")) + list(p.glob("*.rkey"))
            if license_files:
                success = True
                queue_obj.put(f"Found license file: {license_files[0]}\n")
            final_proc = proc2
        else:
            final_proc = proc1

        queue_obj.put(f"Final command output (stdout):\n{final_proc.stdout}\n")
        queue_obj.put(f"Final command output (stderr):\n{final_proc.stderr}\n")

        if success:
            queue_obj.put("\n--- SUCCESS! ---\n")
            queue_obj.put(f"License key successfully created in: {output_folder}\n")
        else:
            if final_proc.returncode != 0:
                error_details = final_proc.stderr if final_proc.stderr else final_proc.stdout
                raise RuntimeError(f"License generation command failed with exit code {final_proc.returncode}. Details: {error_details.strip()}")
            else:
                 raise RuntimeError(f"License generation command finished, but no .lic or .rkey file found in {output_folder}.")

    except Exception as e:
        queue_obj.put(f"\n--- AN ERROR OCCURRED ---\n")
        queue_obj.put(traceback.format_exc() + "\n")
        queue_obj.put(f"{str(e)}\n")
    finally:
        queue_obj.put(("LICENSE_PROCESS_COMPLETE",))


# --- Funzione Semplificata per Sistema con Python Installato (con DLL copiate in obfuscated) ---
def obfuscation_process(source_dir, dest_dir, license_path, queue_obj):
    """
    Funzione semplificata che offusca gli script, crea launcher .bat,
    e copia le DLL VCRuntime e Python Core in obfuscated, assumendo Python installato.
    """
    pyarmor_runtime_folder_name = None
    py_version_major_minor = ".".join(map(str, sys.version_info[:2])) # e.g., "3.10"
    python_dll_name = f'python{py_version_major_minor.replace(".", "")}.dll' # e.g., python310.dll
    try:
        obfuscated_dir = os.path.join(dest_dir, "obfuscated")

        queue_obj.put(f"--- Starting Obfuscation Process (System Python Mode + DLL Copy {py_version_major_minor}) ---\n")

        # 1. Clean up and create directories
        if os.path.exists(dest_dir):
            queue_obj.put(f"Removing existing destination directory: {dest_dir}\n")
            if not dest_dir or len(dest_dir) < 5 or ":" not in dest_dir:
                 raise ValueError(f"Destination directory '{dest_dir}' seems unsafe to remove automatically.")
            shutil.rmtree(dest_dir)
        queue_obj.put(f"Creating destination directory: {dest_dir}\n")
        os.makedirs(obfuscated_dir)

        # 2. Find all Python scripts
        queue_obj.put(f"Searching for Python scripts in: {source_dir}\n")
        all_scripts = glob.glob(os.path.join(source_dir, '*.py'))
        if not all_scripts:
            raise FileNotFoundError("No Python files found in the source directory.")
        script_basenames = [os.path.basename(s) for s in all_scripts]
        queue_obj.put(f"Found {len(script_basenames)} scripts: {', '.join(script_basenames)}\n")

        # 3. Run PyArmor obfuscation
        queue_obj.put("\n--- Running PyArmor ---\n")
        command = [
            "pyarmor", "gen",
            "--outer",
            "--output", obfuscated_dir,
        ] + all_scripts
        process = subprocess.Popen(
            command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='ignore'
        )
        pyarmor_output_lines = []
        for line in iter(process.stdout.readline, ''):
            queue_obj.put(line)
            pyarmor_output_lines.append(line)
            if "INFO     Copy PyArmor runtime files to" in line or "INFO     generate runtime files to" in line:
                 try:
                     path_part = line.split(" to ")[-1].strip()
                     folder_name = os.path.basename(path_part)
                     if folder_name.startswith("pyarmor_runtime_"):
                         pyarmor_runtime_folder_name = folder_name
                         queue_obj.put(f"Detected PyArmor runtime folder: {pyarmor_runtime_folder_name}\n")
                 except Exception as parse_err:
                     queue_obj.put(f"WARNING: Could not parse runtime folder name from line: '{line.strip()}'. Error: {parse_err}\n")
        process.stdout.close()
        return_code = process.wait()

        if not pyarmor_runtime_folder_name:
            runtime_folders = glob.glob(os.path.join(obfuscated_dir, "pyarmor_runtime_*"))
            if runtime_folders:
                 pyarmor_runtime_folder_name = os.path.basename(runtime_folders[0])
                 queue_obj.put(f"Manually found PyArmor runtime folder: {pyarmor_runtime_folder_name}\n")
            else:
                 default_runtime = os.path.join(obfuscated_dir, "pyarmor_runtime_000000")
                 if os.path.isdir(default_runtime):
                     pyarmor_runtime_folder_name = os.path.basename(default_runtime)
                     queue_obj.put(f"Assuming default PyArmor runtime folder: {pyarmor_runtime_folder_name}\n")
                 else:
                     queue_obj.put("ERROR: Could not determine PyArmor runtime folder name!\n")
                     raise RuntimeError("PyArmor runtime folder not found after obfuscation.")
        if return_code != 0:
            queue_obj.put(f"PyArmor process exited with code {return_code}. Check output above for errors.\n")
            raise subprocess.CalledProcessError(return_code, command, output="See status area for PyArmor output")
        queue_obj.put("--- PyArmor Finished Successfully ---\n\n")

        # 3.5 Copia DLL VCRuntime e Python Core in obfuscated_dir
        queue_obj.put(f"--- Copying Dependency DLLs ({python_dll_name}, vcruntime*) to obfuscated folder --- \n")
        host_python_dir = sys.prefix # Directory dell'interprete Python che esegue questo script
        dll_source_dirs = [host_python_dir]
        system32_path = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "System32")
        if system32_path not in dll_source_dirs:
            dll_source_dirs.append(system32_path)

        dlls_to_copy = [python_dll_name] + glob.glob(os.path.join(host_python_dir, "vcruntime140*.dll"))

        found_python_dll = False
        python_dll_src_path = None
        vc_runtime_dll_paths = []

        # Cerca le DLL nelle possibili directory sorgente
        for source_dir_path in dll_source_dirs:
            # Cerca pythonXX.dll
            potential_py_dll = os.path.join(source_dir_path, python_dll_name)
            if not found_python_dll and os.path.exists(potential_py_dll):
                python_dll_src_path = potential_py_dll
                found_python_dll = True
                queue_obj.put(f"Found {python_dll_name} in: {source_dir_path}\n")

            # Cerca vcruntime*.dll
            vc_dlls_in_dir = glob.glob(os.path.join(source_dir_path, "vcruntime140*.dll"))
            for vc_dll in vc_dlls_in_dir:
                if vc_dll not in vc_runtime_dll_paths:
                    vc_runtime_dll_paths.append(vc_dll)
                    queue_obj.put(f"Found {os.path.basename(vc_dll)} in: {source_dir_path}\n")

        # Copia le DLL trovate
        dlls_to_copy_paths = []
        if python_dll_src_path:
            dlls_to_copy_paths.append(python_dll_src_path)
        else:
             queue_obj.put(f"WARNING: Could not find {python_dll_name} in {dll_source_dirs}. The obfuscated app might fail.\n")

        dlls_to_copy_paths.extend(vc_runtime_dll_paths)

        if not dlls_to_copy_paths:
             queue_obj.put(f"WARNING: Could not find any required DLLs ({python_dll_name}, vcruntime*) to copy into obfuscated folder.\n")
        else:
            for dll_path in dlls_to_copy_paths:
                dll_name = os.path.basename(dll_path)
                dest_dll_path = os.path.join(obfuscated_dir, dll_name)
                queue_obj.put(f"Copying {dll_name} to {obfuscated_dir}\n")
                try:
                    if not os.path.exists(dest_dll_path):
                         shutil.copy(dll_path, obfuscated_dir)
                    else:
                         queue_obj.put(f"Skipping copy, {dll_name} already exists in {obfuscated_dir}.\n")
                except Exception as copy_err:
                     queue_obj.put(f"ERROR copying {dll_name} to {obfuscated_dir}: {copy_err}\n")

        queue_obj.put("--- Finished copying Dependency DLLs to obfuscated folder --- \n")


        # 4. Create .bat launchers (Simplified for System Python)
        queue_obj.put("--- Creating .bat Launchers ---\n")
        for script_path in all_scripts:
            script_name = os.path.basename(script_path)
            base_name = os.path.splitext(script_name)[0]
            bat_name = base_name + ".bat"
            bat_path = os.path.join(dest_dir, bat_name)

            # Simplified launcher content
            launcher_content = f'''@echo off
setlocal
REM Cambia la directory corrente alla cartella 'obfuscated' relativa a questo .bat
cd /d "%~dp0obfuscated"

REM Esegui lo script usando il python del sistema (deve essere nel PATH)
echo Running: python "{base_name}.py" %* from %CD%
python "{base_name}.py" %*

endlocal
pause
'''
            queue_obj.put(f"Creating launcher for {script_name}...\n")
            with open(bat_path, 'w', encoding='utf-8') as f:
                f.write(launcher_content)
        queue_obj.put("--- Launchers Created Successfully ---\n\n")

        # 5. Copy non-Python assets (CORRECTED)
        queue_obj.put("--- Copying Assets ---\n")
        ignore_func = shutil.ignore_patterns('*.py', '__pycache__')
        patterns_to_ignore = ('*.py', '__pycache__') # Keep patterns for file check

        if os.path.exists(source_dir):
            for item in os.listdir(source_dir):
                 s = os.path.join(source_dir, item)
                 d = os.path.join(dest_dir, item) # Assets go to the main dest_dir
                 if os.path.isdir(s):
                     should_ignore_dir = any(fnmatch.fnmatch(item, pat) for pat in patterns_to_ignore)
                     if not should_ignore_dir:
                         queue_obj.put(f"Copying directory: {item}\n")
                         shutil.copytree(s, d, ignore=ignore_func, dirs_exist_ok=True)
                 else: # It's a file
                     should_ignore_file = any(fnmatch.fnmatch(item, pat) for pat in patterns_to_ignore)
                     if not should_ignore_file:
                         queue_obj.put(f"Copying file: {item}\n")
                         shutil.copy2(s, d) # copy2 preserves metadata
        else:
             queue_obj.put(f"WARNING: Source directory {source_dir} not found for copying assets.\n")


        # 6. Copy license file if provided (Keep this)
        if license_path:
            if os.path.exists(license_path):
                queue_obj.put(f"Copying license file...\n")
                shutil.copy(license_path, dest_dir)
                shutil.copy(license_path, obfuscated_dir)
            else:
                queue_obj.put(f"Warning: License file not found at '{license_path}'\n")
        queue_obj.put("--- Asset Copying Finished ---\n\n")

        # 7. REMOVED - No portable Python setup needed

        queue_obj.put(f"====== OBFUSCATION COMPLETE (System Python Mode + DLL Copy {py_version_major_minor}) ======\n")
        queue_obj.put(f"Final application is ready in: {dest_dir}\n")
        queue_obj.put(f"NOTE: Ensure Python {py_version_major_minor} is installed and in the system PATH on the target machine.\n")

    except Exception as e:
        queue_obj.put(f"\n--- AN ERROR OCCURRED during obfuscation process ---\n")
        queue_obj.put(traceback.format_exc() + "\n")
        queue_obj.put(f"{str(e)}\n")
    finally:
        queue_obj.put(("PROCESS_COMPLETE",))


# --- Main execution ---
if __name__ == "__main__":
    app = ObfuscatorApp()
    app.run()