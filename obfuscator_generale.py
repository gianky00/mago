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
        import re
        import datetime
        import time
        import pathlib
        import shlex

        queue_obj.put("--- Starting License Generation ---\n")

        # Validate and format date
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", expiry_date):
             raise ValueError("Invalid date format. Please use YYYY-MM-DD.")

        queue_obj.put(f"Expiry: {expiry_date}, Device ID: {device_id}\n")

        # --- Attempt 1: Standard command ---
        cmd1 = f'pyarmor gen key -O "{output_folder}" -e {expiry_date} -b "{device_id}"'
        queue_obj.put(f"Executing: {cmd1}\n")

        proc1 = subprocess.run(shlex.split(cmd1), capture_output=True, text=True, encoding='utf-8', errors='ignore')

        success = False
        p = pathlib.Path(output_folder)
        for ext in ("*.rkey", "*.lic"):
            if list(p.glob(ext)):
                success = True
                break

        if not success:
            queue_obj.put("First attempt failed. Retrying with 'disk:' prefix...\n")
            time.sleep(1)

            # --- Attempt 2 (Fallback): "disk:" prefix ---
            cmd2 = f'pyarmor gen key -O "{output_folder}" -e {expiry_date} -b "disk:{device_id}"'
            queue_obj.put(f"Executing: {cmd2}\n")

            proc2 = subprocess.run(shlex.split(cmd2), capture_output=True, text=True, encoding='utf-8', errors='ignore')

            for ext in ("*.rkey", "*.lic"):
                if list(p.glob(ext)):
                    success = True
                    break

            final_proc = proc2
        else:
            final_proc = proc1

        if success:
            queue_obj.put("\n--- SUCCESS! ---\n")
            queue_obj.put(f"License key successfully created in: {output_folder}\n")
        else:
            error_details = final_proc.stderr if final_proc.stderr else final_proc.stdout
            raise RuntimeError(f"License generation failed. Details: {error_details.strip()}")

    except Exception as e:
        queue_obj.put(f"\n--- AN ERROR OCCURRED ---\n")
        queue_obj.put(f"{str(e)}\n")
    finally:
        queue_obj.put(("LICENSE_PROCESS_COMPLETE",))


def obfuscation_process(source_dir, dest_dir, license_path, queue_obj):
    """
    Questa è la funzione principale che gestisce il processo di offuscamento.
    Verrà eseguita in un thread separato.
    """
    try:
        obfuscated_dir = os.path.join(dest_dir, "obfuscated")

        queue_obj.put("--- Starting Obfuscation Process ---\n")

        # 1. Clean up and create directories
        if os.path.exists(dest_dir):
            queue_obj.put(f"Removing existing destination directory: {dest_dir}\n")
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
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        for line in iter(process.stdout.readline, ''):
            queue_obj.put(line)
        process.stdout.close()
        return_code = process.wait()
        if return_code != 0:
            raise subprocess.CalledProcessError(return_code, command)
        queue_obj.put("--- PyArmor Finished Successfully ---\n\n")

        # 4. Create .bat launchers
        queue_obj.put("--- Creating .bat Launchers ---\n")
        for script_path in all_scripts:
            script_name = os.path.basename(script_path)
            bat_name = os.path.splitext(script_name)[0] + ".bat"
            bat_path = os.path.join(dest_dir, bat_name)

            relative_script_path = os.path.join("obfuscated", script_name)

            launcher_content = f'''@echo off
setlocal
cd /d %~dp0
.\\python.exe ".\\{relative_script_path}"
endlocal
pause
'''
            queue_obj.put(f"Creating launcher for {script_name}...\n")
            with open(bat_path, 'w', encoding='utf-8') as f:
                f.write(launcher_content)
        queue_obj.put("--- Launchers Created Successfully ---\n\n")

        # 5. Copy non-Python assets
        queue_obj.put("--- Copying Assets ---\n")
        for item in glob.glob(os.path.join(source_dir, '*')):
            if not item.endswith('.py'):
                dest_item = os.path.join(dest_dir, os.path.basename(item))
                if os.path.isdir(item):
                    queue_obj.put(f"Copying directory: {os.path.basename(item)}\n")
                    shutil.copytree(item, dest_item)
                else:
                    queue_obj.put(f"Copying file: {os.path.basename(item)}\n")
                    shutil.copy(item, dest_item)

        # 6. Copy license file if provided
        if license_path:
            if os.path.exists(license_path):
                queue_obj.put(f"Copying license file...\n")
                shutil.copy(license_path, dest_dir)
            else:
                queue_obj.put(f"Warning: License file not found at '{license_path}'\n")
        queue_obj.put("--- Asset Copying Finished ---\n\n")

        # 7. Create portable Python runtime by downloading the embeddable package
        queue_obj.put("--- Setting up Portable Python Runtime ---\n")

        # Determine Python version to download
        py_version = ".".join(map(str, sys.version_info[:3]))
        py_version_major_minor = ".".join(map(str, sys.version_info[:2]))
        queue_obj.put(f"Current Python version detected: {py_version}\n")

        # Construct download URL and cache path
        url = f"https://www.python.org/ftp/python/{py_version}/python-{py_version}-embed-amd64.zip"
        cache_dir = os.path.join(tempfile.gettempdir(), "py_embed_cache")
        os.makedirs(cache_dir, exist_ok=True)
        zip_path = os.path.join(cache_dir, os.path.basename(url))

        queue_obj.put(f"Download URL: {url}\n")
        queue_obj.put(f"Cache path: {zip_path}\n")

        # Download if not cached
        if not os.path.exists(zip_path):
            try:
                import urllib.request
                queue_obj.put(f"Downloading Python embeddable package...\n")

                # Simple progress bar logic
                def show_progress(block_num, block_size, total_size):
                    downloaded = block_num * block_size
                    if total_size > 0:
                        percent = downloaded * 100 / total_size
                        queue_obj.put(f"\r  -> {percent:.1f}% of {total_size/1024/1024:.2f} MB",)

                urllib.request.urlretrieve(url, zip_path, show_progress)
                queue_obj.put("\nDownload complete.\n")
            except Exception as e:
                raise RuntimeError(f"Failed to download Python package from {url}. Error: {e}")
        else:
            queue_obj.put("Python package found in cache.\n")

        # Extract the zip file to the destination directory
        try:
            import zipfile
            queue_obj.put(f"Extracting Python runtime to {dest_dir}...\n")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(dest_dir)
            queue_obj.put("Extraction complete.\n")

            # --- Install Tkinter ---
            queue_obj.put("Installing Tkinter for the new runtime...\n")

            # 1. Find the ._pth file to enable site-packages
            pth_files = glob.glob(os.path.join(dest_dir, 'python*._pth'))
            if not pth_files:
                raise FileNotFoundError("Could not find the ._pth file in the extracted Python runtime.")

            with open(pth_files[0], 'a') as f:
                f.write("\nimport site\n")

            # 2. Get pip
            get_pip_url = "https://bootstrap.pypa.io/get-pip.py"
            get_pip_path = os.path.join(cache_dir, "get-pip.py")
            if not os.path.exists(get_pip_path):
                urllib.request.urlretrieve(get_pip_url, get_pip_path)

            # 3. Run get-pip.py with the new python.exe
            new_python_exe = os.path.join(dest_dir, "python.exe")
            subprocess.check_call([new_python_exe, get_pip_path])

            # 4. Use the new pip to install tk
            new_pip_exe = os.path.join(dest_dir, "Scripts", "pip.exe")
            subprocess.check_call([new_pip_exe, "install", "tk"])

            queue_obj.put("Tkinter installed successfully.\n")

        except Exception as e:
            raise RuntimeError(f"Failed to setup Python runtime. Error: {e}")
        queue_obj.put("====== OBFUSCATION COMPLETE ======\n")
        queue_obj.put(f"Final application is ready in: {dest_dir}\n")

    except Exception as e:
        queue_obj.put(f"\n--- AN ERROR OCCURRED ---\n")
        queue_obj.put(f"{str(e)}\n")
    finally:
        queue_obj.put(("PROCESS_COMPLETE",))

if __name__ == "__main__":
    app = ObfuscatorApp()
    app.run()
