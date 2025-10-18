import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import subprocess
import shutil
import os
import glob
import queue
import sys

class ObfuscatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("General Obfuscator")
        self.geometry("600x450")

        # Variabili per i percorsi
        self.source_path = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.license_path = tk.StringVar()
        self.queue = queue.Queue()

        # Frame per la selezione del percorso di origine
        source_frame = tk.Frame(self, pady=5)
        source_frame.pack(fill='x', padx=10, pady=(10,0))
        tk.Label(source_frame, text="Source Folder:").pack(side='left')
        self.source_entry = tk.Entry(source_frame, textvariable=self.source_path, state='readonly', width=50)
        self.source_entry.pack(side='left', expand=True, fill='x', padx=5)
        self.browse_source_button = tk.Button(source_frame, text="Browse...", command=self.select_source)
        self.browse_source_button.pack(side='left')

        # Frame per la selezione del percorso di destinazione
        dest_frame = tk.Frame(self, pady=5)
        dest_frame.pack(fill='x', padx=10)
        tk.Label(dest_frame, text="Destination Folder:").pack(side='left')
        self.dest_entry = tk.Entry(dest_frame, textvariable=self.destination_path, state='readonly', width=50)
        self.dest_entry.pack(side='left', expand=True, fill='x', padx=5)
        self.browse_dest_button = tk.Button(dest_frame, text="Browse...", command=self.select_destination)
        self.browse_dest_button.pack(side='left')

        # Frame per la selezione del file di licenza
        license_frame = tk.Frame(self, pady=5)
        license_frame.pack(fill='x', padx=10)
        tk.Label(license_frame, text="License File (optional):").pack(side='left')
        self.license_entry = tk.Entry(license_frame, textvariable=self.license_path, state='readonly', width=50)
        self.license_entry.pack(side='left', expand=True, fill='x', padx=5)
        self.browse_license_button = tk.Button(license_frame, text="Browse...", command=self.select_license)
        self.browse_license_button.pack(side='left')

        # Pulsante di avvio
        self.start_button = tk.Button(self, text="Start Obfuscation", command=self.start_obfuscation, state='disabled', font=('Helvetica', 12, 'bold'), pady=10)
        self.start_button.pack(pady=10)

        # Area di stato
        status_frame = tk.Frame(self, pady=10)
        status_frame.pack(expand=True, fill='both', padx=10)
        tk.Label(status_frame, text="Status:").pack(anchor='w')
        self.status_text = tk.Text(status_frame, height=10, state='disabled', bg='black', fg='white', font=('Courier', 9))
        self.status_text.pack(expand=True, fill='both')

    def _update_status(self, message):
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message)
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

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
        self.status_text.config(state='normal')
        self.status_text.delete('1.0', tk.END)
        self.status_text.config(state='disabled')

        # Start background thread
        thread = threading.Thread(target=obfuscation_process, args=(source, dest, license_f, self.queue))
        thread.daemon = True
        thread.start()

        # Start processing queue
        self.process_queue()

    def process_queue(self):
        try:
            while True:
                message = self.queue.get_nowait()
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
            self.after(100, self.process_queue)

    def run(self):
        self.mainloop()

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
.\\python.exe .\\{relative_script_path}
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
                shutil.copy(license_path, os.path.join(dest_dir, 'license.lic'))
            else:
                queue_obj.put(f"Warning: License file not found at '{license_path}'\n")
        queue_obj.put("--- Asset Copying Finished ---\n\n")

        # 7. Create portable Python runtime
        queue_obj.put("--- Creating Portable Python Runtime ---\n")
        python_dir = os.path.dirname(sys.executable)
        runtime_files = [
            'python.exe', 'pythonw.exe', 'python3.dll',
            f'python{sys.version_info.major}{sys.version_info.minor}.dll',
            'vcruntime140.dll', 'vcruntime140_1.dll'
        ]
        for filename in runtime_files:
            src_path = os.path.join(python_dir, filename)
            if os.path.exists(src_path):
                queue_obj.put(f"  - Copying {filename}...\n")
                shutil.copy(src_path, dest_dir)

        for folder in ['DLLs', 'Lib', 'tcl']:
            src_path = os.path.join(python_dir, folder)
            dest_path = os.path.join(dest_dir, folder)
            if os.path.exists(src_path):
                queue_obj.put(f"  - Copying '{folder}' subfolder...\n")
                if os.path.exists(dest_path):
                    shutil.rmtree(dest_path)
                shutil.copytree(src_path, dest_path)
        queue_obj.put("--- Portable Runtime Created Successfully ---\n\n")
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
