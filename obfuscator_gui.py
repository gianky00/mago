import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import subprocess
import shutil
import os
import glob
import queue

class ObfuscatorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Application Obfuscator")
        self.geometry("600x400")
        self.destination_path = tk.StringVar()
        self.queue = queue.Queue()

        # Frame for path selection
        path_frame = tk.Frame(self, pady=10)
        path_frame.pack(fill='x', padx=10)

        path_label = tk.Label(path_frame, text="Destination Folder:")
        path_label.pack(side='left')

        self.path_entry = tk.Entry(path_frame, textvariable=self.destination_path, state='readonly', width=60)
        self.path_entry.pack(side='left', expand=True, fill='x', padx=5)

        self.browse_button = tk.Button(path_frame, text="Browse...", command=self.select_destination)
        self.browse_button.pack(side='left')

        # Start Button
        self.start_button = tk.Button(self, text="Start Obfuscation", command=self.start_obfuscation, state='disabled', font=('Helvetica', 12, 'bold'), pady=10)
        self.start_button.pack(pady=10)

        # Status Area
        status_frame = tk.Frame(self, pady=10)
        status_frame.pack(expand=True, fill='both', padx=10)

        status_label = tk.Label(status_frame, text="Status:")
        status_label.pack(anchor='w')

        self.status_text = tk.Text(status_frame, height=10, state='disabled', bg='black', fg='white', font=('Courier', 9))
        self.status_text.pack(expand=True, fill='both')

    def _update_status(self, message):
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message)
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

    def select_destination(self):
        path = filedialog.askdirectory(title="Select Destination Folder")
        if path:
            self.destination_path.set(path)
            self.start_button.config(state='normal')
            self._update_status(f"Destination set to: {path}\n")

    def start_obfuscation(self):
        if not self.destination_path.get():
            messagebox.showerror("Error", "Please select a destination folder first.")
            return

        self.start_button.config(state='disabled')
        self.browse_button.config(state='disabled')
        self.status_text.config(state='normal')
        self.status_text.delete('1.0', tk.END)
        self.status_text.config(state='disabled')

        # Start the queue processor
        self._process_queue()

        # Run the obfuscation in a background thread
        thread = threading.Thread(target=self._run_obfuscation_process)
        thread.daemon = True
        thread.start()

    def _process_queue(self):
        try:
            while not self.queue.empty():
                message = self.queue.get_nowait()
                if isinstance(message, tuple) and message[0] == "PROCESS_COMPLETE":
                    self.start_button.config(state='normal')
                    self.browse_button.config(state='normal')
                    # Clean up build directory
                    build_dir = "build"
                    if os.path.exists(build_dir):
                        shutil.rmtree(build_dir)
                    self._update_status("\nCleanup complete. Ready for next operation.\n")
                else:
                    self._update_status(message)
        finally:
            self.after(100, self._process_queue)

    def _run_obfuscation_process(self):
        try:
            build_dir = "build"
            dest_dir = self.destination_path.get()
            main_script = "gui.py"
            source_dir = os.path.dirname(os.path.abspath(__file__))

            self.queue.put("Starting build process...\n")

            # 1. Clean up
            if os.path.exists(build_dir):
                self.queue.put(f"Deleting old build directory: {build_dir}\n")
                shutil.rmtree(build_dir)

            # 2. Create build directory
            self.queue.put(f"Creating build directory: {build_dir}\n")
            os.makedirs(build_dir)

            # 3. Copy files to build directory
            self.queue.put("Copying application files to temporary build directory...\n")
            files_to_copy = glob.glob('*.py') + glob.glob('*.json') + glob.glob('*.db') + glob.glob('*.xlsm')
            for f in files_to_copy:
                shutil.copy(os.path.join(source_dir, f), build_dir)

            setup_dir_src = os.path.join(source_dir, 'file di setup')
            if os.path.exists(setup_dir_src):
                shutil.copytree(setup_dir_src, os.path.join(build_dir, 'file di setup'))

            self.queue.put("All source files copied.\n")

            # 4. Obfuscate
            self.queue.put("\n--- Running PyArmor ---\n")
            command = [
                "pyarmor-7", "obfuscate",
                "--recursive",
                "--output", dest_dir,
                os.path.join(build_dir, main_script)
            ]

            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, creationflags=subprocess.CREATE_NO_WINDOW)

            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    self.queue.put(output.strip() + '\n')

            rc = process.poll()
            self.queue.put("--- PyArmor Finished ---\n\n")

            if rc == 0:
                self.queue.put("Obfuscation successful!\n")

                # 5. Copy non-Python files to destination
                self.queue.put("Copying data files to destination...\n")
                data_files_to_copy = glob.glob(os.path.join(build_dir, '*.json')) + \
                                     glob.glob(os.path.join(build_dir, '*.db')) + \
                                     glob.glob(os.path.join(build_dir, '*.xlsm'))
                for f in data_files_to_copy:
                    shutil.copy(f, dest_dir)

                build_setup_dir = os.path.join(build_dir, 'file di setup')
                dest_setup_dir = os.path.join(dest_dir, 'file di setup')
                if os.path.exists(build_setup_dir):
                    if os.path.exists(dest_setup_dir):
                        shutil.rmtree(dest_setup_dir)
                    shutil.copytree(build_setup_dir, dest_setup_dir)

                self.queue.put("Data files copied.\n")
                self.queue.put(f"\nSUCCESS: Final application is ready in: {dest_dir}\n")
            else:
                self.queue.put(f"Error: PyArmor process returned error code {rc}.\n")

        except Exception as e:
            self.queue.put(f"\nAn unexpected error occurred: {str(e)}\n")
        finally:
            self.queue.put(("PROCESS_COMPLETE",))


if __name__ == "__main__":
    app = ObfuscatorGUI()
    app.mainloop()
