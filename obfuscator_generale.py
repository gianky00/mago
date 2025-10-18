import argparse
import subprocess
import shutil
import os
import glob
import sys

def main():
    parser = argparse.ArgumentParser(description="A general-purpose tool to obfuscate a directory of Python scripts.")
    parser.add_argument("--source", required=True, help="The source directory containing the Python scripts to obfuscate.")
    parser.add_argument("--dest", required=True, help="The destination directory for the obfuscated output.")
    parser.add_argument("--license", help="Optional path to a .lic file to be included.")
    args = parser.parse_args()

    source_dir = os.path.abspath(args.source)
    dest_dir = os.path.abspath(args.dest)
    obfuscated_dir = os.path.join(dest_dir, "obfuscated")

    if not os.path.isdir(source_dir):
        print(f"Error: Source directory '{source_dir}' does not exist.")
        sys.exit(1)

    print("--- Starting Obfuscation Process ---")

    # 1. Clean up and create directories
    if os.path.exists(dest_dir):
        print(f"Removing existing destination directory: {dest_dir}")
        shutil.rmtree(dest_dir)
    print(f"Creating destination directory: {dest_dir}")
    os.makedirs(obfuscated_dir)

    # 2. Find all Python scripts
    print(f"Searching for Python scripts in: {source_dir}")
    all_scripts = glob.glob(os.path.join(source_dir, '*.py'))
    if not all_scripts:
        print("Error: No Python files found in the source directory.")
        sys.exit(1)

    script_basenames = [os.path.basename(s) for s in all_scripts]
    print(f"Found {len(script_basenames)} scripts: {', '.join(script_basenames)}")

    # 3. Run PyArmor obfuscation
    print("\n--- Running PyArmor ---")
    command = [
        "pyarmor", "gen",
        "--outer",
        "--output", obfuscated_dir,
    ] + all_scripts

    try:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        for line in process.stdout:
            print(line, end='')
        process.wait()
        if process.returncode != 0:
            raise subprocess.CalledProcessError(process.returncode, command)
        print("--- PyArmor Finished Successfully ---\n")
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"\n--- PyArmor Failed ---")
        print(f"An error occurred during obfuscation: {e}")
        sys.exit(1)

    # 4. Create .bat launchers
    print("--- Creating .bat Launchers ---")
    for script_path in all_scripts:
        script_name = os.path.basename(script_path)
        bat_name = os.path.splitext(script_name)[0] + ".bat"
        bat_path = os.path.join(dest_dir, bat_name)

        # Path to the obfuscated script relative to the .bat file
        relative_script_path = os.path.join("obfuscated", script_name)

        launcher_content = f'''@echo off
setlocal
REM Change directory to the script's location to ensure correct asset loading and license checking.
cd /d %~dp0
REM Run the obfuscated application using the LOCAL python.exe.
.\\python.exe .\\{relative_script_path}
endlocal
pause
'''
        print(f"Creating launcher for {script_name} at {bat_path}")
        with open(bat_path, 'w', encoding='utf-8') as f:
            f.write(launcher_content)
    print("--- Launchers Created Successfully ---\n")

    # 5. Copy non-Python assets
    print("--- Copying Assets ---")
    for item in glob.glob(os.path.join(source_dir, '*')):
        if not item.endswith('.py'):
            dest_item = os.path.join(dest_dir, os.path.basename(item))
            if os.path.isdir(item):
                print(f"Copying directory: {os.path.basename(item)}")
                shutil.copytree(item, dest_item)
            else:
                print(f"Copying file: {os.path.basename(item)}")
                shutil.copy(item, dest_item)

    # 6. Copy license file if provided
    if args.license:
        if os.path.exists(args.license):
            print(f"Copying license file from: {args.license}")
            shutil.copy(args.license, os.path.join(dest_dir, 'license.lic'))
        else:
            print(f"Warning: License file not found at '{args.license}'")
    print("--- Asset Copying Finished ---\n")

    # 7. Create portable Python runtime
    print("--- Creating Portable Python Runtime ---")
    python_dir = os.path.dirname(sys.executable)
    runtime_files = [
        'python.exe', 'pythonw.exe', 'python3.dll',
        f'python{sys.version_info.major}{sys.version_info.minor}.dll',
        'vcruntime140.dll', 'vcruntime140_1.dll'
    ]
    for filename in runtime_files:
        src_path = os.path.join(python_dir, filename)
        if os.path.exists(src_path):
            print(f"  - Copying {filename}...")
            shutil.copy(src_path, dest_dir)

    for folder in ['DLLs', 'Lib', 'tcl']:
        src_path = os.path.join(python_dir, folder)
        dest_path = os.path.join(dest_dir, folder)
        if os.path.exists(src_path):
            print(f"  - Copying '{folder}' subfolder...")
            if os.path.exists(dest_path):
                shutil.rmtree(dest_path)
            shutil.copytree(src_path, dest_path)
    print("--- Portable Runtime Created Successfully ---\n")
    print("====== OBFUSCATION COMPLETE ======")
    print(f"Final application is ready in: {dest_dir}")

if __name__ == "__main__":
    main()
