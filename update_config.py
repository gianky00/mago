import json
import sys
import os

CONFIG_FILE = 'config.json'
TESSERACT_PATH_KEY = 'path_tesseract_cmd'
GENERAL_SECTION_KEY = 'generali'

def update_tesseract_path(new_path):
    """
    Updates the Tesseract executable path in the config.json file.

    Args:
        new_path (str): The new, absolute path to the tesseract.exe.
    """
    if not os.path.exists(CONFIG_FILE):
        print(f"ERROR: Configuration file '{CONFIG_FILE}' not found.")
        sys.exit(1)

    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
    except json.JSONDecodeError:
        print(f"ERROR: Could not decode JSON from '{CONFIG_FILE}'.")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: An unexpected error occurred while reading the file: {e}")
        sys.exit(1)

    # Ensure the 'generali' section exists
    if GENERAL_SECTION_KEY not in config_data:
        config_data[GENERAL_SECTION_KEY] = {}

    # Update the path
    current_path = config_data[GENERAL_SECTION_KEY].get(TESSERACT_PATH_KEY)
    if current_path == new_path:
        print(f"Tesseract path is already up-to-date in '{CONFIG_FILE}'.")
        return

    config_data[GENERAL_SECTION_KEY][TESSERACT_PATH_KEY] = new_path

    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)
        print(f"Successfully updated Tesseract path in '{CONFIG_FILE}' to '{new_path}'.")
    except Exception as e:
        print(f"ERROR: Failed to write updated configuration to '{CONFIG_FILE}': {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python update_config.py \"<full_path_to_tesseract_exe>\"")
        sys.exit(1)

    tesseract_exe_path = sys.argv[1]

    if not os.path.isabs(tesseract_exe_path):
        print(f"ERROR: The provided path '{tesseract_exe_path}' is not an absolute path.")
        sys.exit(1)

    # The batch script will verify the file exists, but we can add a check here for safety
    if not os.path.exists(tesseract_exe_path):
        print(f"ERROR: The file specified does not exist: '{tesseract_exe_path}'")
        sys.exit(1)

    update_tesseract_path(tesseract_exe_path)