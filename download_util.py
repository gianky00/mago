import requests
import sys
from tqdm import tqdm

def download_file(url, output_path):
    """
    Downloads a file from a URL to a specified output path with a progress bar.

    Args:
        url (str): The URL of the file to download.
        output_path (str): The local path to save the downloaded file.
    """
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()  # Raise an exception for bad status codes (4xx or 5xx)

        total_size_in_bytes = int(response.headers.get('content-length', 0))
        block_size = 1024  # 1 Kibibyte

        progress_bar = tqdm(total=total_size_in_bytes, unit='iB', unit_scale=True, desc="Downloading")
        with open(output_path, 'wb') as file:
            for data in response.iter_content(block_size):
                progress_bar.update(len(data))
                file.write(data)
        progress_bar.close()

        if total_size_in_bytes != 0 and progress_bar.n != total_size_in_bytes:
            print("ERROR: Downloaded size does not match expected size.")
            sys.exit(1)

        print(f"Download complete. File saved to {output_path}")

    except requests.exceptions.RequestException as e:
        print(f"ERROR: A network error occurred: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python download_util.py <url> <output_path>")
        sys.exit(1)

    download_url = sys.argv[1]
    local_output_path = sys.argv[2]

    download_file(download_url, local_output_path)