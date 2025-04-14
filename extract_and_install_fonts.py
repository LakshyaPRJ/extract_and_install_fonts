import os
import zipfile
import shutil
import win32com.client
import pythoncom
from pathlib import Path
import ctypes
import sys
import subprocess
import time

def is_admin():
    """Check if the script is running with administrative privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def extract_zip(zip_path, extract_to):
    """Extract a zip file to the specified directory."""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        return True
    except Exception as e:
        print(f"Error extracting {zip_path}: {e}")
        return False

def install_font(font_path):
    """Install a font file on Windows using Shell to register it properly."""
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("Shell.Application")
        fonts_folder = shell.NameSpace(0x14)  # Fonts folder (CSIDL_FONTS)

        font_name = os.path.basename(font_path)
        fonts_dir = os.path.join(os.environ["WINDIR"], "Fonts")
        dest_path = os.path.join(fonts_dir, font_name)

        if not os.path.exists(dest_path):
            # Copy the font file to Fonts directory
            shutil.copy(font_path, fonts_dir)

            # Register the font using Shell
            font_folder = shell.NameSpace(fonts_dir)
            font_folder.CopyHere(font_path, 16)  # 16 = no UI, overwrite silently
            print(f"Installed and registered font: {font_name}")
        else:
            print(f"Font already exists: {font_name}")
    except PermissionError:
        print(f"Permission denied: Run the script as Administrator to install {font_name}")
    except Exception as e:
        print(f"Error installing font {font_path}: {e}")
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def refresh_font_cache():
    """Restart the Windows Font Cache Service to refresh fonts."""
    try:
        print("Attempting to refresh font cache...")
        # Stop the Font Cache service
        subprocess.run(
            ["net", "stop", "fontcache"],
            check=True,
            capture_output=True,
            text=True
        )
        time.sleep(1)
        # Start the Font Cache service
        subprocess.run(
            ["net", "start", "fontcache"],
            check=True,
            capture_output=True,
            text=True
        )
        print("Font cache refreshed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error refreshing font cache: {e}. Try restarting your application.")
    except Exception as e:
        print(f"Unexpected error refreshing font cache: {e}")

def process_folder(folder_path):
    """Process all zip files and fonts in the folder."""
    folder = Path(folder_path)
    temp_extract_dir = folder / "temp_extracted"

    # Create temp directory if it doesn't exist
    temp_extract_dir.mkdir(exist_ok=True)

    # Find all zip files
    zip_files = list(folder.glob("*.zip"))

    if not zip_files:
        print("No zip files found in the folder.")
        return

    for zip_file in zip_files:
        print(f"Processing {zip_file.name}...")

        # Create a unique extraction folder for this zip
        extract_path = temp_extract_dir / zip_file.stem
        extract_path.mkdir(exist_ok=True)

        # Extract zip contents
        if extract_zip(zip_file, extract_path):
            # Look for font files in extracted contents
            for root, _, files in os.walk(extract_path):
                for file in files:
                    if file.lower().endswith(('.ttf', '.otf')):
                        font_path = os.path.join(root, file)
                        install_font(font_path)

            # Clean up extracted files
            shutil.rmtree(extract_path, ignore_errors=True)

            # Delete the original zip file
            try:
                zip_file.unlink()
                print(f"Deleted processed zip: {zip_file.name}")
            except Exception as e:
                print(f"Error deleting {zip_file.name}: {e}")

    # Clean up temp directory if empty
    try:
        temp_extract_dir.rmdir()
    except:
        pass

    # Refresh font cache after all fonts are installed
    refresh_font_cache()

def main():
    # Check for admin privileges
    if not is_admin():
        print("This script requires administrative privileges to install fonts.")
        print("Please run the script as Administrator.")
        sys.exit(1)

    # Get the folder path
    folder_path = input("Enter the folder path (press Enter for current directory): ").strip()
    if not folder_path:
        folder_path = os.getcwd()

    if not os.path.isdir(folder_path):
        print("Invalid folder path!")
        sys.exit(1)

    print(f"Processing folder: {folder_path}")
    process_folder(folder_path)
    print("Processing complete!")
    print("If needed, please restart your applications (e.g., Adobe Illustrator) to see the new fonts.")

if __name__ == "__main__":
    main()