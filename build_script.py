
import PyInstaller.__main__
import os
import platform

# --- Configuration ---
APP_NAME = "ContactsMerger"
ENTRY_POINT = "contacts_Merger_Frontend.py"
ICON_PATH = "assets/logo.png"
ASSETS_PATH = "assets"
BACKEND_SCRIPT = "Contacts_Merger_Backend"


def main():
    """
    Runs PyInstaller to build the executable.
    """

    # --- Platform-specific adjustments ---
    separator = ';' if platform.system() == "Windows" else ':'

    # --- Build Command ---
    command = [
        '--name', APP_NAME,
        '--onefile',
        '--windowed',  # Use for GUI applications to not show a console
        f'--icon={ICON_PATH}',
        f'--add-data={ASSETS_PATH}{separator}assets',
        f'--hidden-import={BACKEND_SCRIPT}',
        '--hidden-import=pyodbc',
        ENTRY_POINT,
    ]

    print(f"Running PyInstaller with command: pyinstaller {" ".join(command)}")
    
    try:
        PyInstaller.__main__.run(command)
        print("\nBuild process finished successfully!")
        print(f"Executable created in the 'dist' folder: {APP_NAME}.exe")
    except Exception as e:
        print(f"\nAn error occurred during the build process: {e}")
        print("Please check the PyInstaller logs for more details.")

if __name__ == "__main__":
    print("Starting the build process...")
    # Note for Windows users:
    if platform.system() == "Windows":
        print("On Windows, it is recommended to use a .ico file for the icon.")
        print("If the build fails, consider converting logo.png to logo.ico.")
    
    main()
