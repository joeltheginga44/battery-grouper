import time
import os
import subprocess
import sys

# --- CONFIGURATION ---
# The file you want to watch for changes
SCRIPT_TO_WATCH = "battery_gui.py"

# The command to run when a change is detected
# Note: We use "python -m PyInstaller" to ensure it works even if PATH is tricky
REBUILD_COMMAND = [sys.executable, "-m", "PyInstaller", "--onefile", "--windowed", SCRIPT_TO_WATCH]

# How often to check for changes (in seconds)
CHECK_INTERVAL = 2

def get_file_mod_time(filepath):
    try:
        return os.path.getmtime(filepath)
    except FileNotFoundError:
        return None

def run_rebuild():
    print("\n" + "="*50)
    print(f"🔄 Change detected in {SCRIPT_TO_WATCH}!")
    print("🛠️  Rebuilding EXE... (Please wait)")
    print("="*50)
    
    try:
        # Run the PyInstaller command and show the output
        subprocess.run(REBUILD_COMMAND, check=True)
        print("\n✅ Rebuild complete! Your new .exe is in the 'dist' folder.")
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Build failed! Error: {e}")
        print("Check the output above for details.")
    except FileNotFoundError:
        print("\n❌ Error: PyInstaller not found. Make sure it is installed with 'pip install pyinstaller'")

def main():
    if not os.path.exists(SCRIPT_TO_WATCH):
        print(f"❌ Error: Cannot find {SCRIPT_TO_WATCH} in the current folder.")
        print("Make sure this watcher.py file is in the same folder as your script.")
        return

    print("="*50)
    print(" 👀 BATTERY GROUPER AUTO-BUILDER WATCHER 👀 ")
    print("="*50)
    print(f"Monitoring: {SCRIPT_TO_WATCH}")
    print(f"Rebuild Command: {' '.join(REBUILD_COMMAND)}")
    print("\nInstructions:")
    print("1. Leave this window open.")
    print("2. Edit and SAVE your battery_gui.py file.")
    print("3. This script will automatically rebuild the .exe.")
    print("\nPress Ctrl+C to stop the watcher.")
    print("="*50 + "\n")

    # Get the initial modification time
    last_mod_time = get_file_mod_time(SCRIPT_TO_WATCH)

    try:
        while True:
            time.sleep(CHECK_INTERVAL)
            current_mod_time = get_file_mod_time(SCRIPT_TO_WATCH)

            if current_mod_time is None:
                # File was deleted
                print(f"⚠️  Warning: {SCRIPT_TO_WATCH} was deleted!")
                last_mod_time = None
                continue

            if last_mod_time is None:
                # File was recreated
                print(f"✅ {SCRIPT_TO_WATCH} has been recreated.")
                run_rebuild()
                last_mod_time = current_mod_time
                continue

            # Check if the file has been modified
            if current_mod_time != last_mod_time:
                run_rebuild()
                last_mod_time = current_mod_time

    except KeyboardInterrupt:
        print("\n\n🛑 Watcher stopped. Have a great day!")

if __name__ == "__main__":
    main()