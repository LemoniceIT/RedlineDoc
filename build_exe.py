import subprocess
import sys
import os
import argparse

def main():
    parser = argparse.ArgumentParser(description="Build executable for the project.")
    parser.add_argument(
        "--script",
        default="gui.py",
        help="Main script to build (default: gui.py)"
    )
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="Bundle everything into a single executable"
    )
    parser.add_argument(
        "--noconsole",
        action="store_true",
        help="Do not show console window (for GUI apps)"
    )
    args = parser.parse_args()

    # Check if the script exists
    if not os.path.isfile(args.script):
        print(f"Error: Script '{args.script}' not found.")
        sys.exit(1)

    # Build the PyInstaller command
    cmd = [sys.executable, "-m", "PyInstaller"]
    if args.onefile:
        cmd.append("--onefile")
    if args.noconsole:
        cmd.append("--windowed")  # For GUI apps, no console
    cmd.append(args.script)

    print(f"Building executable from {args.script}...")
    try:
        subprocess.check_call(cmd)
        print("Build completed. Executable is in the 'dist' folder.")
    except subprocess.CalledProcessError as e:
        print(f"Error during build: {e}")
        sys.exit(1)
    except FileNotFoundError:
        print("Error: PyInstaller not found. Please install it first:")
        print("  pip install pyinstaller")
        sys.exit(1)

if __name__ == "__main__":
    main()