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
        default=True,
        help="Bundle everything into a single executable (default)"
    )
    parser.add_argument(
        "--many-files",
        dest="onefile",
        action="store_false",
        help="Create a folder with the executable and dependencies"
    )
    parser.add_argument(
        "--noconsole",
        action="store_true",
        help="Do not show console window (for GUI apps)"
    )
    parser.add_argument(
        "--version",
        type=str,
        default=None,
        help="Set the version string for the executable (e.g., 1.0.0)"
    )
    parser.add_argument(
        "--icon",
        type=str,
        default=None,
        help="Path to an icon file (.ico) for the executable"
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
    if args.version:
        import json
        import tempfile
        version_metadata = {
            "FileVersion": args.version,
            "ProductVersion": args.version
        }
        with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
            json.dump(version_metadata, f)
            version_file = f.name
        cmd.extend(["--version-file", version_file])
    if args.icon:
        if os.path.isfile(args.icon):
            cmd.extend(["--icon", args.icon])
        else:
            print(f"Warning: Icon file '{args.icon}' not found, skipping icon.")
    cmd.append(args.script)

    # Determine version string for display
    version_str = args.version if args.version else "development"

    print(f"Building executable from {args.script} (version: {version_str})...")
    try:
        subprocess.check_call(cmd)
        if args.onefile:
            print(f"Build completed (v{version_str}). Single executable is in the 'dist' folder.")
        else:
            print(f"Build completed (v{version_str}). Executable and dependencies are in the 'dist' folder.")
    except subprocess.CalledProcessError as e:
        print(f"Error during build: {e}")
        sys.exit(1)
    except FileNotFoundError:
        print("Error: PyInstaller not found. Please install it first:")
        print("  pip install pyinstaller")
        sys.exit(1)

if __name__ == "__main__":
    main()