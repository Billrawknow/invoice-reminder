"""
Builds the folder you hand to a client: no Python, no code, just double-click files.

    python build_release.py               # Windows: builds InvoiceReminder.exe with PyInstaller, then packages
    python build_release.py --sample      # include the 13 sample invoices (for demos) instead of a blank sheet
    python build_release.py --no-exe      # package the Python script instead (client then needs Python)

Output:
    release/InvoiceReminder/              the folder the client gets
    release/InvoiceReminder_v1.0.zip      the same folder, zipped, to email or put on a flash disk

What's inside the client folder:
    InvoiceReminder.exe     the tool (or invoice_reminder.py + requirements.txt with --no-exe)
    config.env              the client's settings, with placeholders to fill in
    invoices.xlsx           blank sheet with the right columns and a yes/no dropdown
    setup.bat               step 2 of setup: edit settings, check them, turn on the daily run
    run.bat                 check now and open the Outbox
    remove_schedule.bat     turn the daily run off
    READ ME FIRST.txt       the 3-step setup guide and what each message means

Your own config.env, invoices, Outbox, logs and history are never copied.
"""

import argparse
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
VERSION = "1.0"
NAME = "InvoiceReminder"
RELEASE = HERE / "release"
FOLDER = RELEASE / NAME

COMMON = {                        # source in repo -> name in the client folder
    "config.env.example": "config.env",
    "setup.bat": "setup.bat",
    "run.bat": "run.bat",
    "remove_schedule.bat": "remove_schedule.bat",
    "docs/READ_ME_FIRST.txt": "READ ME FIRST.txt",
}
NEVER_SHIP = {".env", "reminder_history.csv", "automation.log", "Outbox"}


def build_exe():
    if os.name != "nt":
        sys.exit("The .exe must be built on Windows (PyInstaller makes programs for the system it runs on). "
                 "Use --no-exe to package the script instead.")
    import importlib.util
    if importlib.util.find_spec("PyInstaller") is None:
        sys.exit("PyInstaller isn't installed. Run:  pip install pyinstaller")
    print("Building InvoiceReminder.exe (takes a minute)...")
    result = subprocess.run([sys.executable, "-m", "PyInstaller", "--onefile", "--console", "--clean",
                             "--noconfirm", "--name", NAME, "invoice_reminder.py"],
                            cwd=HERE, capture_output=True, text=True)
    exe = HERE / "dist" / f"{NAME}.exe"
    if result.returncode != 0 or not exe.exists():
        print(result.stdout[-2000:], result.stderr[-2000:])
        sys.exit("PyInstaller failed (output above).")
    return exe


def package(exe=None, sample=False):
    if FOLDER.exists():
        shutil.rmtree(FOLDER)
    FOLDER.mkdir(parents=True)

    if exe:
        shutil.copy2(exe, FOLDER / exe.name)
    else:
        for f in ("invoice_reminder.py", "requirements.txt"):
            shutil.copy2(HERE / f, FOLDER / f)
    for src, dest in COMMON.items():
        shutil.copy2(HERE / src, FOLDER / dest)

    sys.path.insert(0, str(HERE))
    import make_sample_invoices
    make_sample_invoices.build(FOLDER / "invoices.xlsx", make_sample_invoices.ROWS if sample else [])

    leaked = [p.name for p in FOLDER.rglob("*") if p.name in NEVER_SHIP]
    if leaked:
        sys.exit(f"Refusing to ship private files: {leaked}")

    zip_path = RELEASE / f"{NAME}_v{VERSION}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(FOLDER.rglob("*")):
            z.write(path, Path(NAME) / path.relative_to(FOLDER))
    return zip_path


def main():
    p = argparse.ArgumentParser(description="Build the client folder and zip.")
    p.add_argument("--no-exe", action="store_true", help="ship the Python script instead of an .exe")
    p.add_argument("--sample", action="store_true", help="include sample invoices instead of a blank sheet")
    args = p.parse_args()
    exe = None if args.no_exe else build_exe()
    zip_path = package(exe, args.sample)
    print(f"\nClient folder: {FOLDER.relative_to(HERE)}")
    for f in sorted(FOLDER.iterdir()):
        print(f"  {f.name:<24} {f.stat().st_size / 1024:>8,.0f} KB")
    print(f"Zip to send:   {zip_path.relative_to(HERE)}  ({zip_path.stat().st_size / 1024 / 1024:.1f} MB)")


if __name__ == "__main__":
    main()
