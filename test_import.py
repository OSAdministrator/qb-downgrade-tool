"""Standalone test: QBFC import Gold Coast II snapshot into blank QB 2021 template.

Usage: python test_import.py
  - Copies blank template to target dir (keeping name 'Blank Template.qbw' for QBFC auth)
  - Launches QB 2021 with the target file
  - Template has no password (pre-authorized)
  - Runs QBFC import from the existing snapshot
  - After import, renames to final name
"""
import shutil
import subprocess
import sys
import time
from pathlib import Path

# Paths
TEMPLATE = Path(r"C:\QBDowngrade\QB-2021 Template\Blank Template.qbw")
SNAPSHOT = Path(r"C:\QBDowngrade\Output\joshs gold coast ii 23\exports\company_snapshot.json")
TARGET_DIR = Path(r"C:\QBDowngrade\Output\joshs gold coast ii 23\target")
WORKING_QBW = TARGET_DIR / "Blank Template.qbw"  # keep original name for QBFC auth
FINAL_QBW = TARGET_DIR / "joshs gold coast ii 21.qbw"
QB2021_EXE = Path(r"C:\Program Files (x86)\Intuit\QuickBooks 2021\QBW32PremierAccountant.exe")

def log(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

def kill_qb():
    """Kill all QB-related processes."""
    for proc in ["QBW32PremierAccountant.exe", "QBWPremierAccountant.exe",
                 "qbupdate.exe", "qbmapi64.exe"]:
        subprocess.run(["taskkill", "/f", "/im", proc], capture_output=True)
    # Also kill CefSharp/Intuit background processes
    for pattern in ["CefSharp*", "Intuit*"]:
        subprocess.run(f'taskkill /f /im "{pattern}"', shell=True, capture_output=True)
    time.sleep(3)

def main():
    # Verify prerequisites
    if not TEMPLATE.exists():
        log(f"ERROR: Template not found at {TEMPLATE}")
        sys.exit(1)
    if not SNAPSHOT.exists():
        log(f"ERROR: Snapshot not found at {SNAPSHOT}")
        sys.exit(1)
    if not QB2021_EXE.exists():
        log(f"ERROR: QB 2021 not found at {QB2021_EXE}")
        sys.exit(1)

    # Step 1: Kill any existing QB
    log("Killing any existing QB processes...")
    kill_qb()

    # Step 2: Copy template (keep name for QBFC auth)
    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    for f in TARGET_DIR.glob("*"):
        try:
            f.unlink()
            log(f"  Removed {f.name}")
        except PermissionError:
            log(f"  WARN: could not remove {f.name} (locked?)")

    log(f"Copying template to {WORKING_QBW}")
    shutil.copy2(TEMPLATE, WORKING_QBW)
    for ext in (".qbw.DSN", ".qbw.ND", ".qbw.tlg"):
        src = TEMPLATE.parent / f"Blank Template{ext}"
        if src.exists():
            dst = TARGET_DIR / f"Blank Template{ext}"
            try:
                shutil.copy2(src, dst)
            except PermissionError:
                pass
    log(f"Template copied ({WORKING_QBW.stat().st_size / 1024 / 1024:.1f} MB)")

    # Step 3: Launch QB 2021
    log(f"Launching QB 2021 with {WORKING_QBW}")
    subprocess.Popen([str(QB2021_EXE), str(WORKING_QBW)])

    log("")
    log("=" * 60)
    log("  QB 2021 is launching with Blank Template.")
    log("  Password: 01Hello02!@!")
    log("  Dismiss any startup dialogs after login.")
    log("  Once the company file is fully open, press ENTER here.")
    log("=" * 60)
    input("\n>>> Press ENTER when QB 2021 has the file open... ")

    # Step 4: QBFC import (qbw_path=None = use already-open file)
    log("Starting QBFC import...")
    try:
        from qbfc_import import import_company_via_qbfc
        results = import_company_via_qbfc(
            snapshot_path=SNAPSHOT,
            qbw_path=None,
            log_fn=log,
            skip_transactions=False,
        )
        log("")
        log("=" * 60)
        log("  IMPORT COMPLETE!")
        log(f"  Total records: {sum(results.values())}")
        for k, v in results.items():
            log(f"    {k}: {v}")
        log("=" * 60)
    except Exception as exc:
        import traceback
        log(f"IMPORT FAILED: {exc}")
        traceback.print_exc()
        sys.exit(1)

    # Step 5: Close QB and rename — wait for file locks to release
    log("Closing QB...")
    kill_qb()

    # Poll for lock release (up to 30s)
    log("Waiting for file locks to release...")
    for attempt in range(15):
        try:
            with open(WORKING_QBW, 'r+b'):
                pass
            log(f"  File unlocked after {attempt * 2}s")
            break
        except (PermissionError, OSError):
            time.sleep(2)
    else:
        log("  WARN: file still locked after 30s, attempting rename anyway")

    log(f"Renaming {WORKING_QBW.name} -> {FINAL_QBW.name}")
    try:
        WORKING_QBW.rename(FINAL_QBW)
        for ext in (".qbw.ND", ".qbw.DSN", ".qbw.tlg"):
            src = TARGET_DIR / f"Blank Template{ext}"
            dst = TARGET_DIR / f"joshs gold coast ii 21{ext}"
            if src.exists():
                try:
                    src.rename(dst)
                except Exception:
                    pass
        log("Rename complete.")
    except Exception as exc:
        log(f"WARN: Rename failed: {exc} — file is still 'Blank Template.qbw'")

    log("Done!")

if __name__ == "__main__":
    main()
