"""Standalone test: QBFC import Gold Coast II snapshot into blank QB 2021 template.

Usage: python test_import.py
  - Copies blank template to target dir
  - Launches QB 2021 with the target file
  - Waits for you to enter the password manually
  - Runs QBFC import from the existing snapshot
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
TARGET_QBW = TARGET_DIR / "joshs gold coast ii 21.qbw"
QB2021_EXE = Path(r"C:\Program Files (x86)\Intuit\QuickBooks 2021\QBW32PremierAccountant.exe")

def log(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

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

    # Step 1: Copy template
    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    if TARGET_QBW.exists():
        log(f"Removing old target: {TARGET_QBW}")
        TARGET_QBW.unlink()
        for ext in (".tlg", ".nd", ".DSN"):
            old = TARGET_QBW.with_suffix(ext)
            if old.exists():
                old.unlink()

    log(f"Copying template to {TARGET_QBW}")
    shutil.copy2(TEMPLATE, TARGET_QBW)
    # Copy companion files
    for ext in (".qbw.DSN", ".qbw.ND"):
        src = TEMPLATE.parent / f"Blank Template{ext}"
        if src.exists():
            dst = TARGET_DIR / f"joshs gold coast ii 21{ext}"
            shutil.copy2(src, dst)
            log(f"  Copied {src.name}")

    log(f"Template copied ({TARGET_QBW.stat().st_size / 1024 / 1024:.1f} MB)")

    # Step 2: Kill any existing QB processes
    log("Killing any existing QB processes...")
    subprocess.run(["taskkill", "/f", "/im", "QBW32PremierAccountant.exe"], 
                    capture_output=True)
    subprocess.run(["taskkill", "/f", "/im", "QBWPremierAccountant.exe"],
                    capture_output=True)
    time.sleep(3)

    # Step 3: Launch QB 2021 with the target file
    log(f"Launching QB 2021 with {TARGET_QBW}")
    subprocess.Popen([str(QB2021_EXE), str(TARGET_QBW)])
    
    log("")
    log("=" * 60)
    log("  QB 2021 is launching. ENTER THE PASSWORD MANUALLY.")
    log("  Password: 3825You171")
    log("  Once the company file is fully open, press ENTER here.")
    log("=" * 60)
    input("\n>>> Press ENTER when QB 2021 has the file open... ")

    # Step 4: QBFC import
    log("Starting QBFC import...")
    try:
        from qbfc_import import import_company_via_qbfc
        results = import_company_via_qbfc(
            snapshot_path=SNAPSHOT,
            qbw_path=TARGET_QBW,
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

if __name__ == "__main__":
    main()
