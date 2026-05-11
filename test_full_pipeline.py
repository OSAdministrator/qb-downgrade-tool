"""Full end-to-end test: Export from QB 2023 -> Import into QB 2021.

Usage: python test_full_pipeline.py

Sequence:
  1. Launch QB 2023 with source file
  2. QBFC export (IIF + CSV + snapshot with per-type transactions)
  3. Close QB 2023
  4. Copy blank QB 2021 template (keeping name for QBFC auth)
  5. Launch QB 2021 with template copy
  6. QBFC import from snapshot
  7. Close QB 2021
  8. Rename to final target name
"""
import shutil
import subprocess
import sys
import time
from pathlib import Path

# --- Configuration ---
SOURCE_QBW = Path(r"C:\Users\AbacusAgent\Desktop\joshs gold coast ii 23.qbw")
SOURCE_PASSWORD = "3825You171"

TEMPLATE = Path(r"C:\QBDowngrade\QB-2021 Template\Blank Template.qbw")
TEMPLATE_PASSWORD = "01Hello02!@!"

OUTPUT_DIR = Path(r"C:\QBDowngrade\Output\joshs gold coast ii 23")
EXPORTS_DIR = OUTPUT_DIR / "exports"
TARGET_DIR = OUTPUT_DIR / "target"

WORKING_QBW = TARGET_DIR / "Blank Template.qbw"  # keeps QBFC auth
FINAL_QBW = TARGET_DIR / "joshs gold coast ii 21.qbw"

QB2023_EXE = Path(r"C:\Program Files\Intuit\QuickBooks 2023\QBWPremierAccountant.exe")
QB2021_EXE = Path(r"C:\Program Files (x86)\Intuit\QuickBooks 2021\QBW32PremierAccountant.exe")


def log(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")


def kill_qb():
    """Kill ALL QB-related processes."""
    for proc in ["QBW32PremierAccountant.exe", "QBWPremierAccountant.exe",
                 "qbupdate.exe", "qbmapi64.exe"]:
        subprocess.run(["taskkill", "/f", "/im", proc], capture_output=True)
    for pattern in ["CefSharp*", "Intuit*"]:
        subprocess.run(f'taskkill /f /im "{pattern}"', shell=True, capture_output=True)
    time.sleep(3)


def wait_for_unlock(path, timeout=30):
    """Wait for a file to be unlocked."""
    for i in range(timeout // 2):
        try:
            with open(path, 'r+b'):
                pass
            log(f"  File unlocked after {i * 2}s")
            return True
        except (PermissionError, OSError):
            time.sleep(2)
    log(f"  WARN: file still locked after {timeout}s")
    return False


def main():
    start = time.time()

    # --- Verify prerequisites ---
    for label, path in [("Source QBW", SOURCE_QBW), ("Template", TEMPLATE),
                        ("QB 2023", QB2023_EXE), ("QB 2021", QB2021_EXE)]:
        if not path.exists():
            log(f"ERROR: {label} not found at {path}")
            sys.exit(1)

    log("="*60)
    log("  QuickBooks TimeWarp — Full Pipeline Test")
    log(f"  Source:   {SOURCE_QBW.name}")
    log(f"  Template: {TEMPLATE.name}")
    log(f"  Target:   {FINAL_QBW.name}")
    log("="*60)

    # =========================================================
    # PHASE 1: EXPORT from QB 2023
    # =========================================================
    log("")
    log(">>> PHASE 1: EXPORT from QB 2023")
    kill_qb()

    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)

    log(f"Launching QB 2023 with {SOURCE_QBW.name}...")
    subprocess.Popen([str(QB2023_EXE), str(SOURCE_QBW)])

    log("")
    log("  Enter password in QB 2023: " + SOURCE_PASSWORD)
    log("  Dismiss any startup dialogs.")
    input(">>> Press ENTER when QB 2023 has the file open... ")

    log("Connecting to QB 2023 via QBFC...")
    from qbfc_export import export_company_via_qbfc

    exported = export_company_via_qbfc(
        qbw_path=None,  # bind to already-open file
        export_dir=EXPORTS_DIR,
        log_fn=log,
    )

    snapshot_path = exported.get("snapshot")
    log(f"Export complete. Snapshot: {snapshot_path}")

    # Show transaction stats from snapshot
    import json
    with open(snapshot_path, 'r', encoding='utf-8') as f:
        snap = json.load(f)
    txns = snap.get('transactions', [])
    has_lines = any('lines' in t for t in txns[:5])
    log(f"  Transactions: {len(txns)}, has_lines={has_lines}")
    if txns:
        type_counts = {}
        for t in txns:
            type_counts[t.get('type', '?')] = type_counts.get(t.get('type', '?'), 0) + 1
        for tname, count in sorted(type_counts.items()):
            log(f"    {tname}: {count}")

    # Close QB 2023
    log("Closing QB 2023...")
    kill_qb()

    # =========================================================
    # PHASE 2: IMPORT into QB 2021
    # =========================================================
    log("")
    log(">>> PHASE 2: IMPORT into QB 2021")

    # Clean target dir
    if TARGET_DIR.exists():
        for f in TARGET_DIR.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
    TARGET_DIR.mkdir(parents=True, exist_ok=True)

    # Copy template
    log(f"Copying template to {WORKING_QBW}")
    shutil.copy2(TEMPLATE, WORKING_QBW)
    for ext in (".qbw.DSN", ".qbw.ND", ".qbw.tlg"):
        src = TEMPLATE.parent / f"Blank Template{ext}"
        if src.exists():
            try:
                shutil.copy2(src, TARGET_DIR / f"Blank Template{ext}")
            except Exception:
                pass
    log(f"  Template copied ({WORKING_QBW.stat().st_size / 1024 / 1024:.1f} MB)")

    # Launch QB 2021
    log(f"Launching QB 2021 with {WORKING_QBW.name}...")
    subprocess.Popen([str(QB2021_EXE), str(WORKING_QBW)])

    log("")
    log("  Enter password in QB 2021: " + TEMPLATE_PASSWORD)
    log("  Dismiss any startup dialogs.")
    input(">>> Press ENTER when QB 2021 has the file open... ")

    # QBFC import
    log("Starting QBFC import...")
    from qbfc_import import import_company_via_qbfc

    results = import_company_via_qbfc(
        snapshot_path=Path(str(snapshot_path)),
        qbw_path=None,  # bind to already-open file
        log_fn=log,
        skip_transactions=False,
    )

    # Close QB 2021 and rename
    log("Closing QB 2021...")
    kill_qb()

    log("Waiting for file locks...")
    wait_for_unlock(WORKING_QBW)

    log(f"Renaming {WORKING_QBW.name} -> {FINAL_QBW.name}")
    try:
        if FINAL_QBW.exists():
            FINAL_QBW.unlink()
        WORKING_QBW.rename(FINAL_QBW)
        for ext in (".qbw.ND", ".qbw.DSN", ".qbw.tlg"):
            src = TARGET_DIR / f"Blank Template{ext}"
            dst = TARGET_DIR / f"joshs gold coast ii 21{ext}"
            if src.exists():
                try:
                    src.rename(dst)
                except Exception:
                    pass
        log("  Rename complete.")
    except Exception as exc:
        log(f"  WARN: Rename failed: {exc}")

    # =========================================================
    # SUMMARY
    # =========================================================
    elapsed = time.time() - start
    log("")
    log("="*60)
    log("  PIPELINE COMPLETE!")
    log(f"  Elapsed: {elapsed/60:.1f} minutes")
    log("")
    log("  Import Results:")
    total = 0
    for k, v in results.items():
        log(f"    {k}: {v}")
        total += v
    log(f"    TOTAL: {total}")
    log("")
    log(f"  Output: {FINAL_QBW}")
    log("="*60)


if __name__ == "__main__":
    main()
