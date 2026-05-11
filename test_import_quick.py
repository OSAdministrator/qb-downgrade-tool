"""Quick test: QBFC import into ALREADY-OPEN QB 2021.

Usage:
  1. Make sure QB 2021 has 'Blank Template.qbw' open in the target dir
  2. Run: python test_import_quick.py
  Skips template copy and QB launch — just does the QBFC import.
"""
import sys
import time
from pathlib import Path

SNAPSHOT = Path(r"C:\QBDowngrade\Output\joshs gold coast ii 23\exports\company_snapshot.json")

def log(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

def main():
    if not SNAPSHOT.exists():
        log(f"ERROR: Snapshot not found at {SNAPSHOT}")
        sys.exit(1)

    log(f"Snapshot: {SNAPSHOT} ({SNAPSHOT.stat().st_size / 1024:.0f} KB)")
    log("")
    log("Make sure QB 2021 has Blank Template.qbw OPEN and all dialogs dismissed.")
    log("Starting QBFC import in 3 seconds...")
    time.sleep(3)

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

if __name__ == "__main__":
    main()
