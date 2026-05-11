"""Quick test: QBFC import into ALREADY-OPEN QB 2021.

Usage: Open QB 2021 with the target file first, then run this.
Skips template copy and QB launch — just does the QBFC import.
"""
import sys
import time
from pathlib import Path

SNAPSHOT = Path(r"C:\QBDowngrade\Output\joshs gold coast ii 23\exports\company_snapshot.json")
TARGET_QBW = Path(r"C:\QBDowngrade\Output\joshs gold coast ii 23\target\joshs gold coast ii 21.qbw")

def log(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

def main():
    if not SNAPSHOT.exists():
        log(f"ERROR: Snapshot not found at {SNAPSHOT}")
        sys.exit(1)
    if not TARGET_QBW.exists():
        log(f"ERROR: Target QBW not found at {TARGET_QBW}")
        sys.exit(1)

    log(f"Snapshot: {SNAPSHOT} ({SNAPSHOT.stat().st_size / 1024:.0f} KB)")
    log(f"Target:   {TARGET_QBW}")
    log("")
    log("Make sure QB 2021 has this file OPEN and all dialogs dismissed.")
    log("Starting QBFC import in 3 seconds...")
    time.sleep(3)

    try:
        from qbfc_import import import_company_via_qbfc
        # Pass qbw_path=None to bind to the already-open company file.
        # Passing the actual path causes QBFC to try launching a new QB instance.
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
