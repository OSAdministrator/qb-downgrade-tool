"""Quick re-export: regenerate the snapshot transactions using per-type queries.

Usage:
  1. Open the Gold Coast II 23 file in QB 2023
  2. Run: python test_reexport.py
  3. Then run test_import.py to import into QB 2021
"""
import sys
import time
from pathlib import Path
import json

SNAPSHOT = Path(r"C:\QBDowngrade\Output\joshs gold coast ii 23\exports\company_snapshot.json")

def log(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

def main():
    if not SNAPSHOT.exists():
        log(f"ERROR: Snapshot not found at {SNAPSHOT}")
        sys.exit(1)

    log(f"Snapshot: {SNAPSHOT}")
    log("")
    log("Make sure QB 2023 has 'joshs gold coast ii 23' OPEN.")
    log("This will re-export ONLY the transactions section with line-item detail.")
    log("")
    log("Starting in 3 seconds...")
    time.sleep(3)

    # Load existing snapshot
    with open(SNAPSHOT, 'r', encoding='utf-8') as f:
        snapshot = json.load(f)

    old_tx_count = len(snapshot.get('transactions', []))
    old_has_lines = any('lines' in t for t in snapshot.get('transactions', [])[:5])
    log(f"Existing snapshot: {old_tx_count} transactions, has_lines={old_has_lines}")

    # Connect to QB 2023
    log("Connecting to QB via QBFC...")
    from qbfc_export import open_qbfc_session, _query_typed_transactions

    session = open_qbfc_session(qbw_path=None, log_fn=log)
    try:
        log("Querying transactions by type (with line items)...")
        txns = _query_typed_transactions(
            session, snapshot.get('items', []), log_fn=log
        )
        log(f"")
        log(f"New transactions: {len(txns)} (was {old_tx_count})")

        # Show type breakdown
        type_counts = {}
        for t in txns:
            type_counts[t['type']] = type_counts.get(t['type'], 0) + 1
        for tname, count in sorted(type_counts.items()):
            log(f"  {tname}: {count}")

        # Show sample
        if txns:
            log(f"")
            log(f"Sample txns[0]:")
            log(f"  {json.dumps(txns[0], indent=2)[:500]}")

        # Update snapshot
        snapshot['transactions'] = txns
        with open(SNAPSHOT, 'w', encoding='utf-8') as f:
            json.dump(snapshot, f, indent=2, ensure_ascii=False, default=str)
        log(f"")
        log(f"Snapshot updated: {SNAPSHOT}")
        log(f"  Size: {SNAPSHOT.stat().st_size / 1024:.0f} KB")
    finally:
        session.end()
        log("Session closed.")

    log("Done! Now close QB 2023 and run test_import.py")

if __name__ == "__main__":
    main()
