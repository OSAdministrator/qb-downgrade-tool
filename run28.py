"""Run #28 launcher — headless, no GUI, just run_job directly.
Fix: SalesReceipt export captures CustomerRef, import uses ItemRef with generic service item.
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from config import ConfigManager, AppConfig
from qb_automation import QuickBooksAutomationEngine, CompanyJob

SOURCE = Path(r"C:\Users\AbacusAgent\Desktop\joshs gold coast ii 23.qbw")
PASSWORD = "3825You171"

def log(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def main():
    log("="*60)
    log("  QuickBooks TimeWarp\u00ae \u2014 Run #28")
    log(f"  Source: {SOURCE}")
    log("  Fix: SalesReceipt customer linking (CustomerRef export + ItemRef import)")
    log("="*60)

    cfg_path = Path(r"C:\QB-TimeWarp\AppFiles\settings.json")
    cfg = ConfigManager(cfg_path).load()
    
    log(f"  Output dir: {cfg.default_output_dir}")
    log(f"  Working dir: {cfg.default_working_dir}")
    log(f"  Template: {cfg.install_paths.qb_2021_template_path}")
    log(f"  dry_run: {cfg.dry_run}")

    engine = QuickBooksAutomationEngine(cfg)

    job = CompanyJob(
        qbw_path=SOURCE,
        password=PASSWORD,
        output_dir=Path(cfg.default_output_dir),
    )

    start = time.time()
    result = engine.run_job(job, log_fn=log)
    elapsed = time.time() - start

    log("")
    log("="*60)
    log(f"  Result: {'SUCCESS' if result.success else 'FAILED'}")
    log(f"  Message: {result.message}")
    log(f"  Elapsed: {elapsed/60:.1f} minutes")
    if hasattr(result, 'generated_files') and result.generated_files:
        for k, v in result.generated_files.items():
            log(f"  {k}: {v}")
    log("="*60)

if __name__ == "__main__":
    main()
