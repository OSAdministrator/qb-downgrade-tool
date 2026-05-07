# QuickBooks Downgrade Tool

A Windows desktop GUI application for automating QuickBooks Desktop **2023 → 2021** migration workflows, including list export/import and CSV-to-IIF transaction conversion.

> ✅ Designed for accountants and service operators who need repeatable batch processing for many company files.

---

## Features

### 1) User-Friendly Desktop GUI (tkinter)
- Multi-file `.QBW` queue
- Per-file admin password support
- JSON password map import
- One-click **Start Processing**
- Real-time status log
- Live progress bar
- Success/failure state per company file

### 2) Automation Engine (modular)
- QB 2023 launch/open hooks
- Export pipeline hooks:
  - IIF list export (accounts/customers/vendors/items/employees)
  - CSV transaction export (Transaction List by Date)
  - Validation report artifact handling
- QB 2021 launch/create/import hooks
- Retry logic + continue-on-error mode
- Full operation logging

### 3) CSV Transaction Parser
- Reads QuickBooks transaction CSV (supports `cp1252` and `latin1` fallback)
- Converts to IIF transaction format:
  - `TRNS`
  - `SPL`
  - `ENDTRNS`
- Handles common transaction types (checks, transfers, deposits, invoices, bills, journals, payroll checks)
- Produces import-ready `transactions_generated.IIF`

### 4) Validation & Reporting
- Validation snapshot model for source vs target
- Generates:
  - `validation_report.xlsx`
  - `validation_report.html`
- Flags discrepancies (PASS/FAIL) for key accounting metrics

### 5) Configurable for Production
- GUI Settings panel + persisted config
- Configurable QuickBooks install paths
- Configurable timeouts
- Output directory preferences
- Retry behavior and dry-run mode

---

## Project Structure

```text
/home/ubuntu/qb_downgrade_tool/
├── qb_downgrade_gui.py         # Main GUI
├── qb_automation.py            # QB automation orchestrator
├── transaction_parser.py       # CSV -> IIF conversion
├── validator.py                # Validation and reports
├── config.py                   # Settings persistence
├── requirements.txt
├── build_exe.bat               # Windows PyInstaller build script
├── sample_config.json
├── README.md
└── docs/
    ├── INSTALLATION.md
    └── screenshots/
        └── README.md
```

---

## Quick Start (Developer)

### 1. Install dependencies
```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows
pip install -r requirements.txt
```

### 2. Launch GUI
```bash
python qb_downgrade_gui.py
```

### 3. Typical workflow
1. Add one or more `.QBW` files
2. Enter passwords (or import password map JSON)
3. Configure settings (paths, output dir, retries)
4. Click **Start Processing**
5. Review results under output folder:
   - `/exports`
   - `/target`
   - `/validation`

---

## Password Map Format

```json
{
  "company_a_23.qbw": "AdminPassword123",
  "company_b_23.qbw": "AnotherPassword"
}
```

---

## Build Standalone .exe (Windows)

Run in CMD or PowerShell from project root:

```bat
build_exe.bat
```

Output executable:

```text
dist\QuickBooksDowngradeTool\QuickBooksDowngradeTool.exe
```

---

## Important Production Notes

- The current code includes a **dry-run mode** for safe testing and non-Windows development.
- Real QuickBooks UI selectors can vary by edition/release/theme. You should tune automation handlers in `qb_automation.py` for your workstation.
- For best reliability:
  - Run in single-user mode where needed
  - Disable popup interruptions
  - Keep consistent screen resolution and Windows scale settings

---

## Validation Strategy

Compare source QB 2023 vs target QB 2021 for:
- Trial Balance total
- Transaction count
- A/R and A/P totals
- List counts (accounts, customers, vendors, items)

Flag any deltas for reconciliation review.

---

## Security & Operational Recommendations

- Use dedicated migration workstations/VMs
- Restrict access to company files and password maps
- Archive logs + validation reports for audit trail
- Run a pilot batch before full 20+ file production run

---

## License

Internal business tooling (proprietary use). Adjust as needed for client distribution.
