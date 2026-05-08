# Installation Guide (Windows 11)

## 0) Clone the Repository (First Time)

Open **PowerShell** or **Command Prompt** and run:

```powershell
cd C:\
git clone https://github.com/OSAdministrator/qb-downgrade-tool.git QBDowngrade
cd QBDowngrade
```

> **Don't have Git?** Download it from https://git-scm.com/download/win and install with default options.

### Pulling Updates

Whenever you need the latest version:

```powershell
cd C:\QBDowngrade
git pull origin main
```

If you have local changes that conflict:

```powershell
git stash
git pull origin main
git stash pop
```

---

## 1) Prerequisites

- Windows 11 Pro
- QuickBooks Desktop 2023 installed
- QuickBooks Desktop 2021 installed
- Python 3.10+
- Local admin rights

## 2) Prepare folders

Recommended:

- `C:\QBDowngrade\Source`
- `C:\QBDowngrade\Results`
- `C:\QBDowngrade\Logs`

## 3) Install dependencies

```powershell
cd C:\path\to\qb_downgrade_tool
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## 4) Configure tool

1. Launch `python qb_downgrade_gui.py`
2. Open **Settings**
3. Verify:
   - QB 2023 executable path
   - QB 2021 executable path
   - Output folder
   - Retry options
4. Keep **Dry-run mode ON** for first smoke test

## 5) First run (safe test)

1. Add 1 test `.QBW` file
2. Enter admin password
3. Click **Start Processing**
4. Confirm output directories and validation reports are created

## 6) Production run

1. Turn **Dry-run mode OFF**
2. Add full batch files
3. Load password map JSON (optional)
4. Run processing during a quiet machine window
5. Review validation report for each company

## 7) Build executable (optional)

```powershell
build_exe.bat
```

Executable path:

`dist\QuickBooksDowngradeTool\QuickBooksDowngradeTool.exe`

## Troubleshooting

- **pywinauto errors**: run command prompt as administrator.
- **Import fails on tax/vendor references**: pre-create missing dependencies in QB 2021 and retry.
- **CSV decode error**: export again from QB with standard CSV format; parser supports cp1252 and latin1 fallback.
- **UI focus issues**: do not interact with keyboard/mouse while automation is running.
