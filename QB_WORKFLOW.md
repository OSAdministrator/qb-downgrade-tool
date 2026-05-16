# QuickBooks TimeWarp® — Standard Workflow & Commands

## Project Overview
- **Repo:** OSAdministrator/qb-downgrade-tool
- **Branch:** tax-man-mike
- **Purpose:** Automate downgrade of QuickBooks company files from QB 2023 to QB 2021
- **Latest Commit:** bc295ac (Run #15 — OB double-counting fix)

## Environments
### Linux (Agent/Development)
- **Repo Path:** /home/ubuntu/qb-downgrade-tool
- **Purpose:** Code editing, commits, pushes

### Windows VM (Testing/Execution)
- **VM:** 64.13.123.196:4420
- **User:** VM-4420-11\AbacusAgent
- **Password:** 01hello02!@!
- **Repo Path:** C:\QB-TimeWarp\AppFiles
- **Purpose:** Running QB automation, testing

## Standard Commands

### RDP Connection (from Linux)
```bash
sudo apt-get install -y freerdp2-x11
xfreerdp /v:64.13.123.196:4420 /u:'VM-4420-11\AbacusAgent' /p:'01hello02!@!' /dynamic-resolution /cert:ignore /auto-reconnect +clipboard &>/tmp/rdp2.log & echo "RDP PID: $!" ; sleep 10 ; echo "Done"
```

### Cleanup Before Each Run
```powershell
taskkill /F /IM QBW32.exe 2>$null; taskkill /F /IM qbw.exe 2>$null
Stop-Service QuickBooksDB33 -Force -ErrorAction SilentlyContinue; Start-Sleep 3
Remove-Item 'C:\QB-TimeWarp\Working\source' -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item 'C:\QB-TimeWarp\Working\Export' -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item 'C:\QBDowngrade\QB-2021 Template\Tax-Man-Mike-Template.qbw' 'C:\QB-TimeWarp\Working\QB-2021 Template\Tax-Man-Mike-Template.qbw' -Force
Remove-Item 'C:\QB-TimeWarp\Working\QB-2021 Template\*.TLG','C:\QB-TimeWarp\Working\QB-2021 Template\*.ND' -Force -ErrorAction SilentlyContinue
Remove-Item 'C:\QB-TimeWarp\Final Output\*' -Recurse -Force -ErrorAction SilentlyContinue
Start-Service QuickBooksDB33
```

### Run Command
```powershell
cd C:\QB-TimeWarp\AppFiles; python run13.py 2>&1 | Tee-Object -FilePath C:\QB-TimeWarp\run16.log
```

### Git Workflow

#### Linux Side (Agent):
```bash
cd /home/ubuntu/qb-downgrade-tool
# Edit files
git add -A
git commit -m "Fix: Description of change"
git push origin tax-man-mike
```

#### Windows Side:
```powershell
cd C:\QB-TimeWarp\AppFiles
git pull origin tax-man-mike
```

## GitHub Credentials
- **Username:** OSAdministrator
- **PAT:** (stored in agent environment — do not commit tokens to repo)

## Test Environment
- **Source File:** joshs gold coast ii 23.qbw (Desktop)
- **Password:** 3825You171
- **Template:** Tax-Man-Mike-Template.qbw (password: 3825You171)
- **Output:** C:\QB-TimeWarp\Final Output\
- **Working:** C:\QB-TimeWarp\Working\
- **Settings:** C:\QB-TimeWarp\AppFiles\settings.json

## Run Results
- Run #15 (bc295ac): 1967 txns, 31 OB adj, TB=825,607.88 (target=843,234.84)
- Run #14e (d8b7650): 1967 txns, 69 OB adj, TB=1,636,733.24 (double-counted)

## Important Notes
- Always commit changes with clear "Fix:" messages explaining WHY
- Always add inline comments in code explaining bug fixes
- Git workflow is FAST and RELIABLE — use it over file transfers
- Watchdog MUST be PAUSED during post-import _close_qb (resume re-hides window → data loss)
- Notepad = Joseph's intercom — STOP and READ when it appears
- Company name update is MANUAL — Joseph does it himself
