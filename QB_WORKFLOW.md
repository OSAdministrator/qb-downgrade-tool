# QuickBooks Downgrade Tool - Standard Workflow & Commands

## Project Overview
- **Repo:** OSAdministrator/qb-downgrade-tool
- **Branch:** fix/export-loop-debug
- **Purpose:** Automate downgrade of QuickBooks company files from QB 2023 to QB 2021

## Environments
### Linux (Agent/Development)
- **Repo Path:** /home/ubuntu/github_repos/qb-downgrade-tool
- **Purpose:** Code editing, commits, pushes

### Windows VM (Testing/Execution)
- **VM:** 64.13.123.196:4420
- **User:** VM-4420-11\AbacusAgent
- **Password:** 01hello02!@!
- **Repo Path:** C:\QBDowngradeApp
- **Purpose:** Running QB automation, testing

## Standard Commands

### RDP Connection (from Linux)
```bash
xfreerdp /v:64.13.123.196:4420 /u:'VM-4420-11\AbacusAgent' /p:'01hello02!@!' /dynamic-resolution /cert:ignore /auto-reconnect +clipboard &>/tmp/rdp2.log & echo "RDP PID: $!" ; sleep 10 ; echo "Done"
```

### Git Workflow

#### Linux Side (Agent):
```bash
cd /home/ubuntu/github_repos/qb-downgrade-tool
# Edit files
git add qb_automation.py  # or specific files
git commit -m "Fix: Description of change"
git push origin fix/export-loop-debug
```

#### Windows Side (User):
```powershell
cd C:\QBDowngradeApp
git pull origin fix/export-loop-debug
```

## GitHub Credentials
- **Username:** OSAdministrator
- **PAT:** (stored in agent environment - do not commit tokens to repo)

## Test Environment
- **Test File:** Air-Masters-QB-2023.qbw (on Desktop)
- **Password:** 3825You171
- **Output:** C:\QBDowngrade\LargeFileTest\Target
- **Source:** C:\QBDowngrade\LargeFileTest\Source

## Important Notes
- Always commit changes with clear "Fix:" messages explaining WHY
- Always add inline comments in code explaining bug fixes
- Git workflow is FAST and RELIABLE - use it over file transfers
- User runs git pull commands on Windows - agent doesn't need RDP for code updates
- RDP connection only needed for debugging/observing logs
