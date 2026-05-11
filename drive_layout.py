"""Drive & partition discovery for QuickBooks TimeWarp® by Our System Administrator.

When TimeWarp is deployed on a flash drive with a fixed partition scheme:
  - Partition 1: (optional) EFI / hidden
  - Partition 2: CODE  — read-only, contains all .py files + launcher
  - Partition 3: WORKSPACE — writable, receives Output/, logs/, settings.json
  - Partition 4: (optional) BACKUP / hidden restore image

The partition numbers are deterministic because we manufacture the drives.
Drive letters are irrelevant — Windows can assign whatever it wants.

Usage:
    from drive_layout import get_workspace_root

    workspace = get_workspace_root()  # Path to writable partition
    # Returns e.g. Path("F:/") if running from a flash drive
    # Returns Path(".") if running from a normal folder (dev mode)
"""

from __future__ import annotations

import subprocess
import json
from pathlib import Path
from typing import Dict, List, Optional, Tuple


# ---------------------------------------------------------------------------
# Configuration — adjust these if partition scheme changes
# ---------------------------------------------------------------------------
CODE_PARTITION_NUMBER = 2       # The partition our code lives on
WORKSPACE_PARTITION_NUMBER = 3  # The writable workspace partition
BACKUP_PARTITION_NUMBER = 4     # Optional backup/restore partition

# Volume labels (optional — used as fallback identification)
CODE_VOLUME_LABEL = "TIMEWARP"      # Label we burn onto the code partition
WORKSPACE_VOLUME_LABEL = "TW_DATA"  # Label we burn onto the workspace partition


# ---------------------------------------------------------------------------
# Core discovery
# ---------------------------------------------------------------------------

def _powershell_json(cmd: str) -> Optional[list]:
    """Run a PowerShell command that outputs JSON and parse it."""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", cmd],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode != 0 or not result.stdout.strip():
            return None
        return json.loads(result.stdout)
    except Exception:
        return None


def _get_disk_number_for_path(file_path: Path) -> Optional[int]:
    """Find which physical disk number a file path lives on.

    Uses PowerShell: resolve drive letter → partition → disk number.
    """
    drive_letter = file_path.resolve().drive.rstrip(":")
    if not drive_letter:
        return None

    cmd = (
        f"Get-Partition -DriveLetter '{drive_letter}' "
        f"| Select-Object DiskNumber, PartitionNumber "
        f"| ConvertTo-Json"
    )
    data = _powershell_json(cmd)
    if not data:
        return None

    # PowerShell may return a single object (not wrapped in array)
    if isinstance(data, dict):
        return data.get("DiskNumber")
    if isinstance(data, list) and data:
        return data[0].get("DiskNumber")
    return None


def _get_partitions_on_disk(disk_number: int) -> List[Dict]:
    """Get all partitions on a specific physical disk.

    Returns list of dicts with: PartitionNumber, DriveLetter, Size, Type,
    IsActive, GptType, MbrType — enough to filter out system partitions.
    """
    cmd = (
        f"Get-Partition -DiskNumber {disk_number} "
        f"| Select-Object PartitionNumber, DriveLetter, Size, Type, "
        f"IsActive, GptType, MbrType "
        f"| ConvertTo-Json"
    )
    data = _powershell_json(cmd)
    if not data:
        return []
    if isinstance(data, dict):
        return [data]
    return data


# GPT partition type GUIDs that are NEVER a valid workspace
_SYSTEM_GPT_TYPES = {
    # EFI System Partition
    "c12a7328-f81f-11d2-ba4b-00a0c93ec93b",
    # Microsoft Reserved (MSR)
    "e3c9e316-0b5c-4db8-817d-f92df00215ae",
    # Windows Recovery Environment
    "de94bba4-06d1-4d40-a16a-bfd50179d6ac",
    # BIOS Boot Partition
    "21686148-6449-6e6f-744e-656564454649",
}

# MBR type IDs that are NEVER a valid workspace
_SYSTEM_MBR_TYPES = {
    0x27,   # Windows Recovery / hidden NTFS
    0xDE,   # Dell utility partition
    0xEF,   # EFI System Partition (MBR)
    0xFE,   # IBM IML / old diagnostic partitions
}


def _is_system_partition(part: Dict) -> bool:
    """Return True if this partition looks like a system/recovery/hidden partition.

    Checks:
      1. Type field: "System", "Reserved", "Recovery", "Unknown" → system
      2. IsActive: False on MBR disks usually means hidden/recovery
         (but on GPT disks IsActive is meaningless, so we combine checks)
      3. GptType GUID: matches known system GUIDs
      4. MbrType ID: matches known system type IDs
      5. No drive letter AND type isn't "Basic" → hidden
    """
    ptype = str(part.get("Type", "")).lower()
    if ptype in ("system", "reserved", "recovery", "unknown"):
        return True

    # Check GPT type GUID
    gpt_type = str(part.get("GptType", "")).lower().strip("{}")
    if gpt_type in _SYSTEM_GPT_TYPES:
        return True

    # Check MBR type ID
    mbr_type = part.get("MbrType")
    if mbr_type is not None:
        try:
            if int(mbr_type) in _SYSTEM_MBR_TYPES:
                return True
        except (ValueError, TypeError):
            pass

    # No drive letter + not a Basic data partition = hidden/system
    if not part.get("DriveLetter") and ptype != "basic":
        return True

    return False


def _find_partition_drive_letter(disk_number: int, partition_number: int) -> Optional[Tuple[str, bool]]:
    """Find the drive letter for a specific partition on a specific disk.

    Returns (drive_letter, is_safe) tuple:
      - drive_letter: the assigned letter (e.g. "F")
      - is_safe: True if the partition is NOT a system/recovery partition
    Returns None if partition not found or has no drive letter.
    """
    partitions = _get_partitions_on_disk(disk_number)
    for p in partitions:
        if p.get("PartitionNumber") == partition_number:
            letter = p.get("DriveLetter")
            if letter:
                safe = not _is_system_partition(p)
                return (str(letter), safe)
    return None


def _find_partition_by_label(label: str) -> Optional[str]:
    """Fallback: find a drive letter by volume label (case-insensitive)."""
    cmd = (
        f"Get-Volume | Where-Object {{ $_.FileSystemLabel -eq '{label}' }} "
        f"| Select-Object DriveLetter | ConvertTo-Json"
    )
    data = _powershell_json(cmd)
    if not data:
        return None
    if isinstance(data, dict):
        return data.get("DriveLetter")
    if isinstance(data, list) and data:
        return data[0].get("DriveLetter")
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_drive_layout() -> Dict[str, Optional[str]]:
    """Discover the full drive layout for the flash drive we're running from.

    Returns:
        {
            "code_drive": "E",       # drive letter of code partition
            "workspace_drive": "F",  # drive letter of workspace partition
            "backup_drive": "G",     # drive letter of backup partition (if any)
            "disk_number": 3,        # physical disk number
            "is_flash_drive": True,  # True if we detected the partition scheme
        }
    """
    program_dir = Path(__file__).resolve().parent
    result = {
        "code_drive": None,
        "workspace_drive": None,
        "backup_drive": None,
        "disk_number": None,
        "is_flash_drive": False,
    }

    # Step 1: What disk are we on?
    disk_num = _get_disk_number_for_path(program_dir)
    if disk_num is None:
        return result

    result["disk_number"] = disk_num
    result["code_drive"] = program_dir.drive.rstrip(":")

    # Step 2: Find sibling partitions on the same physical disk
    partitions = _get_partitions_on_disk(disk_num)
    if len(partitions) < 2:
        # Single partition = normal hard drive / folder deployment
        return result

    # Step 3: Look up workspace partition by number
    ws_result = _find_partition_drive_letter(disk_num, WORKSPACE_PARTITION_NUMBER)
    if ws_result:
        ws_letter, ws_safe = ws_result
        if ws_safe:
            result["workspace_drive"] = ws_letter
            result["is_flash_drive"] = True
        else:
            # Partition exists but it's a system/recovery type — DO NOT USE.
            # This protects against writing to a customer's Recovery partition
            # that happens to be P3 on their disk.
            result["workspace_blocked_reason"] = (
                f"Partition {WORKSPACE_PARTITION_NUMBER} on Disk {disk_num} "
                f"({ws_letter}:) is a system/recovery partition — skipping"
            )
    else:
        # No partition at that number — try by volume label as fallback
        ws_label_letter = _find_partition_by_label(WORKSPACE_VOLUME_LABEL)
        if ws_label_letter:
            result["workspace_drive"] = ws_label_letter
            result["is_flash_drive"] = True

    # Step 4: Optional backup partition
    bk_result = _find_partition_drive_letter(disk_num, BACKUP_PARTITION_NUMBER)
    if bk_result:
        bk_letter, bk_safe = bk_result
        if bk_safe:
            result["backup_drive"] = bk_letter

    return result


def get_workspace_root() -> Path:
    """Get the root path for writable workspace (Output, logs, settings).

    - Flash drive mode: returns the workspace partition root (e.g. F:/)
    - Normal mode: returns the program directory (dev/desktop deployment)

    SAFETY: Only uses a partition as workspace if it contains a
    `.timewarp_workspace` marker file. This prevents accidentally writing
    to a customer's Recovery partition or other random partition that
    happens to be P3 on their disk. When you stamp a flash drive, you
    drop this marker on the workspace partition. No marker = normal mode.

    This is THE function everything else should call.
    """
    try:
        layout = get_drive_layout()
        if layout["is_flash_drive"] and layout["workspace_drive"]:
            ws = Path(f"{layout['workspace_drive']}:/")
            marker = ws / ".timewarp_workspace"
            if not marker.exists():
                # No marker = not our stamped drive, could be a customer's
                # Recovery partition or anything. Fall through to folder mode.
                pass
            else:
                # Marker found — verify it's actually writable
                test_file = ws / ".tw_write_test"
                try:
                    test_file.write_text("ok", encoding="utf-8")
                    test_file.unlink()
                    return ws
                except OSError:
                    pass  # Partition exists but not writable — fall through
    except Exception:
        pass  # PowerShell not available, Linux dev, etc.

    # Fallback: program directory (normal desktop deployment)
    return Path(__file__).resolve().parent


def get_workspace_paths() -> Dict[str, Path]:
    """Get all the writable paths TimeWarp needs.

    Returns:
        {
            "output_dir": Path("F:/Output"),
            "log_dir": Path("F:/logs"),
            "settings_file": Path("F:/settings.json"),
            "workspace_root": Path("F:/"),
        }
    """
    root = get_workspace_root()
    return {
        "output_dir": root / "Final Output",
        "working_dir": root / "working",
        "log_dir": root / "logs",
        "settings_file": root / "settings.json",
        "workspace_root": root,
    }


def print_layout() -> None:
    """Print drive layout for diagnostics / tech troubleshooting."""
    layout = get_drive_layout()
    paths = get_workspace_paths()

    print("=" * 50)
    print("  QuickBooks TimeWarp® — Drive Layout")
    print("=" * 50)

    blocked = layout.get("workspace_blocked_reason")

    if blocked:
        # Found partition but it's a system type — warn the tech
        print(f"  Mode           : Desktop / Folder  (partition blocked)")
        print(f"  Physical Disk  : Disk {layout['disk_number']}")
        print(f"  Code Partition : {layout['code_drive']}:\\")
        print(f"  ⚠ BLOCKED      : {blocked}")
        print(f"  Program Dir    : {Path(__file__).resolve().parent}")
    elif layout["is_flash_drive"]:
        ws_path = Path(f"{layout['workspace_drive']}:/")
        marker = ws_path / ".timewarp_workspace"
        has_marker = marker.exists()
        print(f"  Mode           : Flash Drive")
        print(f"  Physical Disk  : Disk {layout['disk_number']}")
        print(f"  Code Partition : {layout['code_drive']}:\\  (read-only)")
        print(f"  Workspace      : {layout['workspace_drive']}:\\  (writable)")
        print(f"  Marker File    : {'FOUND' if has_marker else 'MISSING — using folder mode'}")
        if not has_marker:
            print(f"")
            print(f"  NOTE: To enable flash drive mode, create an empty file:")
            print(f"         {marker}")
        if layout["backup_drive"]:
            print(f"  Backup         : {layout['backup_drive']}:\\")
    else:
        print(f"  Mode           : Desktop / Folder")
        print(f"  Program Dir    : {Path(__file__).resolve().parent}")

    print(f"")
    print(f"  Output Dir     : {paths['output_dir']}")
    print(f"  Log Dir        : {paths['log_dir']}")
    print(f"  Settings File  : {paths['settings_file']}")
    print("=" * 50)


if __name__ == "__main__":
    print_layout()
