"""Configuration management for QuickBooks TimeWarp\u00ae."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict


# Config lives INSIDE the program folder — portable, no hidden files in ~
_PROGRAM_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = _PROGRAM_DIR / "settings.json"


@dataclass
class TimeoutConfig:
    launch_qb_seconds: int = 90
    open_company_seconds: int = 90
    export_seconds: int = 180
    import_seconds: int = 300
    report_seconds: int = 240
    app_close_seconds: int = 45


@dataclass
class QBInstallPaths:
    qb_2023_path: str = r"C:\Program Files\Intuit\QuickBooks 2023\QBWPremierAccountant.exe"
    qb_2021_path: str = r"C:\Program Files (x86)\Intuit\QuickBooks 2021\QBW32PremierAccountant.exe"


@dataclass
class AppConfig:
    install_paths: QBInstallPaths = field(default_factory=QBInstallPaths)
    timeouts: TimeoutConfig = field(default_factory=TimeoutConfig)
    # Output goes to "Output" subfolder inside the program folder — keeps everything together
    default_output_dir: str = str(_PROGRAM_DIR / "Output")
    retry_attempts: int = 2
    retry_delay_seconds: int = 5
    continue_on_error: bool = True
    dry_run: bool = True

    @classmethod
    def from_dict(cls, payload: Dict[str, Any]) -> "AppConfig":
        install_paths = QBInstallPaths(**payload.get("install_paths", {}))
        timeouts = TimeoutConfig(**payload.get("timeouts", {}))
        return cls(
            install_paths=install_paths,
            timeouts=timeouts,
            default_output_dir=payload.get("default_output_dir", cls().default_output_dir),
            retry_attempts=payload.get("retry_attempts", cls().retry_attempts),
            retry_delay_seconds=payload.get("retry_delay_seconds", cls().retry_delay_seconds),
            continue_on_error=payload.get("continue_on_error", cls().continue_on_error),
            dry_run=payload.get("dry_run", cls().dry_run),
        )

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class ConfigManager:
    """Loads and saves app configuration in JSON format."""

    def __init__(self, path: Path | None = None):
        self.path = path or DEFAULT_CONFIG_PATH

    def load(self) -> AppConfig:
        if not self.path.exists():
            config = AppConfig()
            self.save(config)
            return config

        with self.path.open("r", encoding="utf-8") as f:
            payload = json.load(f)
        return AppConfig.from_dict(payload)

    def save(self, config: AppConfig) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.open("w", encoding="utf-8") as f:
            json.dump(config.to_dict(), f, indent=2)
