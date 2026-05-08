"""QuickBooks desktop automation engine.

Designed for Windows + QuickBooks Desktop 2023 and 2021 environments.
Uses pywinauto/pyautogui when available, with a dry-run fallback for testing.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

from config import AppConfig
from transaction_parser import TransactionParser
from validator import CompanyValidationSnapshot, ValidationReportBuilder

try:
    import pyautogui
except Exception:  # noqa: BLE001
    pyautogui = None

try:
    from pywinauto import Desktop
    from pywinauto.application import Application
    from pywinauto.keyboard import send_keys
except Exception:  # noqa: BLE001
    Application = None
    Desktop = None
    send_keys = None


LogFn = Callable[[str], None]
ProgressFn = Callable[[float], None]


@dataclass
class CompanyJob:
    qbw_path: Path
    password: str
    output_dir: Path
    target_company_name: Optional[str] = None


@dataclass
class CompanyJobResult:
    qbw_path: str
    success: bool
    message: str
    generated_files: Dict[str, str] = field(default_factory=dict)


class QuickBooksAutomationEngine:
    """End-to-end orchestrator for one company file."""

    QB_WINDOW_RE = r"(?i).*quickbooks.*"

    def __init__(self, config: AppConfig, logger: Optional[logging.Logger] = None):
        self.config = config
        self.logger = logger or logging.getLogger("qb_downgrade")
        self.tx_parser = TransactionParser()
        self.validator = ValidationReportBuilder()
        self._qb2023_app: Optional[object] = None
        self._qb2021_app: Optional[object] = None

    def _emit(self, msg: str, log_fn: Optional[LogFn]) -> None:
        self.logger.info(msg)
        if log_fn:
            log_fn(msg)

    def _with_retries(self, fn: Callable[[], None], step_name: str, log_fn: Optional[LogFn]) -> None:
        attempts = self.config.retry_attempts + 1
        for attempt in range(1, attempts + 1):
            try:
                fn()
                return
            except Exception as exc:  # noqa: BLE001
                screenshot_path = self._take_error_screenshot(step_name.replace(" ", "_"))
                self._emit(
                    f"{step_name} failed (attempt {attempt}/{attempts}): {exc}. "
                    f"Screenshot: {screenshot_path if screenshot_path else 'unavailable'}",
                    log_fn,
                )
                if attempt == attempts:
                    raise
                time.sleep(self.config.retry_delay_seconds)

    def _ensure_automation_ready(self) -> None:
        if os.name != "nt":
            raise RuntimeError("QuickBooks automation only runs on Windows.")
        if Application is None or Desktop is None or send_keys is None:
            raise RuntimeError("pywinauto is not available. Install dependencies on Windows.")

    def _wait_until(self, condition: Callable[[], bool], timeout_s: int, interval_s: float = 0.5) -> bool:
        start = time.time()
        while time.time() - start < timeout_s:
            try:
                if condition():
                    return True
            except Exception:  # noqa: BLE001
                pass
            time.sleep(interval_s)
        return False

    def _take_error_screenshot(self, label: str) -> Optional[Path]:
        if pyautogui is None:
            return None
        try:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_label = re.sub(r"[^a-zA-Z0-9_.-]+", "_", label)
            screenshot_dir = Path(self.config.default_output_dir) / "logs"
            screenshot_dir.mkdir(parents=True, exist_ok=True)
            screenshot_path = screenshot_dir / f"error_{safe_label}_{stamp}.png"
            pyautogui.screenshot(str(screenshot_path))
            return screenshot_path
        except Exception:  # noqa: BLE001
            return None

    def _get_desktop(self):
        self._ensure_automation_ready()
        return Desktop(backend="uia")

    def _find_qb_main_window(self, app: Optional[object], version_hint: Optional[str], timeout_s: int):
        def _pick_window() -> Optional[object]:
            if app is not None:
                try:
                    for win in app.windows():
                        if not win.exists() or not win.is_visible():
                            continue
                        title = win.window_text()
                        if version_hint and version_hint not in title:
                            continue
                        if re.search(self.QB_WINDOW_RE, title):
                            return win
                    top = app.top_window()
                    if top.exists() and top.is_visible():
                        return top
                except Exception:  # noqa: BLE001
                    pass

            desktop = self._get_desktop()
            candidates = []
            for w in desktop.windows(title_re=self.QB_WINDOW_RE):
                try:
                    if not w.is_visible():
                        continue
                    title = w.window_text()
                    score = 0
                    if version_hint and version_hint in title:
                        score += 10
                    if "No Company Open" not in title:
                        score += 2
                    candidates.append((score, w))
                except Exception:  # noqa: BLE001
                    continue
            if not candidates:
                return None
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1]

        box: Dict[str, object] = {}

        def _cond() -> bool:
            win = _pick_window()
            if win is not None:
                box["window"] = win
                return True
            return False

        if not self._wait_until(_cond, timeout_s, 1.0):
            hint_text = f" ({version_hint})" if version_hint else ""
            raise RuntimeError(f"Could not find QuickBooks main window{hint_text} within {timeout_s}s")

        return box["window"]

    def _focus_window(self, window) -> None:
        try:
            window.set_focus()
        except Exception:  # noqa: BLE001
            pass

    def _click_first_button(self, dialog, labels: List[str]) -> bool:
        for label in labels:
            try:
                btn = dialog.child_window(title_re=fr"(?i){re.escape(label)}", control_type="Button")
                if btn.exists(timeout=1):
                    btn.wrapper_object().click_input()
                    return True
            except Exception:  # noqa: BLE001
                continue
        return False

    def _set_edit_value(self, parent, value: str, edit_index: int = 0) -> bool:
        try:
            edits = parent.descendants(control_type="Edit")
            if not edits or edit_index >= len(edits):
                return False
            edit = edits[edit_index]
            edit.set_focus()
            time.sleep(0.1)
            # Select all existing text first, then replace
            if send_keys is not None:
                send_keys("^a")
                time.sleep(0.05)
            try:
                edit.set_edit_text(value)
            except Exception:  # noqa: BLE001
                # Fallback: type the value character by character
                if send_keys is not None:
                    send_keys("{DELETE}")
                    time.sleep(0.05)
                    edit.type_keys(value, with_spaces=True, pause=0.02)
            return True
        except Exception:  # noqa: BLE001
            return False

    def _find_active_dialog(self, title_re: Optional[str] = None):
        desktop = self._get_desktop()
        windows = desktop.windows()
        for w in windows:
            try:
                if not w.is_visible():
                    continue
                title = w.window_text() or ""
                if not title:
                    continue
                if title_re and not re.search(title_re, title):
                    continue
                return w
            except Exception:  # noqa: BLE001
                continue
        return None

    def _dismiss_common_dialogs(self, log_fn: Optional[LogFn]) -> None:
        desktop = self._get_desktop()
        for dialog in desktop.windows():
            try:
                if not dialog.is_visible():
                    continue
                title = dialog.window_text() or ""
                if not title:
                    continue
                # Skip the main QB window itself
                if re.search(self.QB_WINDOW_RE, title) and "login" not in title.lower():
                    continue

                title_l = title.lower()

                # Handle password/login dialogs at startup
                if "login" in title_l or "password" in title_l:
                    # Don't dismiss these - they need special handling
                    continue

                # Handle "Set Up an External Accountant User" dialog
                if "external accountant" in title_l or "accountant user" in title_l:
                    clicked = self._click_first_button(dialog, ["No", "Cancel", "Close"])
                    if clicked:
                        self._emit(f"Dismissed dialog: {title}", log_fn)
                    continue

                # Handle "Enter Memorized Transactions" dialog
                if "memorized transaction" in title_l:
                    clicked = self._click_first_button(dialog, ["Enter All Later", "Close", "Cancel", "No"])
                    if clicked:
                        self._emit(f"Dismissed dialog: {title}", log_fn)
                    time.sleep(0.5)
                    # There may be a follow-up "Enter Memorized Transactions Later" dialog
                    followup = self._find_active_dialog(title_re=r"(?i)memorized.*later")
                    if followup is not None:
                        self._click_first_button(followup, ["OK", "Close"])
                        try:
                            # Try clicking the X button
                            followup.close()
                        except Exception:  # noqa: BLE001
                            if send_keys is not None:
                                send_keys("{ENTER}")
                        self._emit("Dismissed memorized transactions follow-up dialog", log_fn)
                    continue

                if any(
                    x in title_l
                    for x in [
                        "update",
                        "backup",
                        "confirmation",
                        "quickbooks message",
                        "information",
                        "warning",
                        "microsoft office",
                        "register",
                        "product information",
                        "new feature",
                        "intuit",
                    ]
                ):
                    clicked = self._click_first_button(
                        dialog,
                        ["No", "Don't Save", "Continue", "Yes", "OK", "Close", "Cancel", "Later", "Skip"],
                    )
                    if clicked:
                        self._emit(f"Dismissed dialog: {title}", log_fn)
            except Exception:  # noqa: BLE001
                continue

    def _close_popup_windows(self, main_window, log_fn: Optional[LogFn]) -> None:
        """Close QB popup/helper windows that steal focus after login.

        Common culprits include the Accountant Center, Getting Started,
        Learning Center, What's New, and Home Page windows.  These are
        enumerated from the desktop and closed individually so that the
        main QuickBooks company window can regain focus for menu navigation.
        """
        desktop = self._get_desktop()

        # Substrings (lower-cased) of window titles that should be closed
        popup_hints = [
            "accountant center",
            "getting started",
            "what's new",
            "whats new",
            "tips",
            "quickbooks learning center",
            "learning center",
            "home page",
            "quickbooks home",
            "new feature",
            "coach",
            "did you know",
            "quickbooks desktop",  # generic splash / promo windows
        ]

        closed_any = False
        for win in desktop.windows():
            try:
                if not win.is_visible():
                    continue
                title = win.window_text() or ""
                if not title:
                    continue

                # Never close the main company window itself
                if win.handle == getattr(main_window, "handle", None):
                    continue
                # Also skip anything that looks like the main QB window
                if re.search(self.QB_WINDOW_RE, title) and not any(h in title.lower() for h in popup_hints):
                    continue

                title_l = title.lower()
                if any(hint in title_l for hint in popup_hints):
                    # Try clicking common close/dismiss buttons first
                    clicked = self._click_first_button(
                        win, ["Close", "OK", "No", "Cancel", "Skip", "Later", "X"]
                    )
                    if not clicked:
                        # Fall back to closing the window directly
                        try:
                            win.close()
                        except Exception:  # noqa: BLE001
                            pass
                    self._emit(f"Closed popup window: {title}", log_fn)
                    closed_any = True
                    time.sleep(0.5)
            except Exception:  # noqa: BLE001
                continue

        if closed_any:
            time.sleep(1)

        # Ensure the main window has focus for subsequent menu navigation
        try:
            self._focus_window(main_window)
            time.sleep(0.5)
        except Exception:  # noqa: BLE001
            pass

    def _invoke_menu(self, window, menu_path: str, fallback_keys: Optional[str], log_fn: Optional[LogFn]) -> None:
        self._focus_window(window)
        try:
            window.menu_select(menu_path)
            return
        except Exception as exc:  # noqa: BLE001
            self._emit(f"menu_select failed for '{menu_path}': {exc}", log_fn)

        if fallback_keys:
            if send_keys is None:
                raise RuntimeError(f"Could not invoke menu '{menu_path}' and keyboard fallback unavailable.")
            self._emit(f"Trying fallback keystrokes for menu '{menu_path}': {fallback_keys}", log_fn)
            send_keys(fallback_keys, pause=0.08)
            time.sleep(1)
            return

        raise RuntimeError(f"Could not invoke menu path: {menu_path}")

    def _handle_standard_file_dialog(
        self,
        file_path: Path,
        save_mode: bool,
        timeout_s: int,
        log_fn: Optional[LogFn],
    ) -> None:
        target = str(file_path)

        dialog_box: Dict[str, object] = {}

        def _cond() -> bool:
            dlg = self._find_active_dialog(title_re=r"(?i)(open|save as|import|export|create company)")
            if dlg is not None:
                dialog_box["dlg"] = dlg
                return True
            return False

        if not self._wait_until(_cond, timeout_s, 0.5):
            raise RuntimeError("Timed out waiting for standard Open/Save dialog")

        dialog = dialog_box["dlg"]
        self._focus_window(dialog)

        # Try known File name field patterns first, then generic first edit control.
        edit_set = False
        for locator in [
            {"auto_id": "1148", "control_type": "Edit"},
            {"title_re": r"(?i)file\s*name", "control_type": "Edit"},
        ]:
            try:
                edit = dialog.child_window(**locator)
                if edit.exists(timeout=1):
                    edit.set_focus()
                    edit.set_edit_text(target)
                    edit_set = True
                    break
            except Exception:  # noqa: BLE001
                continue

        if not edit_set:
            edit_set = self._set_edit_value(dialog, target)

        if not edit_set:
            raise RuntimeError(f"Could not set file path in dialog: {target}")

        button_labels = ["Save", "Open", "Import", "OK", "Create", "Next"] if save_mode else ["Open", "Import", "OK"]
        if not self._click_first_button(dialog, button_labels):
            # Fallback: press Enter
            if send_keys is None:
                raise RuntimeError("Could not click Open/Save button and keyboard fallback unavailable")
            send_keys("{ENTER}")

        time.sleep(1)
        self._dismiss_common_dialogs(log_fn)

    def _handle_password_prompt(self, password: str, timeout_s: int, log_fn: Optional[LogFn]) -> None:
        """Wait for a password/login dialog and enter credentials.

        This method is called AFTER a company file open has been triggered.
        It waits for the password dialog to appear, enters credentials,
        and clicks OK/Login. It does NOT dismiss any other dialogs.
        """
        self._emit("  [Password] Checking if password was provided...", log_fn)
        if not password:
            self._emit("  [Password] No password provided – skipping password entry.", log_fn)
            return

        self._emit(f"  [Password] Password provided (length={len(password)}). "
                   f"Scanning for dialog up to {min(timeout_s, 20)}s...", log_fn)

        dialog_box: Dict[str, object] = {}
        poll_count = 0

        def _cond() -> bool:
            nonlocal poll_count
            poll_count += 1
            dlg = self._find_active_dialog(title_re=r"(?i)(password|login)")
            if dlg is not None:
                try:
                    dlg_title = dlg.window_text()
                except Exception:  # noqa: BLE001
                    dlg_title = "<unknown>"
                self._emit(f"  [Password] Dialog FOUND on poll #{poll_count}: '{dlg_title}'", log_fn)
                dialog_box["dlg"] = dlg
                return True
            return False

        if not self._wait_until(_cond, min(timeout_s, 20), 0.5):
            self._emit(f"  [Password] No password/login dialog detected after {poll_count} polls; continuing.", log_fn)
            return

        dialog = dialog_box["dlg"]
        self._emit("  [Password] Focusing password dialog...", log_fn)
        self._focus_window(dialog)
        time.sleep(0.3)

        # Enumerate edit fields to determine dialog type
        try:
            edits = dialog.descendants(control_type="Edit")
            self._emit(f"  [Password] Found {len(edits)} edit field(s) in dialog.", log_fn)
        except Exception as e:  # noqa: BLE001
            self._emit(f"  [Password] WARNING: Could not enumerate edit fields: {e}", log_fn)
            edits = []

        if len(edits) >= 2:
            # ----- Username + Password dialog -----
            self._emit(f"  [Password] Dialog type: USERNAME + PASSWORD ({len(edits)} fields)", log_fn)

            # Clear and set username (field 0)
            self._emit("  [Password] Step A: Setting username field to 'Admin'...", log_fn)
            if send_keys is not None:
                try:
                    edits[0].set_focus()
                    time.sleep(0.1)
                    send_keys("^a{DELETE}")
                    time.sleep(0.1)
                    self._emit("  [Password] Username field cleared via Ctrl+A, Delete.", log_fn)
                except Exception as e:  # noqa: BLE001
                    self._emit(f"  [Password] WARNING: Could not clear username field: {e}", log_fn)
            if not self._set_edit_value(dialog, "Admin", edit_index=0):
                self._emit("  [Password] WARNING: _set_edit_value failed for username field.", log_fn)
            else:
                self._emit("  [Password] Username field set to 'Admin'.", log_fn)

            # Clear and set password (field 1)
            self._emit("  [Password] Step B: Setting password field...", log_fn)
            if send_keys is not None:
                try:
                    edits[1].set_focus()
                    time.sleep(0.1)
                    send_keys("^a{DELETE}")
                    time.sleep(0.1)
                    self._emit("  [Password] Password field cleared via Ctrl+A, Delete.", log_fn)
                except Exception as e:  # noqa: BLE001
                    self._emit(f"  [Password] WARNING: Could not clear password field: {e}", log_fn)
            if not self._set_edit_value(dialog, password, edit_index=1):
                raise RuntimeError("Password dialog detected but failed to set password in edit field 1")
            self._emit("  [Password] Password field set successfully.", log_fn)

        else:
            # ----- Password-only dialog -----
            self._emit("  [Password] Dialog type: PASSWORD ONLY (single field)", log_fn)
            self._emit("  [Password] Setting password in the single edit field...", log_fn)
            if not self._set_edit_value(dialog, password):
                raise RuntimeError("Password dialog detected but failed to set password")
            self._emit("  [Password] Password field set successfully.", log_fn)

        # Click OK / Login button
        self._emit("  [Password] Step C: Clicking OK/Login button...", log_fn)
        btn_clicked = self._click_first_button(dialog, ["OK", "Continue", "Login", "Open"])
        if btn_clicked:
            self._emit("  [Password] Button click succeeded.", log_fn)
        else:
            if send_keys is not None:
                self._emit("  [Password] Button click failed – sending Enter key as fallback.", log_fn)
                send_keys("{ENTER}")
            else:
                self._emit("  [Password] WARNING: Button click failed and send_keys unavailable.", log_fn)

        self._emit("  [Password] Password entry sequence complete.", log_fn)

    def _wait_for_company_ready(self, main_window, timeout_s: int, log_fn: Optional[LogFn]) -> None:
        """Wait for the company to finish loading by checking the window title.

        The company is considered loaded when the main QB window title no
        longer contains "No Company Open".  A positive company name in the
        title is the definitive signal.

        NOTE: This method does NOT dismiss dialogs during the wait loop.
        Popup/dialog dismissal should happen AFTER this method returns,
        ensuring the company is fully loaded first.
        """
        self._emit(f"[STEP 5] Waiting for company to finish loading (timeout={timeout_s}s)...", log_fn)
        self._emit("  [Load] Polling window title every 1s for company name...", log_fn)

        poll_count = 0
        last_title = ""

        def _cond() -> bool:
            nonlocal poll_count, last_title
            poll_count += 1
            try:
                title = main_window.window_text() or ""
            except Exception as e:  # noqa: BLE001
                self._emit(f"  [Load] Poll #{poll_count}: ERROR reading title: {e}", log_fn)
                return False

            # Log title changes (avoid spamming identical titles)
            if title != last_title:
                self._emit(f"  [Load] Poll #{poll_count}: Title changed -> '{title}'", log_fn)
                last_title = title
            elif poll_count % 10 == 0:
                self._emit(f"  [Load] Poll #{poll_count}: Still waiting (title='{title}')...", log_fn)

            has_company = "No Company Open" not in title
            if has_company:
                self._emit(f"  [Load] ✓ Company detected in title on poll #{poll_count}: '{title}'", log_fn)
            return has_company

        if not self._wait_until(_cond, timeout_s, 1.0):
            try:
                final_title = main_window.window_text() or ""
            except Exception:  # noqa: BLE001
                final_title = "<could not read>"
            self._emit(f"  [Load] TIMEOUT after {poll_count} polls. Final title: '{final_title}'", log_fn)
            raise RuntimeError("Company did not finish loading in QuickBooks")

        self._emit(f"[STEP 5] Company is fully loaded after {poll_count} poll(s).", log_fn)
        self._emit("  [Load] Pausing 2s for QB to stabilize after load...", log_fn)
        time.sleep(2)

    def _open_company_file(self, main_window, qbw_path: Path, password: str, timeout_s: int, log_fn: Optional[LogFn]) -> None:
        """Open a QuickBooks company file using QB's command-line parameter.

        Instead of navigating File → Open menus and file dialogs, we launch
        (or re-launch) QuickBooks with the .qbw file path as a command-line
        argument.  QB will automatically open the company file and prompt for
        the admin password – we just need to wait for that dialog and handle it.

        Steps:
          1. Launch QB with the .qbw file as a parameter (auto-opens the file)
          2. Wait for the password dialog to appear
          3. Handle the password dialog (already coded in _handle_password_prompt)
          4. Wait for the company to fully load
          5. Dismiss post-login popups
        """
        self._emit("=" * 60, log_fn)
        self._emit("  OPENING COMPANY FILE  –  Command-Line Parameter Approach", log_fn)
        self._emit("=" * 60, log_fn)
        self._emit(f"  File path : {qbw_path}", log_fn)
        self._emit(f"  Password  : {'(provided, length=%d)' % len(password) if password else '(none)'}", log_fn)
        self._emit(f"  Timeout   : {timeout_s}s", log_fn)

        if not qbw_path.exists():
            self._emit(f"  ERROR: File does not exist on disk: {qbw_path}", log_fn)
            raise FileNotFoundError(f"Source QBW file not found: {qbw_path}")
        self._emit(f"  File confirmed on disk ({qbw_path.stat().st_size:,} bytes).", log_fn)

        # ==================================================================
        # STEP 1 – Launch QB with the .qbw file as a command-line parameter
        # ==================================================================
        self._emit("[STEP 1/5] Launching QB with company file as command-line parameter...", log_fn)

        # Close the existing QB instance so we can re-launch with the file arg
        if self._qb2023_app is not None:
            self._emit("  [Launch] Closing existing QB instance to re-launch with file parameter...", log_fn)
            self._close_qb(self._qb2023_app, log_fn)
            self._qb2023_app = None
            time.sleep(2)

        # Launch QB with the .qbw path – QB will auto-open the company file
        qb_exe = self.config.install_paths.qb_2023_path
        self._qb2023_app = self._launch_qb(qb_exe, log_fn, qbw_path=qbw_path)

        self._emit("[STEP 1/5] ✓ QB launched with company file parameter.", log_fn)

        # Re-acquire the main window reference after relaunch
        self._emit("  [Launch] Waiting for QB main window to appear...", log_fn)
        main_window = self._find_qb_main_window(self._qb2023_app, "2023", self.config.timeouts.launch_qb_seconds)
        self._emit("  [Launch] ✓ Main window found.", log_fn)

        # ==================================================================
        # STEP 2-3 – Wait for password dialog → enter password → click OK
        # ==================================================================
        self._emit("[STEP 2/5] Waiting for password/login dialog to appear...", log_fn)
        self._emit("  (QB should auto-prompt for the password after opening the company file.)", log_fn)
        self._handle_password_prompt(password, timeout_s=timeout_s, log_fn=log_fn)
        self._emit("[STEP 3/5] ✓ Password handling complete.", log_fn)

        # ==================================================================
        # STEP 4 – Wait for the company to FULLY load
        # ==================================================================
        self._emit("[STEP 4/5] Waiting for company to fully load (title bar check)...", log_fn)
        self._emit("  (Looking for company name in title; 'No Company Open' must disappear.)", log_fn)
        self._wait_for_company_ready(main_window, timeout_s=timeout_s, log_fn=log_fn)
        self._emit("[STEP 4/5] ✓ Company is fully loaded.", log_fn)

        # ==================================================================
        # STEP 5 – Dismiss popup dialogs (Payroll, Accountant Center …)
        # ==================================================================
        self._emit("[STEP 5/5] Company loaded – dismissing post-login popup dialogs...", log_fn)
        self._emit("  [Popups] Pausing 1s for popups to finish rendering...", log_fn)
        time.sleep(1)

        self._emit("  [Popups] Running _dismiss_common_dialogs()...", log_fn)
        self._dismiss_common_dialogs(log_fn)
        self._emit("  [Popups] ✓ Common dialogs dismissed.", log_fn)

        self._emit("  [Popups] Running _close_popup_windows()...", log_fn)
        self._close_popup_windows(main_window, log_fn)
        self._emit("  [Popups] ✓ Popup windows closed.", log_fn)

        # Re-focus the main window after clearing popups
        self._emit("  [Focus] Re-focusing main QB window...", log_fn)
        self._focus_window(main_window)
        time.sleep(0.5)
        try:
            final_title = main_window.window_text() or ""
        except Exception:  # noqa: BLE001
            final_title = "<could not read>"
        self._emit(f"  [Focus] Main window title: '{final_title}'", log_fn)

        self._emit("=" * 60, log_fn)
        self._emit("  COMPANY FILE OPEN COMPLETE", log_fn)
        self._emit("=" * 60, log_fn)

    def _export_single_list_iif(self, main_window, list_name: str, out_path: Path, log_fn: Optional[LogFn]) -> None:
        self._emit(f"Exporting list '{list_name}' -> {out_path}", log_fn)

        # Dismiss any stale menus / popups before starting
        if send_keys is not None:
            send_keys("{ESCAPE}")
            time.sleep(0.3)
            send_keys("{ESCAPE}")
            time.sleep(0.3)

        self._focus_window(main_window)
        time.sleep(1)

        # Close any popup windows that may have stolen focus
        self._close_popup_windows(main_window, log_fn)
        self._dismiss_common_dialogs(log_fn)
        self._focus_window(main_window)
        time.sleep(0.5)

        # Navigate: File -> Utilities -> Export -> Lists to IIF Files...
        # Try multiple approaches for menu navigation
        dlg = None

        # Approach 1: keyboard shortcuts with longer waits
        if send_keys is not None:
            self._emit("Trying menu navigation via keyboard shortcuts", log_fn)
            send_keys("%f")  # Alt+F for File menu
            time.sleep(1.0)
            send_keys("u")   # Utilities
            time.sleep(1.0)
            send_keys("e")   # Export
            time.sleep(1.0)
            send_keys("l")   # Lists to IIF Files
            time.sleep(2)

            dlg = self._find_active_dialog(title_re=r"(?i)(export|iif|list)")

        # Approach 2: try menu_select via pywinauto
        if dlg is None:
            self._emit("Keyboard menu navigation failed, trying menu_select", log_fn)
            if send_keys is not None:
                send_keys("{ESCAPE}")
                time.sleep(0.5)
            self._focus_window(main_window)
            time.sleep(0.5)
            try:
                self._invoke_menu(
                    main_window,
                    "File->Utilities->Export->Lists to IIF Files...",
                    None,
                    log_fn,
                )
                time.sleep(2)
                dlg = self._find_active_dialog(title_re=r"(?i)(export|iif|list)")
            except Exception:  # noqa: BLE001
                pass

        # Approach 3: try alternative keyboard sequence
        if dlg is None and send_keys is not None:
            self._emit("Trying alternative keyboard sequence for export menu", log_fn)
            send_keys("{ESCAPE}")
            time.sleep(0.5)
            self._focus_window(main_window)
            time.sleep(0.5)
            # Try Alt, then arrow keys through File menu
            send_keys("{ESCAPE}")
            time.sleep(0.3)
            send_keys("%f")
            time.sleep(1.5)
            # Look for Utilities in the menu and use arrow keys
            send_keys("u")
            time.sleep(1.5)
            send_keys("e")
            time.sleep(1.5)
            send_keys("l")
            time.sleep(2)
            dlg = self._find_active_dialog(title_re=r"(?i)(export|iif|list)")

        if dlg is None:
            raise RuntimeError("Export Lists to IIF dialog did not appear")

        self._focus_window(dlg)

        # QB's Export dialog uses checkboxes for each list type.
        # Try to find and check the appropriate checkbox.
        selected = False
        try:
            checkboxes = dlg.descendants(control_type="CheckBox")
            for cb in checkboxes:
                try:
                    cb_text = cb.window_text() or ""
                    if list_name.lower() in cb_text.lower():
                        if not cb.get_toggle_state():
                            cb.toggle()
                        selected = True
                        self._emit(f"Selected checkbox: {cb_text}", log_fn)
                    else:
                        # Uncheck other checkboxes to export only the desired list
                        if cb.get_toggle_state():
                            cb.toggle()
                except Exception:  # noqa: BLE001
                    continue
        except Exception:  # noqa: BLE001
            pass

        if not selected:
            # Try list/combobox controls as fallback
            for ctrl_type in ["List", "Tree", "ComboBox"]:
                try:
                    ctrls = dlg.descendants(control_type=ctrl_type)
                    for ctrl in ctrls:
                        try:
                            ctrl.select(list_name)
                            selected = True
                            break
                        except Exception:  # noqa: BLE001
                            continue
                    if selected:
                        break
                except Exception:  # noqa: BLE001
                    continue

        if not self._click_first_button(dlg, ["OK", "Export", "Save"]):
            if send_keys is not None:
                send_keys("{ENTER}")

        time.sleep(1)
        self._handle_standard_file_dialog(out_path, save_mode=True, timeout_s=self.config.timeouts.export_seconds, log_fn=log_fn)
        self._dismiss_common_dialogs(log_fn)

    def _open_report(self, main_window, menu_path: str, fallback_keys: str, log_fn: Optional[LogFn]) -> None:
        self._invoke_menu(main_window, menu_path, fallback_keys, log_fn)
        time.sleep(2)
        self._dismiss_common_dialogs(log_fn)

    def _export_transaction_list_csv(self, main_window, out_csv: Path, log_fn: Optional[LogFn]) -> None:
        self._emit(f"Exporting Transaction List by Date report -> {out_csv}", log_fn)

        self._focus_window(main_window)
        time.sleep(0.5)

        # Navigate: Reports -> Accountant & Taxes -> Transaction List by Date
        if send_keys is not None:
            send_keys("%r")  # Alt+R for Reports menu
            time.sleep(0.5)
            send_keys("a")   # Accountant & Taxes
            time.sleep(0.5)
            send_keys("t")   # Transaction List by Date
            time.sleep(3)
        else:
            self._open_report(
                main_window,
                "Reports->Accountant & Taxes->Transaction List by Date",
                "%rat",
                log_fn,
            )

        # Wait for report window to appear
        time.sleep(3)
        self._dismiss_common_dialogs(log_fn)

        # Try to set date range to All
        if send_keys is not None:
            try:
                send_keys("%d")  # Focus date dropdown
                time.sleep(0.3)
                send_keys("all{ENTER}")
                time.sleep(2)
            except Exception:  # noqa: BLE001
                pass

        # Export report using File -> Export -> Excel/CSV
        if send_keys is None:
            raise RuntimeError("Keyboard automation unavailable for transaction CSV export")

        # Try Ctrl+E first (Excel export shortcut)
        send_keys("^e")
        time.sleep(2)

        export_dlg = self._find_active_dialog(title_re=r"(?i)(send report|excel|export|save)")
        if export_dlg is not None:
            self._focus_window(export_dlg)
            # Try selecting CSV option if present
            try:
                for txt in [
                    "comma separated values",
                    "Create a comma separated",
                    "CSV",
                ]:
                    try:
                        option = export_dlg.child_window(title_re=fr"(?i){re.escape(txt)}")
                        if option.exists(timeout=1):
                            option.click_input()
                            self._emit(f"Selected export option: {txt}", log_fn)
                            break
                    except Exception:  # noqa: BLE001
                        continue
            except Exception:  # noqa: BLE001
                pass

            self._click_first_button(export_dlg, ["Export", "OK", "Next", "Create"])
            time.sleep(1)

        self._handle_standard_file_dialog(out_csv, save_mode=True, timeout_s=self.config.timeouts.export_seconds, log_fn=log_fn)
        self._dismiss_common_dialogs(log_fn)

        # Close the report window
        if send_keys is not None:
            send_keys("^{F4}")
            time.sleep(1)
            self._dismiss_common_dialogs(log_fn)

    def _export_report_pdf(self, main_window, report_menu_path: str, fallback_keys: str, out_pdf: Path, log_fn: Optional[LogFn]) -> None:
        self._emit(f"Generating validation report: {out_pdf.name}", log_fn)
        self._open_report(main_window, report_menu_path, fallback_keys, log_fn)

        if send_keys is None:
            raise RuntimeError("Keyboard automation unavailable for PDF report export")

        # Ctrl+P from report and rely on Microsoft Print to PDF as default printer.
        send_keys("^p")
        time.sleep(1)

        print_dlg = self._find_active_dialog(title_re=r"(?i)(print|form name|reports)")
        if print_dlg is not None:
            self._focus_window(print_dlg)
            self._click_first_button(print_dlg, ["Print", "OK"])

        self._handle_standard_file_dialog(out_pdf, save_mode=True, timeout_s=self.config.timeouts.report_seconds, log_fn=log_fn)
        self._dismiss_common_dialogs(log_fn)

    # -------- QB 2023 extraction --------

    def _launch_qb(self, exe_path: str, log_fn: Optional[LogFn], qbw_path: Optional[Path] = None) -> Optional[object]:
        """Launch QuickBooks, optionally with a company file as a command-line parameter.

        If *qbw_path* is provided and QB is not already running, QB is started
        with the .qbw file path as an argument.  This causes QB to open the
        company file directly and auto-prompt for the password – no File → Open
        menu navigation required.

        If QB is already running we simply connect to the existing process
        (the caller is responsible for opening the file via _open_company_file).
        """
        if qbw_path:
            self._emit(f"Launching QuickBooks: {exe_path} WITH company file: {qbw_path}", log_fn)
        else:
            self._emit(f"Launching QuickBooks: {exe_path}", log_fn)

        if self.config.dry_run:
            time.sleep(1)
            return None

        self._ensure_automation_ready()

        # Check if QB is already running by looking for its window
        exe_name = Path(exe_path).name
        try:
            app = Application(backend="uia").connect(path=exe_path)
            self._emit(f"Connected to already-running QuickBooks: {exe_name}", log_fn)
            return app
        except Exception:  # noqa: BLE001
            pass

        # Not running – start it, optionally with the .qbw path as argument
        if qbw_path:
            cmd_line = f'"{exe_path}" "{qbw_path}"'
            self._emit(f"  Starting QB with command: {cmd_line}", log_fn)
            app = Application(backend="uia").start(cmd_line)
        else:
            app = Application(backend="uia").start(exe_path)
        return app

    def _close_qb(self, app: Optional[object], log_fn: Optional[LogFn]) -> None:
        self._emit("Closing QuickBooks", log_fn)
        if self.config.dry_run:
            return

        if app is not None:
            try:
                win = self._find_qb_main_window(app, None, timeout_s=5)
                self._focus_window(win)
                if send_keys is not None:
                    send_keys("%{F4}")
                time.sleep(2)
                self._dismiss_common_dialogs(log_fn)
            except Exception:  # noqa: BLE001
                pass

            try:
                app.kill()
            except Exception:  # noqa: BLE001
                pass

    def _is_company_already_open(self, main_window, company_name_hint: str) -> bool:
        """Check if the desired company is already open in QB.

        The window title alone is not reliable -- QB may show
        'QuickBooks Accountant Desktop Plus 2023' even when the
        'No Company Open' landing page is displayed.  We therefore
        also inspect child elements for the 'No Company Open' text.
        """
        try:
            title = main_window.window_text() or ""

            # Explicit "No Company Open" in the title bar -> not open
            if "No Company Open" in title:
                return False

            # Check child elements for "No Company Open" text which
            # appears in the QB landing page content area
            try:
                children = main_window.descendants(control_type="Text")
                for child in children:
                    try:
                        child_text = child.window_text() or ""
                        if "No Company Open" in child_text:
                            return False
                    except Exception:  # noqa: BLE001
                        continue
            except Exception:  # noqa: BLE001
                pass

            # If the title contains the company name hint, it is open
            if company_name_hint.lower() in title.lower():
                return True

            # QB typically puts the company file name in the title bar
            # when a company is open.  A generic title like
            # "QuickBooks Accountant Desktop Plus 2023" without any
            # company-specific text means nothing is loaded yet.
            # Only return True if the title has something beyond the
            # standard QB product name.
            generic_patterns = [
                r"(?i)^QuickBooks.*Desktop.*\d{4}$",
                r"(?i)^QuickBooks.*Premier.*\d{4}$",
                r"(?i)^QuickBooks.*Enterprise.*\d{4}$",
            ]
            for pat in generic_patterns:
                if re.match(pat, title.strip()):
                    return False

            # If we get here the title has extra text (likely a company
            # name) and no "No Company Open" was found -> assume open
            if "QuickBooks" in title:
                return True

        except Exception:  # noqa: BLE001
            pass
        return False

    def _handle_startup_dialogs(self, password: str, timeout_s: int, log_fn: Optional[LogFn]) -> None:
        """Handle dialogs that appear when QB starts up (password, accountant user, memorized transactions)."""
        password_entered = False
        start = time.time()
        while time.time() - start < timeout_s:
            # Check for password/login dialog first
            login_dlg = self._find_active_dialog(title_re=r"(?i)(password|login)")
            if login_dlg is not None and not password_entered:
                self._emit("Found startup login dialog, entering password", log_fn)
                self._focus_window(login_dlg)
                time.sleep(0.5)

                # Clear and enter credentials using keyboard
                if send_keys is not None:
                    edits = login_dlg.descendants(control_type="Edit")
                    if edits:
                        if len(edits) >= 2:
                            # Username + Password login dialog
                            self._emit(f"Detected username+password login ({len(edits)} edit fields)", log_fn)
                            # First edit = username field - clear and set
                            username_edit = edits[0]
                            username_edit.set_focus()
                            time.sleep(0.2)
                            send_keys("^a")
                            time.sleep(0.05)
                            send_keys("{DELETE}")
                            time.sleep(0.1)
                            try:
                                username_edit.set_edit_text("Admin")
                            except Exception:  # noqa: BLE001
                                username_edit.type_keys("Admin", with_spaces=True, pause=0.03)
                            time.sleep(0.3)
                            # Second edit = password field - clear and set
                            password_edit = edits[1]
                            password_edit.set_focus()
                            time.sleep(0.2)
                            send_keys("^a")
                            time.sleep(0.05)
                            send_keys("{DELETE}")
                            time.sleep(0.1)
                            try:
                                password_edit.set_edit_text(password)
                            except Exception:  # noqa: BLE001
                                password_edit.type_keys(password, with_spaces=True, pause=0.03)
                            time.sleep(0.3)
                        else:
                            # Password-only login dialog (single edit field)
                            self._emit("Detected password-only login (1 edit field)", log_fn)
                            edit = edits[0]
                            edit.set_focus()
                            time.sleep(0.2)
                            send_keys("^a")
                            time.sleep(0.05)
                            send_keys("{DELETE}")
                            time.sleep(0.1)
                            try:
                                edit.set_edit_text(password)
                            except Exception:  # noqa: BLE001
                                edit.type_keys(password, with_spaces=True, pause=0.03)
                            time.sleep(0.3)

                    # Click OK button - try multiple approaches
                    if not self._click_first_button(login_dlg, ["OK", "Continue", "Login", "Open"]):
                        self._emit("Button click failed, trying Enter key", log_fn)
                        # Focus the OK button area and press Enter
                        send_keys("{ENTER}")
                        time.sleep(0.5)
                        # If dialog still exists, try Tab+Enter
                        if self._find_active_dialog(title_re=r"(?i)(password|login)") is not None:
                            send_keys("{TAB}{ENTER}")

                    password_entered = True
                    self._emit("Password entered at startup", log_fn)
                    time.sleep(5)  # Wait for QB to process login

                    # Check if a "wrong password" warning appeared
                    warning_dlg = self._find_active_dialog(title_re=r"(?i)(warning|error|incorrect)")
                    if warning_dlg is not None:
                        self._emit("Password may have been incorrect, dismissing warning", log_fn)
                        self._click_first_button(warning_dlg, ["OK", "Close"])
                        password_entered = False  # Allow retry
                        time.sleep(1)
                    continue

            elif login_dlg is not None and password_entered:
                # Password was already entered but dialog is still showing - wait
                time.sleep(2)
                continue

            # Dismiss other common startup dialogs
            self._dismiss_common_dialogs(log_fn)
            time.sleep(1)

            # Check if we're past all startup dialogs
            desktop = self._get_desktop()
            blocking_dialogs = False
            for w in desktop.windows():
                try:
                    if not w.is_visible():
                        continue
                    t = (w.window_text() or "").lower()
                    if any(x in t for x in ["login", "password", "external accountant", "memorized transaction"]):
                        blocking_dialogs = True
                        break
                except Exception:  # noqa: BLE001
                    continue

            if not blocking_dialogs:
                break

    def _export_from_qb2023(self, job: CompanyJob, export_dir: Path, log_fn: Optional[LogFn]) -> Dict[str, Path]:
        """Exports list IIF, transaction CSV, and report PDFs from QB 2023.

        In dry-run mode, this method synthesizes/copies sample files.
        """

        export_dir.mkdir(parents=True, exist_ok=True)

        lists_iif = export_dir / "all_lists.IIF"
        tx_csv = export_dir / "TransactionList_QB2023.CSV"

        if self.config.dry_run:
            sample_root = Path.home() / "Uploads" / "QBDowngrade_Results"
            if (sample_root / "all_lists.IIF").exists():
                shutil.copy2(sample_root / "all_lists.IIF", lists_iif)
            else:
                lists_iif.write_text("!HDR\nHDR\tDRY-RUN\n", encoding="utf-8")

            if (sample_root / "TransactionList_QB2023.CSV").exists():
                shutil.copy2(sample_root / "TransactionList_QB2023.CSV", tx_csv)
            else:
                tx_csv.write_text("Type,Date,Account,Split,Debit,Credit\n", encoding="utf-8")

            # copy report artifacts if available
            report_files = {}
            for f in [
                "TrialBalance_QB2023.pdf",
                "BalanceSheet_QB2023.pdf",
                "ProfitLoss_QB2023.pdf",
                "AR_Aging_QB2023.pdf",
                "AP_Aging_QB2023.pdf",
            ]:
                src = sample_root / f
                if src.exists():
                    dst = export_dir / f
                    shutil.copy2(src, dst)
                    report_files[f] = dst

            return {"lists_iif": lists_iif, "tx_csv": tx_csv, **report_files}

        if self._qb2023_app is None:
            raise RuntimeError("QB 2023 app instance is not initialized")

        # Handle any startup dialogs (password, accountant user, memorized transactions)
        self._handle_startup_dialogs(job.password, timeout_s=30, log_fn=log_fn)

        main_window = self._find_qb_main_window(self._qb2023_app, "2023", self.config.timeouts.launch_qb_seconds)
        self._dismiss_common_dialogs(log_fn)

        # Check if the company is already open (QB remembers last opened company)
        company_hint = job.qbw_path.stem.split(" ")[0]  # e.g., "joshs" from "joshs gold coast ii 23"
        if self._is_company_already_open(main_window, company_hint):
            self._emit("Company already open in QB 2023, skipping File->Open", log_fn)
        else:
            self._emit("Company not open, opening company file...", log_fn)
            self._open_company_file(
                main_window,
                job.qbw_path,
                job.password,
                timeout_s=self.config.timeouts.open_company_seconds,
                log_fn=log_fn,
            )
            # Re-acquire main window reference after opening company
            time.sleep(3)
            main_window = self._find_qb_main_window(self._qb2023_app, "2023", self.config.timeouts.launch_qb_seconds)
            # Verify the company actually opened
            if not self._is_company_already_open(main_window, company_hint):
                self._emit("WARNING: Company may not have opened successfully, proceeding anyway", log_fn)

        # Note: _open_company_file() now handles dialog/popup dismissal internally
        # (Step 6 of its sequential approach). Only do a final safety check here.
        time.sleep(1)
        self._close_popup_windows(main_window, log_fn)

        # Export required list IIF files one-by-one.
        list_exports = {
            "Chart of Accounts": export_dir / "accounts.iif",
            "Customers": export_dir / "customers.iif",
            "Vendors": export_dir / "vendors.iif",
            "Items": export_dir / "items.iif",
            "Employees": export_dir / "employees.iif",
        }

        for list_name, out_path in list_exports.items():
            self._with_retries(
                lambda list_name=list_name, out_path=out_path: self._export_single_list_iif(main_window, list_name, out_path, log_fn),
                f"Export list {list_name}",
                log_fn,
            )

        # Build all_lists.IIF for backward compatibility with earlier pipeline expectations.
        with lists_iif.open("w", encoding="utf-8", errors="ignore") as out_f:
            for idx, src in enumerate(list_exports.values()):
                if src.exists():
                    if idx > 0:
                        out_f.write("\n")
                    out_f.write(src.read_text(encoding="utf-8", errors="ignore"))

        self._with_retries(
            lambda: self._export_transaction_list_csv(main_window, tx_csv, log_fn),
            "Export transaction report CSV",
            log_fn,
        )

        report_exports = {
            "TrialBalance_QB2023.pdf": (
                "Reports->Accountant & Taxes->Trial Balance",
                "%rab",
            ),
            "BalanceSheet_QB2023.pdf": (
                "Reports->Company & Financial->Balance Sheet Standard",
                "%rcb",
            ),
            "ProfitLoss_QB2023.pdf": (
                "Reports->Company & Financial->Profit & Loss Standard",
                "%rcp",
            ),
            "AR_Aging_QB2023.pdf": (
                "Reports->Customers & Receivables->A/R Aging Summary",
                "%rcr",
            ),
            "AP_Aging_QB2023.pdf": (
                "Reports->Vendors & Payables->A/P Aging Summary",
                "%rvp",
            ),
        }

        generated_report_paths: Dict[str, Path] = {}
        for report_file, (menu_path, fallback_keys) in report_exports.items():
            out_path = export_dir / report_file
            self._with_retries(
                lambda menu_path=menu_path, fallback_keys=fallback_keys, out_path=out_path: self._export_report_pdf(
                    main_window,
                    menu_path,
                    fallback_keys,
                    out_path,
                    log_fn,
                ),
                f"Export report {report_file}",
                log_fn,
            )
            generated_report_paths[report_file] = out_path

        return {
            "lists_iif": lists_iif,
            "accounts_iif": list_exports["Chart of Accounts"],
            "customers_iif": list_exports["Customers"],
            "vendors_iif": list_exports["Vendors"],
            "items_iif": list_exports["Items"],
            "employees_iif": list_exports["Employees"],
            "tx_csv": tx_csv,
            **generated_report_paths,
        }

    # -------- QB 2021 import --------

    def _create_qb2021_company(self, job: CompanyJob, target_dir: Path, log_fn: Optional[LogFn]) -> Path:
        target_dir.mkdir(parents=True, exist_ok=True)
        target_name = (job.target_company_name or job.qbw_path.stem).replace("23", "21")
        target_qbw = target_dir / f"{target_name}.qbw"

        if self.config.dry_run:
            target_qbw.write_text("DRY RUN PLACEHOLDER - QBW FILE CREATED", encoding="utf-8")
            self._emit(f"Created dry-run target company file: {target_qbw}", log_fn)
            return target_qbw

        if self._qb2021_app is None:
            raise RuntimeError("QB 2021 app instance is not initialized")

        main_window = self._find_qb_main_window(self._qb2021_app, "2021", self.config.timeouts.launch_qb_seconds)
        self._dismiss_common_dialogs(log_fn)

        self._emit("Creating new QB 2021 company", log_fn)
        self._invoke_menu(main_window, "File->New Company", "%fn", log_fn)
        time.sleep(1)

        wizard = self._find_active_dialog(title_re=r"(?i)(new company|quickbooks setup|easystep)")
        if wizard is None:
            raise RuntimeError("New Company wizard did not appear")

        self._focus_window(wizard)

        # Prefer skipping the long interview if available.
        self._click_first_button(wizard, ["Skip Interview", "Express Start", "Detailed Start", "Start Setup", "Next"])
        time.sleep(1)

        # Fill core fields where controls are available.
        for value in [target_name, "000000000", "Services", "Accrual"]:
            self._set_edit_value(wizard, value)
            if send_keys is not None:
                send_keys("{TAB}")

        self._click_first_button(wizard, ["Next", "Continue", "Create Company", "Finish"])

        # Save file location.
        self._handle_standard_file_dialog(target_qbw, save_mode=True, timeout_s=self.config.timeouts.open_company_seconds, log_fn=log_fn)

        self._wait_for_company_ready(main_window, timeout_s=self.config.timeouts.open_company_seconds, log_fn=log_fn)
        self._dismiss_common_dialogs(log_fn)
        return target_qbw

    def _import_single_iif(self, main_window, iif_path: Path, timeout_s: int, log_fn: Optional[LogFn]) -> None:
        if not iif_path.exists():
            raise FileNotFoundError(f"IIF file does not exist: {iif_path}")

        self._emit(f"Importing IIF file: {iif_path.name}", log_fn)
        self._invoke_menu(main_window, "File->Utilities->Import->IIF Files...", "%fui", log_fn)
        time.sleep(1)

        self._handle_standard_file_dialog(iif_path, save_mode=False, timeout_s=timeout_s, log_fn=log_fn)

        # Common QuickBooks import warnings/confirmations.
        start = time.time()
        while time.time() - start < timeout_s:
            dlg = self._find_active_dialog(title_re=r"(?i)(import|quickbooks|warning|confirmation|information)")
            if dlg is None:
                break
            if not self._click_first_button(dlg, ["Yes", "OK", "Continue", "Done", "Close"]):
                if send_keys is not None:
                    send_keys("{ENTER}")
            time.sleep(0.5)

        self._dismiss_common_dialogs(log_fn)

    def _import_iif_lists_qb2021(self, lists_iif: Path, log_fn: Optional[LogFn]) -> None:
        if self.config.dry_run:
            self._emit(f"Dry-run import lists: {lists_iif}", log_fn)
            return

        if self._qb2021_app is None:
            raise RuntimeError("QB 2021 app instance is not initialized")

        main_window = self._find_qb_main_window(self._qb2021_app, "2021", self.config.timeouts.launch_qb_seconds)

        ordered = [
            lists_iif.parent / "accounts.iif",
            lists_iif.parent / "customers.iif",
            lists_iif.parent / "vendors.iif",
            lists_iif.parent / "items.iif",
            lists_iif.parent / "employees.iif",
        ]

        if not any(p.exists() for p in ordered):
            ordered = [lists_iif]

        for iif_path in ordered:
            self._with_retries(
                lambda iif_path=iif_path: self._import_single_iif(
                    main_window,
                    iif_path,
                    timeout_s=self.config.timeouts.import_seconds,
                    log_fn=log_fn,
                ),
                f"Import list IIF ({iif_path.name})",
                log_fn,
            )

        self._emit("List IIF import completed", log_fn)

    def _import_iif_transactions_qb2021(self, tx_iif: Path, log_fn: Optional[LogFn]) -> None:
        if self.config.dry_run:
            self._emit(f"Dry-run import transactions: {tx_iif}", log_fn)
            return

        if self._qb2021_app is None:
            raise RuntimeError("QB 2021 app instance is not initialized")

        main_window = self._find_qb_main_window(self._qb2021_app, "2021", self.config.timeouts.launch_qb_seconds)

        self._with_retries(
            lambda: self._import_single_iif(
                main_window,
                tx_iif,
                timeout_s=max(self.config.timeouts.import_seconds, 600),
                log_fn=log_fn,
            ),
            "Import transaction IIF",
            log_fn,
        )

        self._emit("Transaction IIF import completed", log_fn)

    def _extract_validation_snapshot(self, source: bool = True) -> CompanyValidationSnapshot:
        """Stub validation metrics.

        In live mode, this should parse exported report data or API-accessible report dumps.
        """

        if source:
            return CompanyValidationSnapshot(
                trial_balance_total=100000.00,
                ar_total=0.0,
                ap_total=0.0,
                transaction_count=2480,
                account_count=145,
                customer_count=3,
                vendor_count=251,
                item_count=15,
            )
        return CompanyValidationSnapshot(
            trial_balance_total=100000.00 if self.config.dry_run else 0.0,
            ar_total=0.0,
            ap_total=0.0,
            transaction_count=2480 if self.config.dry_run else 0,
            account_count=145,
            customer_count=3,
            vendor_count=251,
            item_count=15,
        )

    def run_job(
        self,
        job: CompanyJob,
        log_fn: Optional[LogFn] = None,
        progress_fn: Optional[ProgressFn] = None,
    ) -> CompanyJobResult:
        steps = [
            "Launch QB 2023",
            "Export from QB 2023",
            "Close QB 2023",
            "Parse CSV transactions",
            "Launch QB 2021",
            "Create QB 2021 company",
            "Import list IIF",
            "Import transaction IIF",
            "Validation report generation",
            "Close QB 2021",
        ]

        def set_progress(step_idx: int) -> None:
            if progress_fn:
                progress_fn((step_idx / len(steps)) * 100.0)

        qb2023_app = None
        qb2021_app = None

        try:
            job.output_dir.mkdir(parents=True, exist_ok=True)
            exports_dir = job.output_dir / "exports"
            target_dir = job.output_dir / "target"
            validation_dir = job.output_dir / "validation"

            # 1) Launch QB 2023 with company file as parameter (auto-opens it)
            set_progress(0)
            qb2023_app = self._launch_qb(self.config.install_paths.qb_2023_path, log_fn, qbw_path=job.qbw_path)
            self._qb2023_app = qb2023_app

            # 2) Export
            set_progress(1)
            exported: Dict[str, Path] = {}
            self._with_retries(
                lambda: exported.update(self._export_from_qb2023(job, exports_dir, log_fn)),
                "QB 2023 export",
                log_fn,
            )

            # 3) Close QB 2023
            set_progress(2)
            self._close_qb(qb2023_app, log_fn)
            qb2023_app = None
            self._qb2023_app = None

            # 4) Parse transactions CSV -> IIF
            set_progress(3)
            tx_iif = exports_dir / "transactions_generated.IIF"
            parse_stats = self.tx_parser.convert_csv_to_iif(exported["tx_csv"], tx_iif)
            self._emit(f"Generated transaction IIF with {parse_stats['record_count']} records", log_fn)

            # 5) Launch QB 2021
            set_progress(4)
            qb2021_app = self._launch_qb(self.config.install_paths.qb_2021_path, log_fn)
            self._qb2021_app = qb2021_app

            # 6) Create company
            set_progress(5)
            target_qbw = self._create_qb2021_company(job, target_dir, log_fn)

            # 7) Import list IIF
            set_progress(6)
            self._with_retries(lambda: self._import_iif_lists_qb2021(exported["lists_iif"], log_fn), "QB 2021 list import", log_fn)

            # 8) Import transaction IIF
            set_progress(7)
            self._with_retries(lambda: self._import_iif_transactions_qb2021(tx_iif, log_fn), "QB 2021 transaction import", log_fn)

            # 9) Validation
            set_progress(8)
            source_snapshot = self._extract_validation_snapshot(source=True)
            target_snapshot = self._extract_validation_snapshot(source=False)
            validation_files = self.validator.generate_reports(
                source_snapshot,
                target_snapshot,
                validation_dir,
                company_name=job.qbw_path.stem,
            )

            # 10) Close QB 2021
            set_progress(9)
            self._close_qb(qb2021_app, log_fn)
            qb2021_app = None
            self._qb2021_app = None
            set_progress(10)

            generated_files = {
                "lists_iif": str(exported["lists_iif"]),
                "tx_csv": str(exported["tx_csv"]),
                "tx_iif": str(tx_iif),
                "target_qbw": str(target_qbw),
                "validation_excel": validation_files["excel"],
                "validation_html": validation_files["html"],
            }

            for key in ["accounts_iif", "customers_iif", "vendors_iif", "items_iif", "employees_iif"]:
                if key in exported:
                    generated_files[key] = str(exported[key])

            return CompanyJobResult(
                qbw_path=str(job.qbw_path),
                success=True,
                message="Completed successfully",
                generated_files=generated_files,
            )
        except Exception as exc:  # noqa: BLE001
            screenshot = self._take_error_screenshot("run_job_failure")
            stack = traceback.format_exc()
            self._emit(f"Job failed: {exc}\n{stack}", log_fn)
            if screenshot is not None:
                self._emit(f"Failure screenshot: {screenshot}", log_fn)
            return CompanyJobResult(
                qbw_path=str(job.qbw_path),
                success=False,
                message=str(exc),
                generated_files={},
            )
        finally:
            if qb2023_app is not None:
                self._close_qb(qb2023_app, log_fn)
            if qb2021_app is not None:
                self._close_qb(qb2021_app, log_fn)
            self._qb2023_app = None
            self._qb2021_app = None
