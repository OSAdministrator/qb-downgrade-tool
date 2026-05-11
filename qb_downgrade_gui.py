"""Main tkinter GUI for QuickBooks TimeWarp\u00ae by Our System Administrator."""

from __future__ import annotations

import json
import logging
import os
import queue
import threading
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, VERTICAL, BooleanVar, N, S, E, W
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from config import AppConfig, ConfigManager
from qb_automation import CompanyJob, QuickBooksAutomationEngine


@dataclass
class QueueItem:
    qbw_path: Path
    password: str


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, config: AppConfig, on_save):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("700x430")
        self.resizable(False, False)
        self.config_data = config
        self.on_save = on_save

        pad = {"padx": 8, "pady": 6}

        self.var_qb2023 = tk.StringVar(value=config.install_paths.qb_2023_path)
        self.var_qb2021 = tk.StringVar(value=config.install_paths.qb_2021_path)
        self.var_output = tk.StringVar(value=config.default_output_dir)
        self.var_retry_attempts = tk.IntVar(value=config.retry_attempts)
        self.var_retry_delay = tk.IntVar(value=config.retry_delay_seconds)
        self.var_continue = BooleanVar(value=config.continue_on_error)
        self.var_dry_run = BooleanVar(value=config.dry_run)

        row = 0
        tk.Label(self, text="QB 2023 executable").grid(row=row, column=0, sticky=E, **pad)
        tk.Entry(self, textvariable=self.var_qb2023, width=62).grid(row=row, column=1, sticky=W, **pad)
        tk.Button(self, text="Browse", command=lambda: self._pick_file(self.var_qb2023)).grid(row=row, column=2, **pad)

        row += 1
        tk.Label(self, text="QB 2021 executable").grid(row=row, column=0, sticky=E, **pad)
        tk.Entry(self, textvariable=self.var_qb2021, width=62).grid(row=row, column=1, sticky=W, **pad)
        tk.Button(self, text="Browse", command=lambda: self._pick_file(self.var_qb2021)).grid(row=row, column=2, **pad)

        row += 1
        tk.Label(self, text="Default output directory").grid(row=row, column=0, sticky=E, **pad)
        tk.Entry(self, textvariable=self.var_output, width=62).grid(row=row, column=1, sticky=W, **pad)
        tk.Button(self, text="Browse", command=lambda: self._pick_dir(self.var_output)).grid(row=row, column=2, **pad)

        row += 1
        tk.Label(self, text="Retry attempts").grid(row=row, column=0, sticky=E, **pad)
        tk.Spinbox(self, from_=0, to=10, textvariable=self.var_retry_attempts, width=10).grid(row=row, column=1, sticky=W, **pad)

        row += 1
        tk.Label(self, text="Retry delay (seconds)").grid(row=row, column=0, sticky=E, **pad)
        tk.Spinbox(self, from_=1, to=120, textvariable=self.var_retry_delay, width=10).grid(row=row, column=1, sticky=W, **pad)

        row += 1
        tk.Checkbutton(self, text="Continue processing if a file fails", variable=self.var_continue).grid(row=row, column=1, sticky=W, **pad)

        row += 1
        tk.Checkbutton(self, text="Dry-run mode (recommended while testing)", variable=self.var_dry_run).grid(row=row, column=1, sticky=W, **pad)

        row += 1
        tk.Label(self, text="Timeouts (seconds)", font=("Segoe UI", 9, "bold")).grid(row=row, column=0, sticky=E, **pad)
        timeout_frame = tk.Frame(self)
        timeout_frame.grid(row=row, column=1, sticky=W)

        self.timeout_vars = {
            "launch_qb_seconds": tk.IntVar(value=config.timeouts.launch_qb_seconds),
            "open_company_seconds": tk.IntVar(value=config.timeouts.open_company_seconds),
            "export_seconds": tk.IntVar(value=config.timeouts.export_seconds),
            "import_seconds": tk.IntVar(value=config.timeouts.import_seconds),
            "report_seconds": tk.IntVar(value=config.timeouts.report_seconds),
            "app_close_seconds": tk.IntVar(value=config.timeouts.app_close_seconds),
        }

        col = 0
        for key, var in self.timeout_vars.items():
            tk.Label(timeout_frame, text=key.replace("_", " ")).grid(row=0, column=col, sticky=W, padx=4)
            tk.Spinbox(timeout_frame, from_=10, to=600, textvariable=var, width=8).grid(row=1, column=col, padx=4)
            col += 1

        row += 1
        footer = tk.Frame(self)
        footer.grid(row=row, column=0, columnspan=3, pady=12)
        tk.Button(footer, text="Save", command=self._save, width=15).pack(side=LEFT, padx=6)
        tk.Button(footer, text="Cancel", command=self.destroy, width=15).pack(side=LEFT, padx=6)

    def _pick_file(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename()
        if path:
            var.set(path)

    def _pick_dir(self, var: tk.StringVar) -> None:
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _save(self) -> None:
        cfg = self.config_data
        cfg.install_paths.qb_2023_path = self.var_qb2023.get().strip()
        cfg.install_paths.qb_2021_path = self.var_qb2021.get().strip()
        cfg.default_output_dir = self.var_output.get().strip()
        cfg.retry_attempts = int(self.var_retry_attempts.get())
        cfg.retry_delay_seconds = int(self.var_retry_delay.get())
        cfg.continue_on_error = bool(self.var_continue.get())
        cfg.dry_run = bool(self.var_dry_run.get())

        cfg.timeouts.launch_qb_seconds = int(self.timeout_vars["launch_qb_seconds"].get())
        cfg.timeouts.open_company_seconds = int(self.timeout_vars["open_company_seconds"].get())
        cfg.timeouts.export_seconds = int(self.timeout_vars["export_seconds"].get())
        cfg.timeouts.import_seconds = int(self.timeout_vars["import_seconds"].get())
        cfg.timeouts.report_seconds = int(self.timeout_vars["report_seconds"].get())
        cfg.timeouts.app_close_seconds = int(self.timeout_vars["app_close_seconds"].get())

        self.on_save(cfg)
        self.destroy()


class QuickBooksDowngradeGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("QuickBooks TimeWarp\u00ae by Our System Administrator")
        self.root.geometry("1100x430")

        self.config_manager = ConfigManager()
        self.config = self.config_manager.load()

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self._setup_logging()
        self._build_ui()
        self.root.after(200, self._poll_log_queue)

    def _setup_logging(self) -> None:
        # Logs go to workspace (flash drive partition or program folder — auto-detected)
        try:
            from drive_layout import get_workspace_paths
            log_dir = get_workspace_paths()["log_dir"]
        except Exception:
            log_dir = Path(__file__).resolve().parent / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "qb_downgrade_tool.log"

        self.logger = logging.getLogger("qb_downgrade")
        self.logger.setLevel(logging.INFO)
        self.logger.handlers.clear()

        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        self.logger.addHandler(fh)

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        ttk.Button(top, text="Add QBW Files", command=self._add_files).pack(side=LEFT, padx=4)
        ttk.Button(top, text="Remove Selected", command=self._remove_selected).pack(side=LEFT, padx=4)
        ttk.Button(top, text="Settings", command=self._open_settings).pack(side=LEFT, padx=4)

        self.btn_start = ttk.Button(top, text="Start Processing", command=self._start_processing)
        self.btn_start.pack(side=RIGHT, padx=4)

        # Drop-zone: visual area that doubles as a click-to-add target.
        # If tkinterdnd2 is available, actual drag-and-drop works here too.
        self.drop_frame = tk.Frame(self.root, bg="#2a3a4a", relief="groove", bd=2)
        self.drop_frame.pack(fill=tk.X, padx=10, pady=(4, 2))
        self.drop_label = tk.Label(
            self.drop_frame,
            text="📂  Drag .QBW files onto Launch TimeWarp  —  or click Add QBW Files above",
            bg="#2a3a4a", fg="#8ab4d8", font=("Segoe UI", 10),
            padx=20, pady=6, cursor="hand2",
        )
        self.drop_label.pack(fill=tk.X)
        self.drop_label.bind("<Button-1>", lambda e: self._add_files())

        # Try to enable native drag-and-drop via tkinterdnd2 (optional dependency)
        self._setup_dnd()

        columns = ("file", "status", "message")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=2)
        self.tree.pack(fill=BOTH, expand=False, padx=10, pady=4)

        self.tree.heading("file", text="QBW File")
        self.tree.heading("status", text="Status")
        self.tree.heading("message", text="Message")

        self.tree.column("file", width=620)
        self.tree.column("status", width=110)
        self.tree.column("message", width=330)

        progress_frame = ttk.Frame(self.root, padding=(10, 4))
        progress_frame.pack(fill=tk.X)
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, side=LEFT, expand=True)

        # Heartbeat / elapsed timer — shows user the process is alive
        self.elapsed_var = tk.StringVar(value="")
        self.elapsed_label = ttk.Label(progress_frame, textvariable=self.elapsed_var,
                                        font=("Segoe UI", 9), width=22, anchor="e")
        self.elapsed_label.pack(side=RIGHT, padx=(8, 0))
        self._heartbeat_start: float | None = None
        self._heartbeat_phase: str = ""
        self._heartbeat_after_id: str | None = None

        # Sub-status line: shows the current step in human-friendly text
        # (e.g. "Importing customers: 234 / 891"). Sits between progress
        # bar and the scrolling log so the operator always knows what
        # phase we're in without reading the log.
        self.status_var = tk.StringVar(value="Idle — drop or add .QBW files to begin")
        status_label = tk.Label(
            self.root, textvariable=self.status_var,
            font=("Segoe UI", 10, "bold"),
            fg="#0f172a", bg="#e2e8f0",
            anchor="w", padx=12, pady=6,
        )
        status_label.pack(fill=tk.X, padx=10, pady=(0, 4))

        log_frame = ttk.LabelFrame(self.root, text="Activity Log", padding=8)
        log_frame.pack(fill=BOTH, expand=True, padx=10, pady=4)

        self.log_text = tk.Text(
            log_frame, height=6, wrap="word",
            bg="#0f172a", fg="#e2e8f0",
            font=("Consolas", 9),
            insertbackground="#e2e8f0",
            selectbackground="#334155",
        )
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)

        # Color tags for the log: green=success, yellow=warn/retry, red=error,
        # cyan=phase header, gray=watchdog / housekeeping.
        self.log_text.tag_configure("ok",       foreground="#4ade80")
        self.log_text.tag_configure("warn",     foreground="#facc15")
        self.log_text.tag_configure("err",      foreground="#f87171")
        self.log_text.tag_configure("phase",    foreground="#22d3ee", font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("watchdog", foreground="#94a3b8")
        self.log_text.tag_configure("info",     foreground="#e2e8f0")

        scroll = ttk.Scrollbar(log_frame, orient=VERTICAL, command=self.log_text.yview)
        scroll.pack(side=RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scroll.set)

        # Branded footer — every window is a billboard
        footer = tk.Frame(self.root, bg="#1e293b")
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(
            footer,
            text="oursystemadmin.com  •  liveremotesupport.com",
            bg="#1e293b", fg="#64748b", font=("Segoe UI", 8),
            padx=10, pady=4, cursor="hand2",
        ).pack(side=LEFT)
        tk.Label(
            footer,
            text="© 2026 Our System Administrator, LLC",
            bg="#1e293b", fg="#64748b", font=("Segoe UI", 8),
            padx=10, pady=4,
        ).pack(side=RIGHT)

    def _setup_dnd(self) -> None:
        """Enable native Windows drag-and-drop INTO the GUI if windnd is available.

        windnd is a lightweight (<50 KB) package that hooks Windows OLE drag-and-drop.
        Install with:  pip install windnd
        If not installed, the GUI still works — just use the bat file or Add button.
        """
        try:
            import windnd  # type: ignore[import-untyped]

            def _on_drop(file_list: list) -> None:
                added = 0
                for raw in file_list:
                    # windnd gives bytes on some versions, str on others
                    fp = raw.decode("utf-8") if isinstance(raw, bytes) else str(raw)
                    fp = fp.strip().strip('"').strip("'")
                    if fp.lower().endswith(".qbw"):
                        self.tree.insert("", END, values=(fp, "Queued", ""))
                        added += 1
                if added:
                    self.drop_label.config(text=f"✅  {added} file(s) added — click Start")
                    self._log(f"Drag-and-drop: added {added} .QBW file(s)")

            windnd.hook_dropfiles(self.root, func=_on_drop)
            self.drop_label.config(
                text="📂  Drop .QBW files here  —  or click Add QBW Files above"
            )
        except ImportError:
            pass  # windnd not installed — bat file + Add button still work

    def _add_files(self) -> None:
        try:
            files = filedialog.askopenfilenames(filetypes=[("QuickBooks Company", "*.qbw")])
            for f in files:
                self.tree.insert("", END, values=(f, "Queued", ""))
        except Exception:
            pass
        # If no files were added (e.g. file-in-use error), offer manual path entry
        if not self.tree.get_children():
            self._add_file_manually()

    def _add_file_manually(self) -> None:
        """Allow typing a file path directly — bypasses file-in-use locks."""
        from tkinter import simpledialog
        path = simpledialog.askstring(
            "Add QBW File",
            "Enter the full path to the .qbw file:\n\n"
            "(Use this when the file is already open in QuickBooks)",
            parent=self.root
        )
        if path and path.strip():
            path = path.strip().strip('"').strip("'")
            if os.path.exists(path) or path.lower().endswith('.qbw'):
                self.tree.insert("", END, values=(path, "Queued", ""))
            else:
                messagebox.showwarning("Invalid Path", f"File not found: {path}")

    def _remove_selected(self) -> None:
        for row in self.tree.selection():
            self.tree.delete(row)

    def _open_settings(self) -> None:
        SettingsDialog(self.root, self.config, self._save_settings)

    def _save_settings(self, new_cfg: AppConfig) -> None:
        self.config = new_cfg
        self.config_manager.save(self.config)
        self._log("Settings saved.")

    def _queue_items(self) -> list[tuple[str, QueueItem]]:
        items = []
        for row in self.tree.get_children():
            file_path, status, _ = self.tree.item(row, "values")
            if not file_path:
                continue
            items.append((row, QueueItem(qbw_path=Path(file_path), password="")))
        return items

    def _start_processing(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Processing", "Batch processing is already running.")
            return

        queue_items = self._queue_items()
        if not queue_items:
            messagebox.showwarning("No files", "Add one or more .QBW files first.")
            return

        self.btn_start.configure(state=tk.DISABLED)
        self.progress["value"] = 0
        self._start_heartbeat("Starting")

        self.worker_thread = threading.Thread(target=self._process_batch, args=(queue_items,), daemon=True)
        self.worker_thread.start()

    def _process_batch(self, queue_items: list[tuple[str, QueueItem]]) -> None:
        engine = QuickBooksAutomationEngine(self.config, logger=self.logger)
        output_root = Path(self.config.default_output_dir)
        output_root.mkdir(parents=True, exist_ok=True)

        total = len(queue_items)
        for idx, (row_id, item) in enumerate(queue_items, start=1):
            self._update_row(row_id, "Running", "")
            self._log(f"[{idx}/{total}] Processing {item.qbw_path.name}")

            company_output = output_root / item.qbw_path.stem
            job = CompanyJob(qbw_path=item.qbw_path, password=item.password, output_dir=company_output)

            result = engine.run_job(
                job,
                log_fn=self._log,
                progress_fn=lambda p, idx=idx, total=total: self._set_batch_progress(idx, total, p),
            )

            if result.success:
                self._update_row(row_id, "Success", "Completed")
                self._log(f"✔ {item.qbw_path.name}: success")
            else:
                self._update_row(row_id, "Failed", result.message)
                self._log(f"✖ {item.qbw_path.name}: {result.message}")
                if not self.config.continue_on_error:
                    self._log("Stopping batch because continue_on_error is disabled.")
                    break

            self.progress["value"] = (idx / total) * 100

        self.btn_start.configure(state=tk.NORMAL)
        self._stop_heartbeat()
        self._log("Batch processing finished.")

    def _set_batch_progress(self, item_index: int, total_items: int, item_percent: float) -> None:
        base = (item_index - 1) / total_items * 100
        slot = (1 / total_items) * 100
        self.progress["value"] = base + (slot * (item_percent / 100.0))

    def _update_row(self, row_id: str, status: str, message: str) -> None:
        vals = list(self.tree.item(row_id, "values"))
        vals[1] = status
        vals[2] = message
        self.tree.item(row_id, values=vals)

    def _start_heartbeat(self, phase: str = "Processing") -> None:
        """Start the elapsed-time heartbeat ticker."""
        import time as _time
        self._heartbeat_start = _time.time()
        self._heartbeat_phase = phase
        self._tick_heartbeat()

    def _tick_heartbeat(self) -> None:
        """Update elapsed time display every second while processing."""
        import time as _time
        if self._heartbeat_start is None:
            return
        elapsed = int(_time.time() - self._heartbeat_start)
        m, s = divmod(elapsed, 60)
        h, m = divmod(m, 60)
        if h:
            elapsed_str = f"{h}:{m:02d}:{s:02d}"
        else:
            elapsed_str = f"{m}:{s:02d}"

        # Spinning indicator so user sees it's alive
        spinner = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        spin_char = spinner[elapsed % len(spinner)]

        self.elapsed_var.set(f"{spin_char} {self._heartbeat_phase} {elapsed_str}")
        self._heartbeat_after_id = self.root.after(1000, self._tick_heartbeat)

    def _stop_heartbeat(self) -> None:
        """Stop the heartbeat ticker."""
        self._heartbeat_start = None
        if self._heartbeat_after_id:
            self.root.after_cancel(self._heartbeat_after_id)
            self._heartbeat_after_id = None
        self.elapsed_var.set("")

    def _update_heartbeat_phase(self, message: str) -> None:
        """Auto-detect the current phase from log messages."""
        lower = message.lower()
        if "exporting lists" in lower:
            self._heartbeat_phase = "Exporting Lists"
        elif "exporting transactions" in lower:
            self._heartbeat_phase = "Exporting Txns"
        elif "rendered pdf" in lower or "report" in lower:
            self._heartbeat_phase = "Generating PDFs"
        elif "launching" in lower:
            self._heartbeat_phase = "Launching QB"
        elif "beginsession" in lower:
            self._heartbeat_phase = "QBFC Connected"
        elif "succeeded" in lower or "success" in lower:
            self._heartbeat_phase = "Complete"

    def _log(self, message: str) -> None:
        self.log_queue.put(message)
        self._update_heartbeat_phase(message)

    def _classify_log(self, msg: str) -> str:
        """Pick a color tag based on log message content."""
        m = msg.lower()
        if msg.startswith("===") or "phase" in m and ":" in msg:
            return "phase"
        if "[watchdog]" in m:
            return "watchdog"
        if "✔" in msg or "succeeded" in m or "complete" in m or "✓" in msg:
            return "ok"
        if "✖" in msg or "failed" in m or "error" in m or "traceback" in m:
            return "err"
        if "warn" in m or "retry" in m or "skipped" in m or "timeout" in m:
            return "warn"
        return "info"

    def _update_status_from_log(self, msg: str) -> None:
        """Update the bold sub-status line based on key log markers."""
        # Phase headers like "=== QBFC Import: Phase 3 — Entities ==="
        if msg.startswith("===") and msg.endswith("==="):
            self.status_var.set(msg.strip("= ").strip())
            return
        # Per-step lines: "  Snapshot: 234 customers" etc.
        for marker in ("Snapshot:", "Importing", "Exporting", "Validating",
                       "Opening", "Closing", "QBFC import complete",
                       "Company:", "Restored"):
            if marker in msg:
                # Trim leading whitespace/bullets so status reads cleanly
                self.status_var.set(msg.strip().lstrip("•-> "))
                return

    def _poll_log_queue(self) -> None:
        while True:
            try:
                msg = self.log_queue.get_nowait()
                tag = self._classify_log(msg)
                self.log_text.insert(END, msg + "\n", tag)
                self.log_text.see(END)
                self._update_status_from_log(msg)
            except queue.Empty:
                break
        self.root.after(200, self._poll_log_queue)


def _auto_update() -> None:
    """Check for updates on GitHub and pull if behind. Runs before GUI launches."""
    import subprocess
    import sys

    script_dir = Path(__file__).resolve().parent

    # Make sure we're in a git repo
    if not (script_dir / ".git").exists():
        return

    try:
        # Fetch latest from remote (silent)
        subprocess.run(
            ["git", "fetch", "--quiet"],
            cwd=str(script_dir), capture_output=True, timeout=15
        )

        # Compare local HEAD vs remote tracking branch
        local = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=str(script_dir), capture_output=True, text=True, timeout=5
        ).stdout.strip()

        remote = subprocess.run(
            ["git", "rev-parse", "@{u}"],
            cwd=str(script_dir), capture_output=True, text=True, timeout=5
        ).stdout.strip()

        if local == remote:
            print("[Auto-Update] Already up to date.")
            return

        print(f"[Auto-Update] Update available: {local[:8]} -> {remote[:8]}")
        print("[Auto-Update] Pulling latest changes...")

        result = subprocess.run(
            ["git", "pull", "--ff-only"],
            cwd=str(script_dir), capture_output=True, text=True, timeout=30
        )
        print(result.stdout)

        if result.returncode == 0:
            print("[Auto-Update] Updated successfully. Restarting...")
            # Re-launch ourselves with the same arguments
            os.execv(sys.executable, [sys.executable] + sys.argv)
            # execv replaces the process — code below never runs
        else:
            print(f"[Auto-Update] Pull failed (non-fatal): {result.stderr}")

    except Exception as exc:  # noqa: BLE001
        print(f"[Auto-Update] Check failed (non-fatal): {exc}")


def main() -> None:
    import sys

    # Auto-update from git before launching GUI
    _auto_update()

    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = QuickBooksDowngradeGUI(root)
    # Accept command-line file paths to pre-populate the queue
    # (This is how drag-and-drop onto the .bat launcher works)
    cli_count = 0
    for arg in sys.argv[1:]:
        if arg.lower().endswith('.qbw'):
            app.tree.insert("", END, values=(arg, "Queued", ""))
            cli_count += 1
    if cli_count:
        app.drop_label.config(text=f"✅  {cli_count} file(s) loaded — click Start")
    root.mainloop()


if __name__ == "__main__":
    main()
