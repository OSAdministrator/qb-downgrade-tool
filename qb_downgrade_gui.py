"""Main tkinter GUI for QuickBooks Downgrade Tool."""

from __future__ import annotations

import json
import logging
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
        self.root.title("QuickBooks Downgrade Tool")
        self.root.geometry("1100x650")

        self.config_manager = ConfigManager()
        self.config = self.config_manager.load()

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self._setup_logging()
        self._build_ui()
        self.root.after(200, self._poll_log_queue)

    def _setup_logging(self) -> None:
        Path("logs").mkdir(exist_ok=True)
        log_path = Path("logs") / "qb_downgrade_tool.log"

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
        ttk.Button(top, text="Load Passwords JSON", command=self._load_password_map).pack(side=LEFT, padx=4)
        ttk.Button(top, text="Settings", command=self._open_settings).pack(side=LEFT, padx=4)

        self.btn_start = ttk.Button(top, text="Start Processing", command=self._start_processing)
        self.btn_start.pack(side=RIGHT, padx=4)

        columns = ("file", "password", "status", "message")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=8)
        self.tree.pack(fill=BOTH, expand=False, padx=10, pady=8)

        self.tree.heading("file", text="QBW File")
        self.tree.heading("password", text="Admin Password")
        self.tree.heading("status", text="Status")
        self.tree.heading("message", text="Message")

        self.tree.column("file", width=470)
        self.tree.column("password", width=160)
        self.tree.column("status", width=110)
        self.tree.column("message", width=330)

        pw_frame = ttk.Frame(self.root)
        pw_frame.pack(fill=tk.X, padx=10)
        ttk.Label(pw_frame, text="Default password for selected rows:").pack(side=LEFT, padx=4)
        self.default_password_var = tk.StringVar()
        ttk.Entry(pw_frame, textvariable=self.default_password_var, width=30).pack(side=LEFT, padx=4)
        ttk.Button(pw_frame, text="Apply", command=self._apply_default_password).pack(side=LEFT, padx=4)

        progress_frame = ttk.Frame(self.root, padding=(10, 4))
        progress_frame.pack(fill=tk.X)
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X)

        log_frame = ttk.LabelFrame(self.root, text="Status Log", padding=8)
        log_frame.pack(fill=BOTH, expand=True, padx=10, pady=8)

        self.log_text = tk.Text(log_frame, height=10, wrap="word")
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)

        scroll = ttk.Scrollbar(log_frame, orient=VERTICAL, command=self.log_text.yview)
        scroll.pack(side=RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scroll.set)

    def _add_files(self) -> None:
        files = filedialog.askopenfilenames(filetypes=[("QuickBooks Company", "*.qbw")])
        for f in files:
            self.tree.insert("", END, values=(f, "", "Queued", ""))

    def _remove_selected(self) -> None:
        for row in self.tree.selection():
            self.tree.delete(row)

    def _apply_default_password(self) -> None:
        default_pw = self.default_password_var.get().strip()
        if not default_pw:
            messagebox.showwarning("Password", "Enter a default password first.")
            return

        targets = self.tree.selection() or self.tree.get_children()
        for row in targets:
            vals = list(self.tree.item(row, "values"))
            vals[1] = default_pw
            self.tree.item(row, values=vals)

    def _load_password_map(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return

        with open(path, "r", encoding="utf-8") as f:
            mapping = json.load(f)

        updated = 0
        for row in self.tree.get_children():
            vals = list(self.tree.item(row, "values"))
            file_name = Path(vals[0]).name
            if file_name in mapping:
                vals[1] = mapping[file_name]
                self.tree.item(row, values=vals)
                updated += 1

        self._log(f"Loaded password map: updated {updated} queue entries")

    def _open_settings(self) -> None:
        SettingsDialog(self.root, self.config, self._save_settings)

    def _save_settings(self, new_cfg: AppConfig) -> None:
        self.config = new_cfg
        self.config_manager.save(self.config)
        self._log("Settings saved.")

    def _queue_items(self) -> list[tuple[str, QueueItem]]:
        items = []
        for row in self.tree.get_children():
            file_path, password, status, _ = self.tree.item(row, "values")
            if not file_path:
                continue
            items.append((row, QueueItem(qbw_path=Path(file_path), password=password or "")))
        return items

    def _start_processing(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Processing", "Batch processing is already running.")
            return

        queue_items = self._queue_items()
        if not queue_items:
            messagebox.showwarning("No files", "Add one or more .QBW files first.")
            return

        missing_pw = [item.qbw_path.name for _, item in queue_items if not item.password]
        if missing_pw:
            proceed = messagebox.askyesno(
                "Missing passwords",
                f"{len(missing_pw)} files have blank passwords. Continue anyway?",
            )
            if not proceed:
                return

        self.btn_start.configure(state=tk.DISABLED)
        self.progress["value"] = 0

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
        self._log("Batch processing finished.")

    def _set_batch_progress(self, item_index: int, total_items: int, item_percent: float) -> None:
        base = (item_index - 1) / total_items * 100
        slot = (1 / total_items) * 100
        self.progress["value"] = base + (slot * (item_percent / 100.0))

    def _update_row(self, row_id: str, status: str, message: str) -> None:
        vals = list(self.tree.item(row_id, "values"))
        vals[2] = status
        vals[3] = message
        self.tree.item(row_id, values=vals)

    def _log(self, message: str) -> None:
        self.log_queue.put(message)

    def _poll_log_queue(self) -> None:
        while True:
            try:
                msg = self.log_queue.get_nowait()
                self.log_text.insert(END, msg + "\n")
                self.log_text.see(END)
            except queue.Empty:
                break
        self.root.after(200, self._poll_log_queue)


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = QuickBooksDowngradeGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
