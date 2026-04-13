from __future__ import annotations

import importlib
import locale
import os
import queue
import smtplib
import subprocess
import sys
import threading
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Any
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None

try:
    DateEntry = importlib.import_module("tkcalendar").DateEntry
except Exception:  # pragma: no cover
    DateEntry = None


DEFAULT_VENUE_CODE_MAP = {
    "1": "风雨体育馆",
    "2": "松园体育馆",
    "3": "竹园体育馆",
    "4": "星湖体育馆",
    "5": "卓尔体育馆",
    "6": "杏林体育馆",
}


class MonitorGuiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ZHLJ Monitor GUI")
        self.root.geometry("1080x760")

        self.proc: subprocess.Popen[str] | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()

        self.workspace_dir = Path(__file__).resolve().parent
        self.monitor_script = self.workspace_dir / "monitor_pc.py"
        self.default_config = self.workspace_dir / "config.yaml"

        self._build_form()
        self._build_log_panel()
        self._set_running(False)
        self._refresh_startup_status_label()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(120, self._drain_log_queue)

    def _build_form(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.X)

        today_text = date.today().strftime("%Y-%m-%d")
        self.sport_var = tk.StringVar(value="badminton")
        self.date_var = tk.StringVar(value=today_text)
        self.config_var = tk.StringVar(value=str(self.default_config))
        self.log_level_var = tk.StringVar(value="INFO")
        self.run_mode_var = tk.StringVar(value="loop")
        self.email_alert_var = tk.BooleanVar(value=False)

        self.venue_display_to_value = self._load_venue_display_map(Path(self.config_var.get()))
        self.venue_option_items: list[tuple[str, str]] = list(self.venue_display_to_value.items())
        self.time_slot_values = [f"{h:02d}:00-{h + 1:02d}:00" for h in range(8, 21)]

        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=4)
        ttk.Label(row1, text="Sport").pack(side=tk.LEFT)
        ttk.Combobox(
            row1,
            textvariable=self.sport_var,
            values=["badminton", "pingpong"],
            width=14,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(8, 16))

        ttk.Label(row1, text="Date").pack(side=tk.LEFT)
        if DateEntry is not None:
            self.date_picker = DateEntry(row1, textvariable=self.date_var, date_pattern="yyyy-mm-dd", width=12)
            self.date_picker.pack(side=tk.LEFT, padx=(8, 4))
        else:
            ttk.Entry(row1, textvariable=self.date_var, width=14).pack(side=tk.LEFT, padx=(8, 4))
        ttk.Button(row1, text="Today", command=self.set_today).pack(side=tk.LEFT)

        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=4)

        venue_panel = ttk.Frame(row2)
        venue_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        ttk.Label(venue_panel, text="Venues (multi-select)").pack(anchor="w")
        self.venue_listbox = tk.Listbox(venue_panel, selectmode=tk.MULTIPLE, exportselection=False, height=6)
        self.venue_listbox.pack(fill=tk.BOTH, expand=True)
        for label, _ in self.venue_option_items:
            self.venue_listbox.insert(tk.END, label)
        self._select_all_venues()
        venue_btns = ttk.Frame(venue_panel)
        venue_btns.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(venue_btns, text="All", command=self._select_all_venues).pack(side=tk.LEFT)
        ttk.Button(venue_btns, text="Clear", command=self._clear_venues).pack(side=tk.LEFT, padx=(6, 0))

        time_panel = ttk.Frame(row2)
        time_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))
        ttk.Label(time_panel, text="Time Slots (multi-select, default 08:00-21:00)").pack(anchor="w")
        self.time_listbox = tk.Listbox(time_panel, selectmode=tk.MULTIPLE, exportselection=False, height=6)
        self.time_listbox.pack(fill=tk.BOTH, expand=True)
        for slot in self.time_slot_values:
            self.time_listbox.insert(tk.END, slot)
        self._select_all_time_slots()
        time_btns = ttk.Frame(time_panel)
        time_btns.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(time_btns, text="All", command=self._select_all_time_slots).pack(side=tk.LEFT)
        ttk.Button(time_btns, text="Clear", command=self._clear_time_slots).pack(side=tk.LEFT, padx=(6, 0))

        row3 = ttk.Frame(frame)
        row3.pack(fill=tk.X, pady=4)
        ttk.Label(row3, text="Config").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.config_var, width=60).pack(side=tk.LEFT, padx=(8, 8))
        ttk.Button(row3, text="Reload Venues", command=self.reload_venue_options).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(row3, text="Log Level").pack(side=tk.LEFT)
        ttk.Combobox(
            row3,
            textvariable=self.log_level_var,
            values=["DEBUG", "INFO", "WARNING", "ERROR"],
            width=10,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(8, 0))

        row4 = ttk.Frame(frame)
        row4.pack(fill=tk.X, pady=4)
        ttk.Label(row4, text="Run Mode").pack(side=tk.LEFT)
        ttk.Radiobutton(row4, text="Continuous", variable=self.run_mode_var, value="loop").pack(side=tk.LEFT, padx=(8, 0))
        ttk.Radiobutton(row4, text="Once", variable=self.run_mode_var, value="once").pack(side=tk.LEFT, padx=(8, 0))
        ttk.Checkbutton(row4, text="Email Alert", variable=self.email_alert_var).pack(side=tk.LEFT, padx=(16, 0))

        self.start_btn = ttk.Button(row4, text="Start", command=self.start_monitor)
        self.start_btn.pack(side=tk.LEFT, padx=(16, 6))
        self.stop_btn = ttk.Button(row4, text="Stop", command=self.stop_monitor)
        self.stop_btn.pack(side=tk.LEFT)
        ttk.Button(row4, text="Test Email Config", command=self.test_email_config).pack(side=tk.LEFT, padx=(12, 0))

        ttk.Button(row4, text="Open Snapshot HTML", command=self.open_snapshot_html).pack(side=tk.LEFT, padx=(16, 6))
        ttk.Button(row4, text="Open Snapshot PNG", command=self.open_snapshot_png).pack(side=tk.LEFT)

        row5 = ttk.Frame(frame)
        row5.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(row5, text="Enable Startup", command=self.enable_startup).pack(side=tk.LEFT)
        ttk.Button(row5, text="Disable Startup", command=self.disable_startup).pack(side=tk.LEFT, padx=(6, 12))
        self.startup_status_var = tk.StringVar(value="Startup: unknown")
        ttk.Label(row5, textvariable=self.startup_status_var).pack(side=tk.LEFT)

    def _build_log_panel(self) -> None:
        log_frame = ttk.Frame(self.root, padding=(12, 6, 12, 12))
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 10))
        y_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=y_scroll.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._append_log("GUI ready. Click Start to run monitor_pc.py.\n")
        if DateEntry is None:
            self._append_log("Tip: install tkcalendar for graphical date picker: pip install tkcalendar\n")

    def _select_all_venues(self) -> None:
        if self.venue_listbox.size() > 0:
            self.venue_listbox.select_set(0, tk.END)

    def _clear_venues(self) -> None:
        self.venue_listbox.selection_clear(0, tk.END)

    def _select_all_time_slots(self) -> None:
        if self.time_listbox.size() > 0:
            self.time_listbox.select_set(0, tk.END)

    def _clear_time_slots(self) -> None:
        self.time_listbox.selection_clear(0, tk.END)

    def _selected_venue_values(self) -> list[str]:
        values: list[str] = []
        for idx in self.venue_listbox.curselection():
            if idx < 0 or idx >= len(self.venue_option_items):
                continue
            value = self.venue_option_items[idx][1]
            if value:
                values.append(value)
        return values

    def _selected_time_ranges(self) -> list[str]:
        values: list[str] = []
        for idx in self.time_listbox.curselection():
            if idx < 0 or idx >= len(self.time_slot_values):
                continue
            values.append(self.time_slot_values[idx])
        return values

    def _load_venue_display_map(self, config_path: Path) -> dict[str, str]:
        code_map: dict[str, str] = {}
        if yaml is not None and config_path.exists():
            try:
                data = yaml.safe_load(config_path.read_text(encoding="utf-8"))
                monitor = data.get("monitor", {}) if isinstance(data, dict) else {}
                configured = monitor.get("venue_code_map", {}) if isinstance(monitor, dict) else {}
                if isinstance(configured, dict):
                    for code, name in configured.items():
                        code_text = str(code).strip()
                        name_text = str(name).strip()
                        if code_text and name_text:
                            code_map[code_text] = name_text
            except Exception:
                code_map = {}

        if not code_map:
            code_map = dict(DEFAULT_VENUE_CODE_MAP)

        result: dict[str, str] = {}
        sorted_items = sorted(
            code_map.items(),
            key=lambda x: (not str(x[0]).isdigit(), int(x[0]) if str(x[0]).isdigit() else 9999, str(x[0])),
        )
        for code, name in sorted_items:
            result[f"{code} - {name}"] = code
        return result

    def reload_venue_options(self) -> None:
        cfg_text = self.config_var.get().strip() or str(self.default_config)
        self.venue_display_to_value = self._load_venue_display_map(Path(cfg_text))
        self.venue_option_items = list(self.venue_display_to_value.items())

        self.venue_listbox.delete(0, tk.END)
        for label, _ in self.venue_option_items:
            self.venue_listbox.insert(tk.END, label)
        self._select_all_venues()

        self._append_log("Venue options reloaded from config.\n")

    def set_today(self) -> None:
        self.date_var.set(date.today().strftime("%Y-%m-%d"))

    def _load_config_dict(self) -> dict[str, Any]:
        if yaml is None:
            raise ValueError("PyYAML is not installed. Please run: pip install pyyaml")
        cfg_text = self.config_var.get().strip() or str(self.default_config)
        config_path = Path(cfg_text)
        if not config_path.exists():
            raise ValueError(f"Config file not found: {config_path}")
        data = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("Config format invalid: root must be an object.")
        return data

    def test_email_config(self) -> None:
        try:
            data = self._load_config_dict()
            notification = data.get("notification", {}) if isinstance(data, dict) else {}
            email_cfg = notification.get("email", {}) if isinstance(notification, dict) else {}
            if not isinstance(email_cfg, dict):
                raise ValueError("notification.email config is missing or invalid.")

            smtp_host = str(email_cfg.get("smtp_host", "")).strip()
            smtp_port = int(email_cfg.get("smtp_port", 465) or 465)
            use_ssl = bool(email_cfg.get("use_ssl", True))
            use_starttls = bool(email_cfg.get("use_starttls", True))
            username = str(email_cfg.get("username", "")).strip()
            password = str(email_cfg.get("password", "")).strip()
            from_addr = str(email_cfg.get("from_addr", username)).strip()

            to_addrs_raw = email_cfg.get("to_addrs", [])
            to_addrs: list[str] = []
            if isinstance(to_addrs_raw, list):
                to_addrs = [str(x).strip() for x in to_addrs_raw if str(x).strip()]
            elif isinstance(to_addrs_raw, str):
                to_addrs = [x.strip() for x in to_addrs_raw.replace(";", ",").split(",") if x.strip()]

            if not smtp_host or not from_addr or not to_addrs:
                raise ValueError("Please set smtp_host/from_addr/to_addrs in notification.email.")

            subject_prefix = str(email_cfg.get("subject_prefix", "[场馆余量提醒]")).strip() or "[场馆余量提醒]"
            msg = EmailMessage()
            msg["Subject"] = f"{subject_prefix} GUI测试 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            msg["From"] = from_addr
            msg["To"] = ", ".join(to_addrs)
            msg.set_content(
                "This is a test email from monitor_gui.py.\n"
                f"time={datetime.now().isoformat(timespec='seconds')}\n"
                f"smtp_host={smtp_host}:{smtp_port}\n",
                charset="utf-8",
            )

            self._append_log("Testing email config...\n")
            if use_ssl:
                with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=15) as server:
                    if username and password:
                        server.login(username, password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
                    if use_starttls:
                        server.starttls()
                    if username and password:
                        server.login(username, password)
                    server.send_message(msg)

            self._append_log(f"Test email sent successfully to: {', '.join(to_addrs)}\n")
            messagebox.showinfo("Email Test", "Test email sent successfully.")
        except Exception as exc:
            self._append_log(f"Test email failed: {exc}\n")
            messagebox.showerror("Email Test Failed", str(exc))

    def _append_log(self, text: str) -> None:
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)

    def _drain_log_queue(self) -> None:
        while True:
            try:
                item = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log(item)
        self.root.after(120, self._drain_log_queue)

    def _build_command(self) -> list[str]:
        cmd = [
            sys.executable,
            str(self.monitor_script),
            "--config",
            self.config_var.get().strip() or str(self.default_config),
        ]

        sport = self.sport_var.get().strip() or "badminton"
        cmd.extend(["--sport", sport])

        selected_venues = self._selected_venue_values()
        if selected_venues:
            cmd.extend(["--venue", ",".join(selected_venues)])

        selected_ranges = self._selected_time_ranges()
        if selected_ranges:
            cmd.extend(["--time-ranges", ",".join(selected_ranges)])

        date_text = self.date_var.get().strip()
        if date_text:
            cmd.extend(["--date", date_text])

        if self.run_mode_var.get() == "once":
            cmd.append("--once")

        if self.email_alert_var.get():
            cmd.append("--email-alert")

        level = self.log_level_var.get().strip() or "INFO"
        cmd.extend(["--log-level", level])

        return cmd

    def _set_running(self, running: bool) -> None:
        state_start = tk.DISABLED if running else tk.NORMAL
        state_stop = tk.NORMAL if running else tk.DISABLED
        self.start_btn.configure(state=state_start)
        self.stop_btn.configure(state=state_stop)

    def start_monitor(self) -> None:
        if self.proc is not None and self.proc.poll() is None:
            messagebox.showinfo("Running", "Monitor process is already running.")
            return

        if not self.monitor_script.exists():
            messagebox.showerror("Missing File", f"File not found: {self.monitor_script}")
            return

        cmd = self._build_command()
        self._append_log("\n=== START ===\n")
        self._append_log("Command: " + " ".join(cmd) + "\n")

        creationflags = 0
        if os.name == "nt":
            creationflags = subprocess.CREATE_NO_WINDOW

        try:
            stdout_encoding = locale.getpreferredencoding(False) or "utf-8"
            self.proc = subprocess.Popen(
                cmd,
                cwd=str(self.workspace_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding=stdout_encoding,
                errors="replace",
                bufsize=1,
                creationflags=creationflags,
            )
        except Exception as exc:
            messagebox.showerror("Start Failed", str(exc))
            return

        self._set_running(True)
        threading.Thread(target=self._read_proc_output, daemon=True).start()
        threading.Thread(target=self._wait_proc_exit, daemon=True).start()

    def _read_proc_output(self) -> None:
        if self.proc is None or self.proc.stdout is None:
            return
        for line in self.proc.stdout:
            self.log_queue.put(line)

    def _wait_proc_exit(self) -> None:
        if self.proc is None:
            return
        code = self.proc.wait()
        self.log_queue.put(f"\n=== EXIT: {code} ===\n")
        self.root.after(0, lambda: self._set_running(False))

    def stop_monitor(self) -> None:
        if self.proc is None or self.proc.poll() is not None:
            return
        self._append_log("\nStopping process...\n")
        try:
            self.proc.terminate()
        except Exception:
            pass

    def _startup_launcher_path(self) -> Path:
        appdata = os.environ.get("APPDATA", "")
        startup_dir = Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
        return startup_dir / "zhlj-monitor-gui.cmd"

    def _refresh_startup_status_label(self) -> None:
        launcher = self._startup_launcher_path()
        self.startup_status_var.set("Startup: enabled" if launcher.exists() else "Startup: disabled")

    def enable_startup(self) -> None:
        if os.name != "nt":
            messagebox.showwarning("Unsupported", "Startup shortcut currently supports Windows only.")
            return

        launcher = self._startup_launcher_path()
        launcher.parent.mkdir(parents=True, exist_ok=True)

        pythonw = Path(sys.executable).with_name("pythonw.exe")
        py_exec = str(pythonw if pythonw.exists() else Path(sys.executable))
        gui_script = self.workspace_dir / "monitor_gui.py"

        content = "\n".join(
            [
                "@echo off",
                f'cd /d "{self.workspace_dir}"',
                f'start "" "{py_exec}" "{gui_script}"',
            ]
        )
        launcher.write_text(content, encoding="utf-8")
        self._refresh_startup_status_label()
        messagebox.showinfo("Startup", f"Enabled at: {launcher}")

    def disable_startup(self) -> None:
        launcher = self._startup_launcher_path()
        if launcher.exists():
            launcher.unlink()
        self._refresh_startup_status_label()
        messagebox.showinfo("Startup", "Startup launcher removed.")

    def open_snapshot_html(self) -> None:
        self._open_file(self.workspace_dir / "page_snapshot.html")

    def open_snapshot_png(self) -> None:
        self._open_file(self.workspace_dir / "page_snapshot.png")

    def _open_file(self, path: Path) -> None:
        if not path.exists():
            messagebox.showwarning("Not Found", f"File not found: {path}")
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror("Open Failed", str(exc))

    def on_close(self) -> None:
        if self.proc is not None and self.proc.poll() is None:
            self.stop_monitor()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    MonitorGuiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
