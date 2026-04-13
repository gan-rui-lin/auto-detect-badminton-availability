from __future__ import annotations

import os
import locale
import queue
import subprocess
import sys
import threading
from datetime import date
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
    from tkcalendar import DateEntry
except ImportError:  # pragma: no cover
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
        self.root.geometry("980x680")

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
        self.start_hour_var = tk.StringVar(value="20")
        self.end_hour_var = tk.StringVar(value="21")
        self.date_var = tk.StringVar(value=today_text)
        self.config_var = tk.StringVar(value=str(self.default_config))
        self.log_level_var = tk.StringVar(value="INFO")
        self.venue_display_to_value = self._load_venue_display_map(Path(self.config_var.get()))
        venue_values = list(self.venue_display_to_value.keys())
        self.venue_var = tk.StringVar(value=venue_values[0] if venue_values else "全部场馆")

        self.once_var = tk.BooleanVar(value=True)
        self.email_alert_var = tk.BooleanVar(value=False)

        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=4)
        ttk.Label(row1, text="Sport").pack(side=tk.LEFT)
        ttk.Combobox(row1, textvariable=self.sport_var, values=["badminton", "pingpong"], width=14, state="readonly").pack(side=tk.LEFT, padx=(8, 16))
        ttk.Label(row1, text="Venue").pack(side=tk.LEFT)
        self.venue_combo = ttk.Combobox(row1, textvariable=self.venue_var, values=venue_values, width=28, state="readonly")
        self.venue_combo.pack(side=tk.LEFT, padx=(8, 16))
        ttk.Label(row1, text="Time Range").pack(side=tk.LEFT)
        hour_start_values = ["不限"] + [str(h) for h in range(0, 24)]
        hour_end_values = ["不限"] + [str(h) for h in range(1, 25)]
        ttk.Combobox(row1, textvariable=self.start_hour_var, values=hour_start_values, width=6, state="readonly").pack(side=tk.LEFT, padx=(8, 4))
        ttk.Label(row1, text="to").pack(side=tk.LEFT)
        ttk.Combobox(row1, textvariable=self.end_hour_var, values=hour_end_values, width=6, state="readonly").pack(side=tk.LEFT, padx=(4, 16))
        ttk.Label(row1, text="Date").pack(side=tk.LEFT)
        if DateEntry is not None:
            self.date_picker = DateEntry(row1, textvariable=self.date_var, date_pattern="yyyy-mm-dd", width=12)
            self.date_picker.pack(side=tk.LEFT, padx=(8, 4))
        else:
            ttk.Entry(row1, textvariable=self.date_var, width=14).pack(side=tk.LEFT, padx=(8, 4))
        ttk.Button(row1, text="Today", command=self.set_today).pack(side=tk.LEFT)

        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=4)
        ttk.Label(row2, text="Config").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.config_var, width=62).pack(side=tk.LEFT, padx=(8, 8))
        ttk.Button(row2, text="Reload Venues", command=self.reload_venue_options).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(row2, text="Log Level").pack(side=tk.LEFT)
        ttk.Combobox(row2, textvariable=self.log_level_var, values=["DEBUG", "INFO", "WARNING", "ERROR"], width=10, state="readonly").pack(side=tk.LEFT, padx=(8, 0))

        row3 = ttk.Frame(frame)
        row3.pack(fill=tk.X, pady=4)
        ttk.Checkbutton(row3, text="Once", variable=self.once_var).pack(side=tk.LEFT)
        ttk.Checkbutton(row3, text="Email Alert", variable=self.email_alert_var).pack(side=tk.LEFT, padx=(12, 0))

        self.start_btn = ttk.Button(row3, text="Start", command=self.start_monitor)
        self.start_btn.pack(side=tk.LEFT, padx=(16, 6))
        self.stop_btn = ttk.Button(row3, text="Stop", command=self.stop_monitor)
        self.stop_btn.pack(side=tk.LEFT)

        ttk.Button(row3, text="Open Snapshot HTML", command=self.open_snapshot_html).pack(side=tk.LEFT, padx=(16, 6))
        ttk.Button(row3, text="Open Snapshot PNG", command=self.open_snapshot_png).pack(side=tk.LEFT)

        row4 = ttk.Frame(frame)
        row4.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(row4, text="Enable Startup", command=self.enable_startup).pack(side=tk.LEFT)
        ttk.Button(row4, text="Disable Startup", command=self.disable_startup).pack(side=tk.LEFT, padx=(6, 12))
        self.startup_status_var = tk.StringVar(value="Startup: unknown")
        ttk.Label(row4, textvariable=self.startup_status_var).pack(side=tk.LEFT)

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

        result = {"全部场馆": ""}
        sorted_items = sorted(code_map.items(), key=lambda x: (not str(x[0]).isdigit(), int(x[0]) if str(x[0]).isdigit() else 9999, str(x[0])))
        for code, name in sorted_items:
            result[f"{code} - {name}"] = code
        return result

    def reload_venue_options(self) -> None:
        cfg_text = self.config_var.get().strip() or str(self.default_config)
        self.venue_display_to_value = self._load_venue_display_map(Path(cfg_text))
        venue_values = list(self.venue_display_to_value.keys())
        self.venue_combo.configure(values=venue_values)

        current = self.venue_var.get().strip()
        if current not in self.venue_display_to_value:
            self.venue_var.set(venue_values[0] if venue_values else "")

        self._append_log("Venue options reloaded from config.\n")

    def set_today(self) -> None:
        self.date_var.set(date.today().strftime("%Y-%m-%d"))

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
        cmd = [sys.executable, str(self.monitor_script), "--config", self.config_var.get().strip() or str(self.default_config)]

        sport = self.sport_var.get().strip() or "badminton"
        cmd.extend(["--sport", sport])

        venue_key = self.venue_var.get().strip()
        venue = self.venue_display_to_value.get(venue_key, "")
        if venue:
            cmd.extend(["--venue", venue])

        start_txt = self.start_hour_var.get().strip()
        end_txt = self.end_hour_var.get().strip()
        unlimited = "不限"
        if start_txt == unlimited and end_txt == unlimited:
            pass
        elif start_txt != unlimited and end_txt != unlimited:
            try:
                start_hour = int(start_txt)
                end_hour = int(end_txt)
            except ValueError as exc:
                raise ValueError("Time Range must be integer hours.") from exc
            if not (0 <= start_hour <= 23 and 1 <= end_hour <= 24 and start_hour < end_hour):
                raise ValueError("Time Range requires start < end, start in [0,23], end in [1,24].")
            cmd.extend(["--time-range", f"{start_hour:02d}:00-{end_hour:02d}:00"])
        else:
            raise ValueError("Please choose both start and end hours, or both set to 不限.")

        date_text = self.date_var.get().strip()
        if date_text:
            cmd.extend(["--date", date_text])

        if self.once_var.get():
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

        try:
            cmd = self._build_command()
        except ValueError as exc:
            messagebox.showerror("Invalid Parameters", str(exc))
            return
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
    app = MonitorGuiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
