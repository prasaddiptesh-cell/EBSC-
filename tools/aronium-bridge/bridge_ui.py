#!/usr/bin/env python3
"""Simple Windows desktop interface for the Easy Buys Aronium bridge foundation."""
from __future__ import annotations
import json
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

HERE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG = HERE / "config.json"
EXAMPLE = HERE / "config.example.json"

class BridgeUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Easy Buys | Aronium Sync Bridge")
        self.geometry("850x620")
        self.minsize(720, 500)
        self.configure(padx=18, pady=16)
        ttk.Label(self, text="Easy Buys Aronium Sync Bridge", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(self, text="Local setup and diagnostics • cloud upload remains disabled in this foundation",
                  wraplength=800).pack(anchor="w", pady=(3, 14))
        self.summary = ttk.LabelFrame(self, text="Installation")
        self.summary.pack(fill="x", pady=(0, 12))
        self.install_label = ttk.Label(self.summary, text="Checking configuration…", padding=10)
        self.install_label.pack(anchor="w")
        actions = ttk.Frame(self)
        actions.pack(fill="x", pady=(0, 12))
        action_items = [
            ("Configure", self.open_config_dialog),
            ("Validate database", lambda: self.run_command("validate")),
            ("Inspect schema", lambda: self.run_command("inspect-schema")),
            ("Profile schema", lambda: self.run_command("profile-schema")),
            ("Initialize queue", lambda: self.run_command("init-queue")),
            ("Extract mapped data", lambda: self.run_command("extract-mapped")),
            ("Sync pending queue", lambda: self.run_command("sync-queue")),
            ("Refresh status", self.refresh_status),
        ]
        for index, (label, command) in enumerate(action_items):
            ttk.Button(actions, text=label, command=command).grid(
                row=index // 4, column=index % 4, sticky="ew", padx=4, pady=4
            )
        for column in range(4):
            actions.columnconfigure(column, weight=1)
        ttk.Label(self, text="Activity log", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.log = tk.Text(self, height=18, wrap="word", state="disabled", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, pady=(5, 0))
        self.refresh_status()

    def append_log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip() + "\n\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def refresh_status(self):
        if not CONFIG.exists():
            self.install_label.configure(text="Not configured. Select Configure to create your local config.")
            self.append_log("No config.json found. No POS database has been accessed.")
            return
        try:
            cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
            self.install_label.configure(
                text=f"Shop: {cfg.get('shop_name', cfg.get('shop_id', 'Not set'))}   |   "
                     f"Source: {cfg.get('source_id', 'Not set')}   |   "
                     f"Cloud upload: {'OFF' if not cfg.get('cloud_sync_enabled') else 'Configured but not implemented'}\n"
                     f"Database: {cfg.get('database_path', 'Not set')}"
            )
            self.run_command("status", background=False)
        except Exception as exc:
            self.install_label.configure(text="Could not read config.json. Use Configure to correct it.")
            self.append_log(f"Configuration error: {exc}")

    def open_config_dialog(self):
        if not CONFIG.exists():
            try:
                CONFIG.write_text(EXAMPLE.read_text(encoding="utf-8"), encoding="utf-8")
            except Exception as exc:
                messagebox.showerror("Configuration", str(exc))
                return
        try:
            cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
        except Exception as exc:
            messagebox.showerror("Configuration", f"Cannot parse config.json: {exc}")
            return
        win = tk.Toplevel(self)
        win.title("Configure this Aronium till")
        win.transient(self)
        win.grab_set()
        win.geometry("780x570")
        form = ttk.Frame(win, padding=16)
        form.pack(fill="both", expand=True)
        fields = [
            ("shop_id", "Shop ID (e.g. S1)"),
            ("source_id", "Unique source ID (e.g. S1-TILL1)"),
            ("shop_name", "Shop name"),
            ("register_name", "Till name"),
            ("database_path", "Aronium database path"),
            ("queue_path", "Local queue path"),
            ("schema_report_path", "Schema report path"),
            ("cloud_api_base_url", "Cloud web app URL (HTTPS)"),
            ("cloud_device_token_env", "Token environment variable name"),
        ]
        variables = {}
        for i, (key, label) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", pady=5)
            var = tk.StringVar(value=str(cfg.get(key, "")))
            variables[key] = var
            ttk.Entry(form, textvariable=var, width=58).grid(row=i, column=1, sticky="ew", padx=(12, 0), pady=5)
            if key == "database_path":
                def browse(v=var):
                    picked = filedialog.askopenfilename(
                        title="Select Aronium database", filetypes=[("SQLite database", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")]
                    )
                    if picked:
                        v.set(picked)
                ttk.Button(form, text="Browse…", command=browse).grid(row=i, column=2, padx=(6, 0))
        form.columnconfigure(1, weight=1)
        cloud_enabled = tk.BooleanVar(value=bool(cfg.get("cloud_sync_enabled", False)))
        cloud_enabled.set(False)
        ttk.Checkbutton(form, text="Cloud uploads (not yet enabled for this test build)",
                        variable=cloud_enabled, state="disabled").grid(row=len(fields), column=0, columnspan=3, sticky="w", pady=(10, 4))
        ttk.Label(form, text="Keep this off until a test endpoint and approved source mapping are ready. The bridge token must be stored in a Windows environment variable.",
                  wraplength=700).grid(row=len(fields)+1, column=0, columnspan=3, sticky="w", pady=(2, 10))
        def save():
            for key, var in variables.items():
                cfg[key] = var.get().strip()
            cfg["cloud_sync_enabled"] = False
            try:
                CONFIG.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
            except Exception as exc:
                messagebox.showerror("Save failed", str(exc), parent=win)
                return
            win.destroy()
            self.append_log("Saved local configuration. Cloud upload setting: " + ("ON" if cfg.get("cloud_sync_enabled") else "OFF") + ".")
            self.refresh_status()
        ttk.Button(form, text="Save configuration", command=save).grid(row=len(fields)+1, column=2, sticky="e")

    def run_command(self, command: str, background: bool = True):
        def work():
            try:
                if getattr(sys, "frozen", False):
                    cmd = [sys.executable, "--bridge-command", command]
                    display = f"$ EasyBuysAroniumBridge.exe {command}"
                else:
                    cmd = [sys.executable, str(HERE / "bridge.py"), command]
                    display = f"$ python bridge.py {command}"
                result = subprocess.run(cmd, cwd=str(HERE), capture_output=True, text=True, timeout=90)
                output = f"{display}\n{result.stdout}{result.stderr}\nExit code: {result.returncode}"
            except Exception as exc:
                output = f"Could not run {command}: {exc}"
            self.after(0, lambda: self.append_log(output))
        if background:
            threading.Thread(target=work, daemon=True).start()
        else:
            work()

if __name__ == "__main__":
    BridgeUI().mainloop()
