#!/usr/bin/env python3
"""Safe foundation for the Easy Buys Aronium sync bridge."""
from __future__ import annotations
import argparse
import json
import os
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"

def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")

def load_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError("Copy config.example.json to config.json and edit it.")
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    required = ("shop_id", "source_id", "database_path", "queue_path")
    missing = [k for k in required if not str(cfg.get(k, "")).strip()]
    if missing:
        raise ValueError("Missing configuration fields: " + ", ".join(missing))
    return cfg

def path_of(value: str) -> Path:
    return Path(os.path.expandvars(os.path.expanduser(value))).resolve()

def connect_source(path: Path) -> sqlite3.Connection:
    if not path.is_file():
        raise FileNotFoundError(f"Aronium database not found: {path}")
    conn = sqlite3.connect(path.as_uri() + "?mode=ro", uri=True, timeout=5)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA query_only = ON")
    conn.execute("PRAGMA busy_timeout = 5000")
    return conn

def validate(_: argparse.Namespace) -> int:
    cfg = load_config()
    db = path_of(cfg["database_path"])
    print(f"Shop/source: {cfg['shop_id']} / {cfg['source_id']}")
    print(f"Database: {db}")
    print(f"Cloud sync enabled: {bool(cfg.get('cloud_sync_enabled', False))}")
    if not db.is_file():
        print("Database file not found. Confirm its location in Aronium Settings > Database.")
        return 2
    try:
        with connect_source(db) as conn:
            check = conn.execute("PRAGMA quick_check").fetchone()[0]
        print(f"Read-only database check: {check}")
        return 0 if check == "ok" else 3
    except sqlite3.Error as exc:
        print(f"Could not open source read-only: {exc}")
        return 4

def inspect_schema(_: argparse.Namespace) -> int:
    cfg = load_config()
    db = path_of(cfg["database_path"])
    out = path_of(cfg.get("schema_report_path", "schema-report.json"))
    try:
        with connect_source(db) as conn:
            rows = conn.execute(
                "SELECT name, type FROM sqlite_master "
                "WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' "
                "ORDER BY type,name"
            ).fetchall()
            tables = []
            for row in rows:
                name = str(row["name"])
                quoted = '"' + name.replace('"', '""') + '"'
                cols = conn.execute(f"PRAGMA table_info({quoted})").fetchall()
                tables.append({
                    "name": name, "type": row["type"],
                    "columns": [{"name": c["name"], "type": c["type"],
                                 "not_null": bool(c["notnull"]),
                                 "primary_key_position": c["pk"]} for c in cols]
                })
    except (OSError, sqlite3.Error) as exc:
        print(f"Schema inspection failed: {exc}")
        return 2
    report = {
        "report_type": "schema-only", "generated_at_utc": utc_now(),
        "shop_id": cfg["shop_id"], "source_id": cfg["source_id"],
        "database_file_name": db.name, "row_data_included": False,
        "tables": tables
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Found {len(tables)} tables/views. Schema-only report: {out}")
    print("No transaction or customer row data was exported.")
    return 0

def queue_connection(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript("""
      CREATE TABLE IF NOT EXISTS bridge_meta (
        key TEXT PRIMARY KEY, value TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS sync_queue (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_id TEXT NOT NULL, entity_type TEXT NOT NULL,
        source_record_id TEXT NOT NULL, idempotency_key TEXT NOT NULL UNIQUE,
        payload_json TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'pending'
          CHECK(status IN ('pending','in_progress','sent','needs_review')),
        attempts INTEGER NOT NULL DEFAULT 0,
        created_at_utc TEXT NOT NULL, updated_at_utc TEXT NOT NULL,
        last_error TEXT
      );
      CREATE INDEX IF NOT EXISTS idx_queue_status_created
        ON sync_queue(status, created_at_utc);
      CREATE INDEX IF NOT EXISTS idx_queue_source_record
        ON sync_queue(source_id, entity_type, source_record_id);
    """)
    conn.execute("INSERT OR IGNORE INTO bridge_meta(key,value) VALUES('schema_version','1')")
    conn.commit()
    return conn

def init_queue(_: argparse.Namespace) -> int:
    cfg = load_config()
    path = path_of(cfg["queue_path"])
    with queue_connection(path) as conn:
        count = conn.execute("SELECT COUNT(*) FROM sync_queue").fetchone()[0]
    print(f"Queue initialized: {path}; queued records: {count}")
    print("No cloud upload was attempted.")
    return 0

def status(_: argparse.Namespace) -> int:
    cfg = load_config()
    queue = path_of(cfg["queue_path"])
    print(f"Bridge: {cfg.get('bridge_name', 'Easy Buys Aronium Bridge')}")
    print(f"Shop/source: {cfg['shop_id']} / {cfg['source_id']}")
    print(f"Cloud sync: {'CONFIGURED BUT NOT IMPLEMENTED' if cfg.get('cloud_sync_enabled') else 'DISABLED'}")
    print(f"Queue: {queue}")
    if not queue.exists():
        print("Queue not initialized. Run: python bridge.py init-queue")
        return 0
    with queue_connection(queue) as conn:
        rows = conn.execute("SELECT status,COUNT(*) AS n FROM sync_queue GROUP BY status").fetchall()
    counts = {r["status"]: r["n"] for r in rows}
    for state in ("pending", "in_progress", "sent", "needs_review"):
        print(f"{state}: {counts.get(state, 0)}")
    print("Last successful cloud sync: not configured")
    return 0

def main() -> int:
    parser = argparse.ArgumentParser(description="Easy Buys Aronium bridge foundation")
    sub = parser.add_subparsers(dest="command", required=True)
    for name, help_text, handler in [
        ("validate", "Validate configuration and read-only database access", validate),
        ("inspect-schema", "Write schema-only report without transaction rows", inspect_schema),
        ("init-queue", "Initialize local durable queue", init_queue),
        ("status", "Show bridge configuration and queue status", status),
    ]:
        p = sub.add_parser(name, help=help_text)
        p.set_defaults(handler=handler)
    args = parser.parse_args()
    try:
        return int(args.handler(args))
    except (FileNotFoundError, ValueError, json.JSONDecodeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
