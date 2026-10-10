#!/usr/bin/env python3
"""Safe foundation for the Easy Buys Aronium sync bridge."""
from __future__ import annotations
import argparse
import json
import os
import sqlite3
import sys
import urllib.request
import urllib.error
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

def enqueue_json(args: argparse.Namespace) -> int:
    """Queue already-mapped events from a local JSON file; does not read POS rows."""
    cfg = load_config()
    input_path = path_of(args.file)
    try:
        data = json.loads(input_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"Cannot read event JSON: {exc}")
        return 2
    events = data.get("events") if isinstance(data, dict) else None
    if not isinstance(events, list):
        print("Input must be a JSON object containing an events array.")
        return 2
    source_id = str(cfg["source_id"])
    shop_id = str(cfg["shop_id"])
    now = utc_now()
    inserted = 0
    with queue_connection(path_of(cfg["queue_path"])) as conn:
        for ev in events:
            if not isinstance(ev, dict):
                continue
            required = ("event_id", "entity_type", "source_record_id")
            if any(not str(ev.get(k, "")).strip() for k in required):
                continue
            if ev.get("source_id") and str(ev["source_id"]) != source_id:
                print(f"Skipping event {ev.get('event_id')}: source_id does not match this configured till.")
                continue
            ev["source_id"] = source_id
            ev["shop_id"] = str(ev.get("shop_id") or shop_id)
            idem = str(ev["event_id"])
            payload = json.dumps(ev, separators=(",", ":"), ensure_ascii=False)
            try:
                conn.execute(
                    "INSERT INTO sync_queue(source_id,entity_type,source_record_id,idempotency_key,payload_json,status,attempts,created_at_utc,updated_at_utc) "
                    "VALUES(?,?,?,?,?,'pending',0,?,?)",
                    (source_id, str(ev["entity_type"]), str(ev["source_record_id"]), idem, payload, now, now)
                )
                inserted += 1
            except sqlite3.IntegrityError:
                # Existing event ID is intentionally not duplicated.
                continue
        conn.commit()
    print(f"Queued {inserted} new mapped events. Existing event IDs were not duplicated.")
    return 0


def sync_queue(args: argparse.Namespace) -> int:
    """Upload queued, mapped events to the dedicated POS endpoint."""
    cfg = load_config()
    if not cfg.get("cloud_sync_enabled", False):
        print("Cloud sync is disabled. No data was sent.")
        return 2
    base_url = str(cfg.get("cloud_api_base_url", "")).strip()
    token_env = str(cfg.get("cloud_device_token_env", "EBSC_BRIDGE_TOKEN"))
    token = os.environ.get(token_env, "")
    if not base_url.startswith("https://") or not token:
        print("Configure an HTTPS cloud_api_base_url and set the bridge token in the named environment variable.")
        return 2
    batch_limit = max(1, min(100, int(getattr(args, "limit", 50))))
    queue_path = path_of(cfg["queue_path"])
    with queue_connection(queue_path) as conn:
        rows = conn.execute(
            "SELECT id,source_id,entity_type,source_record_id,idempotency_key,payload_json,attempts "
            "FROM sync_queue WHERE status='pending' ORDER BY id LIMIT ?", (batch_limit,)
        ).fetchall()
    if not rows:
        print("No pending events.")
        return 0
    events = []
    for row in rows:
        try:
            ev = json.loads(row["payload_json"])
            # Ensure event identity used by the server is the persistent queue key.
            ev["event_id"] = row["idempotency_key"]
            ev["source_id"] = row["source_id"]
            ev["entity_type"] = row["entity_type"]
            ev["source_record_id"] = row["source_record_id"]
            events.append(ev)
        except (TypeError, json.JSONDecodeError):
            with queue_connection(queue_path) as conn:
                conn.execute("UPDATE sync_queue SET status='needs_review',last_error=?,updated_at_utc=? WHERE id=?",
                             ("Invalid local JSON payload", utc_now(), row["id"]))
                conn.commit()
    if not events:
        print("No valid pending events in this batch.")
        return 2
    body = json.dumps({"kind": "pos_ingest", "bridge_token": token, "events": events}).encode("utf-8")
    request = urllib.request.Request(base_url, data=body, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=25) as response:
            result = json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, OSError, json.JSONDecodeError) as exc:
        # Preserve pending status and record attempt/error so a retry is safe.
        with queue_connection(queue_path) as conn:
            for row in rows:
                conn.execute("UPDATE sync_queue SET attempts=attempts+1,updated_at_utc=?,last_error=? WHERE id=? AND status='pending'",
                             (utc_now(), str(exc)[:1000], row["id"]))
            conn.commit()
        print(f"Upload failed; events remain queued for retry: {exc}")
        return 1
    if not result.get("ok"):
        with queue_connection(queue_path) as conn:
            for row in rows:
                conn.execute("UPDATE sync_queue SET attempts=attempts+1,updated_at_utc=?,last_error=? WHERE id=? AND status='pending'",
                             (utc_now(), str(result.get("error", "Cloud rejected batch"))[:1000], row["id"]))
            conn.commit()
        print(f"Cloud rejected batch; events remain queued: {result.get('error', 'unknown error')}")
        return 1
    statuses = {str(item.get("event_id")): str(item.get("status")) for item in result.get("results", []) if isinstance(item, dict)}
    rejected_ids = set()
    for item in result.get("rejected", []):
        if isinstance(item, dict) and item.get("event_id"):
            rejected_ids.add(str(item["event_id"]))
    sent = review = 0
    with queue_connection(queue_path) as conn:
        for row in rows:
            key = str(row["idempotency_key"])
            status = statuses.get(key)
            if status in ("accepted", "already_processed"):
                conn.execute("UPDATE sync_queue SET status='sent',attempts=attempts+1,updated_at_utc=?,last_error=NULL WHERE id=?",
                             (utc_now(), row["id"]))
                sent += 1
            elif key in rejected_ids or status is None:
                reason = "Rejected by cloud; review required" if key in rejected_ids else "Cloud response did not acknowledge this event"
                conn.execute("UPDATE sync_queue SET status='needs_review',attempts=attempts+1,updated_at_utc=?,last_error=? WHERE id=?",
                             (utc_now(), reason, row["id"]))
                review += 1
            else:
                conn.execute("UPDATE sync_queue SET attempts=attempts+1,updated_at_utc=?,last_error=? WHERE id=?",
                             (utc_now(), "Unexpected server acknowledgement: " + status, row["id"]))
        conn.commit()
    print(f"Cloud acknowledged: {sent}; moved to review: {review}.")
    if result.get("rejected"):
        print("Cloud rejections:")
        print(json.dumps(result["rejected"], indent=2))
    return 0 if review == 0 else 1


def watch_queue(args: argparse.Namespace) -> int:
    """Continuously attempt to upload queued mapped events until stopped."""
    import time
    interval = max(5, min(300, int(args.interval)))
    print(f"Queue sync loop started (every {interval}s). Press Ctrl+C to stop.")
    try:
        while True:
            result = sync_queue(argparse.Namespace(limit=args.limit))
            if result not in (0, 2):
                print("Some events remain queued or need review; the loop will retry pending events.")
            time.sleep(interval)
    except KeyboardInterrupt:
        print("\\nQueue sync loop stopped. Local queue is preserved.")
        return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Easy Buys Aronium bridge foundation")
    sub = parser.add_subparsers(dest="command", required=True)
    for name, help_text, handler in [
        ("validate", "Validate configuration and read-only database access", validate),
        ("inspect-schema", "Write schema-only report without transaction rows", inspect_schema),
        ("init-queue", "Initialize local durable queue", init_queue),
        ("status", "Show bridge configuration and queue status", status),
        ("enqueue-json", "Queue pre-mapped events from a local JSON file", enqueue_json),
        ("sync-queue", "Upload pending events to configured Easy Buys POS API", sync_queue),
        ("watch-queue", "Continuously retry pending queue uploads", watch_queue),
    ]:
        p = sub.add_parser(name, help=help_text)
        p.set_defaults(handler=handler)
        if name == "enqueue-json": p.add_argument("file", help="Path to a JSON file containing an events array")
        if name in ("sync-queue", "watch-queue"): p.add_argument("--limit", type=int, default=50, help="Maximum events per upload batch (1-100)")\n        if name == "watch-queue": p.add_argument("--interval", type=int, default=10, help="Seconds between retries (5-300)")
    args = parser.parse_args()
    try:
        return int(args.handler(args))
    except (FileNotFoundError, ValueError, json.JSONDecodeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
