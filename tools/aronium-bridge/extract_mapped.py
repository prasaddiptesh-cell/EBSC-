#!/usr/bin/env python3
"""Extract explicitly mapped Aronium fields into the bridge queue; mappings default disabled."""
import json
import os
import sqlite3
from datetime import datetime, timezone
from contextlib import closing
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG = HERE / "config.json"
MAPPING = HERE / "mapping.json"

def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")

def p(value):
    return Path(os.path.expandvars(os.path.expanduser(value))).resolve()

def qi(name):
    if not name:
        raise ValueError("Empty SQL identifier in mapping")
    return '"' + str(name).replace('"', '""') + '"'

def main():
    if not CONFIG.exists() or not MAPPING.exists():
        raise SystemExit("Copy config.example.json to config.json and mapping.example.json to mapping.json.")
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    mapping = json.loads(MAPPING.read_text(encoding="utf-8"))
    db = p(cfg["database_path"])
    queue_path = p(cfg["queue_path"])
    if not db.is_file():
        raise SystemExit("Aronium database not found: " + str(db))
    entities = mapping.get("entities", {})
    inserted = 0
    stamp = now()
    with closing(sqlite3.connect(db.as_uri() + "?mode=ro", uri=True, timeout=5)) as source:
        source.row_factory = sqlite3.Row
        source.execute("PRAGMA query_only = ON")
        table_rows = source.execute("SELECT name FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%'").fetchall()
        tables = {str(r["name"]) for r in table_rows}
        columns = {t: {str(r["name"]) for r in source.execute("PRAGMA table_info(" + qi(t) + ")").fetchall()} for t in tables}
        active = []
        errors = []
        for kind, spec in entities.items():
            if not isinstance(spec, dict) or not spec.get("enabled", False):
                continue
            table = str(spec.get("table", ""))
            pk = str(spec.get("primary_key", ""))
            if table not in tables:
                errors.append(kind + ": table not found: " + table)
                continue
            if pk not in columns[table]:
                errors.append(kind + ": primary key not found: " + pk)
                continue
            fields = {str(k): str(v) for k, v in (spec.get("columns") or {}).items() if str(v)}
            time_col = str(spec.get("event_time_column", ""))
            parent_col = str(spec.get("parent_sale_key_column", ""))
            for col in list(fields.values()) + ([time_col] if time_col else []) + ([parent_col] if parent_col else []):
                if col not in columns[table]:
                    errors.append(kind + ": mapped column not found: " + col)
            active.append((kind, table, pk, fields, time_col, parent_col))
        if errors:
            raise SystemExit("Mapping validation failed:\n - " + "\n - ".join(errors))
        if not active:
            print("No mapping is enabled. No source records were extracted.")
            return
        queue_path.parent.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(queue_path, timeout=10)) as queue:
            queue.execute("PRAGMA journal_mode=WAL")
            queue.executescript("""
            CREATE TABLE IF NOT EXISTS bridge_meta(key TEXT PRIMARY KEY,value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS sync_queue(
              id INTEGER PRIMARY KEY AUTOINCREMENT, source_id TEXT NOT NULL,
              entity_type TEXT NOT NULL, source_record_id TEXT NOT NULL,
              idempotency_key TEXT NOT NULL UNIQUE, payload_json TEXT NOT NULL,
              status TEXT NOT NULL DEFAULT 'pending'
                CHECK(status IN ('pending','in_progress','sent','needs_review')),
              attempts INTEGER NOT NULL DEFAULT 0, created_at_utc TEXT NOT NULL,
              updated_at_utc TEXT NOT NULL, last_error TEXT);
            CREATE INDEX IF NOT EXISTS idx_queue_status_created ON sync_queue(status,created_at_utc);
            """)
            for kind, table, pk, fields, time_col, parent_col in active:
                selected = list(dict.fromkeys([pk] + ([time_col] if time_col else []) + ([parent_col] if parent_col else []) + list(fields.values())))
                sql = "SELECT " + ",".join(qi(x) for x in selected) + " FROM " + qi(table)
                for row in source.execute(sql):
                    record_id = row[pk]
                    if record_id is None:
                        continue
                    payload = {field: row[col] for field, col in fields.items()}
                    if parent_col and row[parent_col] is not None:
                        payload["parent_source_record_id"] = row[parent_col]
                    event_id = str(cfg["source_id"]) + ":" + kind + ":" + str(record_id)
                    event = {
                        "event_id": event_id, "source_id": str(cfg["source_id"]),
                        "shop_id": str(cfg["shop_id"]), "entity_type": kind,
                        "source_record_id": str(record_id),
                        "event_time": str(row[time_col]) if time_col and row[time_col] is not None else "",
                        "payload": payload, "mapping_version": mapping.get("version", 1)
                    }
                    cur = queue.execute(
                        "INSERT OR IGNORE INTO sync_queue(source_id,entity_type,source_record_id,idempotency_key,payload_json,status,attempts,created_at_utc,updated_at_utc) VALUES(?,?,?,?,?,'pending',0,?,?)",
                        (str(cfg["source_id"]), kind, str(record_id), event_id,
                         json.dumps(event, separators=(",", ":"), ensure_ascii=False, default=str), stamp, stamp)
                    )
                    inserted += cur.rowcount
            queue.commit()
    print("New mapped records queued: " + str(inserted))
    print("Existing IDs were not duplicated. This first version does not update changed existing events; changes need a correction workflow.")
    print("Cloud sync is separate. Do not enable uploads until schema mappings and sample totals are reconciled.")

if __name__ == "__main__":
    main()
