#!/usr/bin/env python3
"""Schema-only Aronium database profile helper. Exports no transaction rows."""
import json
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG = HERE / "config.json"

def main():
    if not CONFIG.exists():
        raise SystemExit("Copy config.example.json to config.json first.")
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    db = Path(os.path.expandvars(os.path.expanduser(cfg["database_path"]))).resolve()
    output = Path(os.path.expandvars(os.path.expanduser(cfg.get("schema_profile_path", "schema-profile.json")))).resolve()
    if not db.is_file():
        raise SystemExit("Database not found: " + str(db))
    tokens = {
        "sale": ("sale", "receipt", "ticket", "order", "document", "transaction"),
        "sale_line": ("line", "detail", "item", "position"),
        "payment": ("payment", "tender", "settlement"),
        "refund": ("refund", "return", "void", "cancel"),
        "product": ("product", "item", "article", "stock"),
        "cashier": ("cashier", "user", "operator", "employee"),
    }
    uri = db.as_uri() + "?mode=ro"
    results = []
    with sqlite3.connect(uri, uri=True, timeout=5) as conn:
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA query_only = ON")
        rows = conn.execute("SELECT name,type FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY name").fetchall()
        for row in rows:
            name = str(row["name"])
            quoted = '"' + name.replace('"', '""') + '"'
            try:
                count = int(conn.execute("SELECT COUNT(*) FROM " + quoted).fetchone()[0])
            except sqlite3.Error:
                count = None
            scores = {role: sum(1 for token in words if token in name.lower()) for role, words in tokens.items()}
            candidates = sorted([{"entity_type": k, "score": v} for k, v in scores.items() if v], key=lambda x: (-x["score"], x["entity_type"]))
            cols = conn.execute("PRAGMA table_info(" + quoted + ")").fetchall()
            results.append({
                "table": name, "type": row["type"], "row_count": count,
                "candidate_roles": candidates[:3],
                "columns": [{"name": c["name"], "type": c["type"], "primary_key_position": c["pk"]} for c in cols]
            })
    report = {
        "report_type": "schema-profile-only",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "shop_id": cfg.get("shop_id", "TEST"),
        "source_id": cfg.get("source_id", "TEST-TILL1"),
        "database_file_name": db.name,
        "row_data_included": False,
        "tables": results
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print("Profile written to: " + str(output))
    print("Contains table/column names, row counts and candidate-role scores only; no transaction rows.")

if __name__ == "__main__":
    main()
