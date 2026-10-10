import importlib.util
import sqlite3
import tempfile
from contextlib import closing
import unittest
from pathlib import Path

MODULE_PATH = Path(__file__).resolve().parents[1] / "bridge.py"
spec = importlib.util.spec_from_file_location("ebsc_bridge", MODULE_PATH)
bridge = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bridge)

class BridgeFoundationTests(unittest.TestCase):
    def test_source_connection_is_read_only(self):
        with tempfile.TemporaryDirectory() as temp:
            db = Path(temp) / "pos.db"
            conn = sqlite3.connect(db)
            conn.execute("CREATE TABLE sample(id INTEGER PRIMARY KEY)")
            conn.commit()
            conn.close()
            with closing(bridge.connect_source(db)) as source:
                self.assertEqual(source.execute("SELECT COUNT(*) FROM sample").fetchone()[0], 0)
                with self.assertRaises(sqlite3.OperationalError):
                    source.execute("CREATE TABLE must_not_exist(id INTEGER)")
            with sqlite3.connect(db) as verify:
                names = [r[0] for r in verify.execute("SELECT name FROM sqlite_master WHERE type='table'")]
            self.assertEqual(names, ["sample"])

    def test_queue_has_unique_idempotency_key(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "queue.db"
            with closing(bridge.queue_connection(path)) as conn:
                conn.execute(
                    "INSERT INTO sync_queue(source_id,entity_type,source_record_id,idempotency_key,payload_json,created_at_utc,updated_at_utc) "
                    "VALUES(?,?,?,?,?,?,?)",
                    ("S1-TILL1", "sale", "sale-1", "S1-TILL1:sale:sale-1", "{}", "now", "now")
                )
                with self.assertRaises(sqlite3.IntegrityError):
                    conn.execute(
                        "INSERT INTO sync_queue(source_id,entity_type,source_record_id,idempotency_key,payload_json,created_at_utc,updated_at_utc) "
                        "VALUES(?,?,?,?,?,?,?)",
                        ("S1-TILL1", "sale", "sale-1-copy", "S1-TILL1:sale:sale-1", "{}", "now", "now")
                    )

    def test_queue_initializes_empty(self):
        with tempfile.TemporaryDirectory() as temp:
            with closing(bridge.queue_connection(Path(temp) / "queue.db")) as conn:
                self.assertEqual(conn.execute("SELECT COUNT(*) FROM sync_queue").fetchone()[0], 0)

if __name__ == "__main__":
    unittest.main()
