#!/usr/bin/env python3
"""将 price_db.json 迁移到 SQLite"""
import json, sqlite3, shutil
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
JSON_PATH = DATA_DIR / "price_db.json"
DB_PATH = DATA_DIR / "price.db"

def init_db():
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model TEXT NOT NULL,
            name TEXT DEFAULT '',
            unit_price REAL DEFAULT 0,
            retail_price REAL DEFAULT 0,
            brand TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS price_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action TEXT NOT NULL,
            model TEXT,
            old_value TEXT,
            new_value TEXT,
            operator TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
    """)
    conn.commit()
    conn.close()

def _to_float(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0

def migrate():
    if not JSON_PATH.exists():
        print("price_db.json not found, skipping migration")
        return
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        items = json.load(f)
    
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("PRAGMA journal_mode=WAL")
    count = 0
    for item in items:
        model = str(item.get('型号', '')).strip()
        if not model:
            continue
        conn.execute(
            "INSERT OR IGNORE INTO products (model, name, unit_price, retail_price, brand) VALUES (?,?,?,?,?)",
            (model,
             str(item.get('名称', '')).strip(),
             _to_float(item.get('单价', 0)),
             _to_float(item.get('面价', 0)),
             str(item.get('产地/厂家', '')).strip()))
        count += 1
    conn.commit()
    conn.close()
    
    shutil.copy2(str(JSON_PATH), str(JSON_PATH) + ".bak")
    print(f"Migrated {count} products to SQLite. Backup: price_db.json.bak")

if __name__ == "__main__":
    init_db()
    migrate()
