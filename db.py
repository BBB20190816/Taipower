"""
db.py — 資料庫初始化與共用查詢工具
使用 SQLite，包含主資料表、標籤索引表、欄位定義表、補充欄位表、匯入紀錄表
"""
import sqlite3
import os
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "artifacts.db"


def get_conn() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    """建立所有資料表（若不存在）"""
    conn = get_conn()
    cur = conn.cursor()

    # ── 主資料表：每筆文物一列，欄位以 JSON 儲存動態欄位 ──────────────
    cur.execute("""
        CREATE TABLE IF NOT EXISTS artifacts (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            metadata_id   TEXT    NOT NULL UNIQUE,   -- metadataID，主鍵
            batch_label   TEXT    NOT NULL,           -- 來源批次標籤
            imported_at   TEXT    NOT NULL,           -- 匯入時間 ISO8601
            core_fields   TEXT    NOT NULL            -- JSON：所有欄位值
        )
    """)

    # ── 標籤索引表：keywords / subjectMatter 拆分後的獨立標籤 ──────────
    cur.execute("""
        CREATE TABLE IF NOT EXISTS tags (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            metadata_id TEXT NOT NULL REFERENCES artifacts(metadata_id) ON DELETE CASCADE,
            field_code  TEXT NOT NULL,   -- 來源欄位，如 keywords
            tag         TEXT NOT NULL
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_tags_tag ON tags(tag)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_tags_mid ON tags(metadata_id)")

    # ── 補充欄位表：每次「補充欄位匯入」新增的自訂欄位值 ────────────────
    cur.execute("""
        CREATE TABLE IF NOT EXISTS supplement_fields (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            metadata_id   TEXT NOT NULL REFERENCES artifacts(metadata_id) ON DELETE CASCADE,
            field_code    TEXT NOT NULL,   -- 補充欄位代碼（使用者自訂）
            field_label   TEXT NOT NULL,   -- 補充欄位中文名稱
            field_value   TEXT,            -- 欄位值
            batch_label   TEXT NOT NULL,   -- 來源批次標籤
            imported_at   TEXT NOT NULL,
            UNIQUE(metadata_id, field_code)  -- 同欄位同筆資料只能有一個值
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_supp_mid ON supplement_fields(metadata_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_supp_field ON supplement_fields(field_code)")

    # ── 欄位定義表：記錄所有出現過的補充欄位 ────────────────────────────
    cur.execute("""
        CREATE TABLE IF NOT EXISTS supplement_field_defs (
            field_code  TEXT PRIMARY KEY,
            field_label TEXT NOT NULL,
            created_at  TEXT NOT NULL
        )
    """)

    # ── 匯入紀錄表 ────────────────────────────────────────────────────────
    cur.execute("""
        CREATE TABLE IF NOT EXISTS import_log (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            import_type  TEXT NOT NULL,   -- 'main' | 'supplement'
            batch_label  TEXT NOT NULL,
            filename     TEXT NOT NULL,
            imported_at  TEXT NOT NULL,
            total_rows   INTEGER,
            inserted     INTEGER,
            updated      INTEGER,
            skipped      INTEGER,
            unmatched    INTEGER,
            notes        TEXT
        )
    """)

    # ── 全文搜尋虛擬表（FTS5）────────────────────────────────────────────
    cur.execute("""
        CREATE VIRTUAL TABLE IF NOT EXISTS artifacts_fts USING fts5(
            metadata_id UNINDEXED,
            main_title,
            abstract,
            significance,
            keywords_raw,
            content='',
            tokenize='unicode61'
        )
    """)

    # ── 篩選快照表 ────────────────────────────────────────────────────────
    cur.execute("""
        CREATE TABLE IF NOT EXISTS saved_filters (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT NOT NULL UNIQUE,
            filter_json TEXT NOT NULL,
            created_at  TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()
    print(f"[db] 資料庫初始化完成：{DB_PATH}")


if __name__ == "__main__":
    init_db()
