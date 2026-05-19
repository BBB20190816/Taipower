"""
db.py — 資料庫初始化與共用查詢工具
使用 PostgreSQL (Supabase)
連線字串從 Streamlit secrets["DATABASE_URL"] 或環境變數 DATABASE_URL 讀取
"""
import os

import psycopg2
import psycopg2.extras


def _get_url() -> str:
    try:
        import streamlit as st
        return st.secrets["DATABASE_URL"]
    except Exception:
        return os.environ.get("DATABASE_URL", "")


class _ConnWrapper:
    """
    讓 psycopg2 connection 的介面與原本 sqlite3 一致。
    - execute() 自動將 ? 轉為 %s（psycopg2 的佔位符）
    - 每次 execute() 建立新 cursor，回傳 cursor（可呼叫 .fetchone() / .fetchall()）
    """

    def __init__(self, conn):
        self._conn = conn

    def execute(self, sql: str, params=()):
        sql = sql.replace("?", "%s")
        cur = self._conn.cursor()
        cur.execute(sql, params if params else None)
        return cur

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()


def get_conn() -> _ConnWrapper:
    conn = psycopg2.connect(_get_url())
    return _ConnWrapper(conn)


def init_db():
    """建立所有資料表（若不存在）"""
    raw = psycopg2.connect(_get_url())
    cur = raw.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS artifacts (
            id            SERIAL PRIMARY KEY,
            metadata_id   TEXT   NOT NULL UNIQUE,
            batch_label   TEXT   NOT NULL,
            imported_at   TEXT   NOT NULL,
            core_fields   JSONB  NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS tags (
            id          SERIAL PRIMARY KEY,
            metadata_id TEXT NOT NULL REFERENCES artifacts(metadata_id) ON DELETE CASCADE,
            field_code  TEXT NOT NULL,
            tag         TEXT NOT NULL
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_tags_tag ON tags(tag)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_tags_mid ON tags(metadata_id)")

    cur.execute("""
        CREATE TABLE IF NOT EXISTS supplement_fields (
            id            SERIAL PRIMARY KEY,
            metadata_id   TEXT NOT NULL REFERENCES artifacts(metadata_id) ON DELETE CASCADE,
            field_code    TEXT NOT NULL,
            field_label   TEXT NOT NULL,
            field_value   TEXT,
            batch_label   TEXT NOT NULL,
            imported_at   TEXT NOT NULL,
            UNIQUE(metadata_id, field_code)
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_supp_mid   ON supplement_fields(metadata_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_supp_field ON supplement_fields(field_code)")

    cur.execute("""
        CREATE TABLE IF NOT EXISTS supplement_field_defs (
            field_code  TEXT PRIMARY KEY,
            field_label TEXT NOT NULL,
            created_at  TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS import_log (
            id           SERIAL PRIMARY KEY,
            import_type  TEXT NOT NULL,
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

    # 全文搜尋表（PostgreSQL tsvector）
    cur.execute("""
        CREATE TABLE IF NOT EXISTS artifacts_fts (
            metadata_id   TEXT PRIMARY KEY
                REFERENCES artifacts(metadata_id) ON DELETE CASCADE,
            search_vector TSVECTOR
        )
    """)
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_fts_vector ON artifacts_fts USING GIN(search_vector)"
    )

    cur.execute("""
        CREATE TABLE IF NOT EXISTS saved_filters (
            id          SERIAL PRIMARY KEY,
            name        TEXT NOT NULL UNIQUE,
            filter_json TEXT NOT NULL,
            created_at  TEXT NOT NULL
        )
    """)

    raw.commit()
    raw.close()
    print("[db] PostgreSQL 資料庫初始化完成")


if __name__ == "__main__":
    init_db()
