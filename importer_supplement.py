"""
importer_supplement.py — 補充欄位 Excel 匯入邏輯

支援兩種模式：
  A. 固定格式（auto）
     系統自動偵測常見 key 欄名稱（文物登錄號、metadataID、登錄號、archieveID），
     其餘欄位全部視為補充欄位，無需使用者手動設定。
     適合：每次格式相同的例行補充檔。

  B. 自由映射（manual）
     使用者在介面上手動指定 key 欄、補充欄位、各欄位中文標籤。
     適合：來自不同來源、欄位名稱每次不同的補充檔。

重複匯入行為：
  預設直接覆蓋（on_duplicate='overwrite'）。
  同一筆 metadataID 的同一欄位若已有值，新值直接取代舊值。

補充欄位獨立存放在 supplement_fields 表，查詢時自動合併顯示。
"""

import re
from datetime import datetime

import pandas as pd

from db import get_conn

# 系統自動偵測的 key 欄候選名稱（固定格式模式用）
AUTO_KEY_CANDIDATES = [
    "文物登錄號", "metadataID", "metadata_id",
    "登錄號", "archieveID", "archieve_id", "id",
]

# 固定格式模式下，自動排除的非補充欄位
AUTO_SKIP_COLS = ["上傳排序", "序號", "排序", "編號"]


def _make_field_code(col_name: str) -> str:
    """將欄位名稱轉為合法的 field_code（英數底線）"""
    if re.match(r'^[a-zA-Z][a-zA-Z0-9_]*$', col_name):
        return col_name
    safe = re.sub(r'[^\w]', '_', col_name)[:30]
    return f"supp_{safe}"


def _read_file(file_or_path) -> tuple[pd.DataFrame, str]:
    """讀取 Excel，回傳 (df, sheet_name)，同時接受路徑或 file-like object"""
    xl = pd.ExcelFile(file_or_path)
    sheet = xl.sheet_names[0]
    df = pd.read_excel(file_or_path, sheet_name=sheet, dtype=str).fillna("")
    return df, sheet


def preview_supplement(file_or_path) -> dict:
    """
    Step 1：讀取補充檔，自動偵測模式，回傳欄位清單與建議設定。

    回傳 dict：
      sheet          — 工作表名稱
      columns        — 所有欄位名稱
      total          — 總列數
      sample_rows    — 前5筆（list of dict）
      auto_mode      — True 表示成功自動偵測到 key 欄
      auto_key_col   — 自動偵測到的 key 欄名稱（auto_mode=True 時有效）
      auto_value_cols — 自動偵測到的補充欄位列表
      df             — DataFrame（供後續步驟使用）
    """
    df, sheet = _read_file(file_or_path)
    columns = list(df.columns)

    # 嘗試自動偵測 key 欄
    auto_key_col = None
    for candidate in AUTO_KEY_CANDIDATES:
        if candidate in columns:
            auto_key_col = candidate
            break

    auto_value_cols = []
    if auto_key_col:
        auto_value_cols = [
            c for c in columns
            if c != auto_key_col and c not in AUTO_SKIP_COLS
        ]

    return {
        "sheet": sheet,
        "columns": columns,
        "total": len(df),
        "sample_rows": df.head(5).to_dict(orient="records"),
        "auto_mode": auto_key_col is not None,
        "auto_key_col": auto_key_col,
        "auto_value_cols": auto_value_cols,
        "df": df,
    }


def validate_supplement(
    df: pd.DataFrame,
    key_col: str,
    value_cols: list[str],
) -> dict:
    """
    Step 2：比對補充檔的 key 欄與主資料庫，回傳配對統計。

    回傳 dict：
      total          — 補充檔總列數
      matched        — 能對應到主資料庫的筆數
      unmatched      — 對應不到的筆數
      unmatched_ids  — 對應不到的 key 值（最多20個）
      will_overwrite — 各欄位將被覆蓋的筆數（已有舊值）
    """
    conn = get_conn()
    all_ids = set(
        r[0] for r in conn.execute("SELECT metadata_id FROM artifacts").fetchall()
    )

    keys = df[key_col].astype(str).str.strip()
    matched_mask = keys.isin(all_ids)

    will_overwrite = {}
    for vc in value_cols:
        fc = _make_field_code(vc)
        count = conn.execute(
            "SELECT COUNT(*) FROM supplement_fields WHERE field_code=?", (fc,)
        ).fetchone()[0]
        if count > 0:
            will_overwrite[vc] = count

    conn.close()
    return {
        "total": len(df),
        "matched": int(matched_mask.sum()),
        "unmatched": int((~matched_mask).sum()),
        "unmatched_ids": keys[~matched_mask].tolist()[:20],
        "will_overwrite": will_overwrite,
    }


def run_supplement(
    df: pd.DataFrame,
    key_col: str,
    value_cols: list[str],
    col_labels: dict[str, str],
    batch_label: str,
    on_duplicate: str = "overwrite",   # ← 預設直接覆蓋
    progress_callback=None,
) -> dict:
    """
    Step 3：將補充欄位值寫入 supplement_fields 表。

    on_duplicate:
        'overwrite'（預設）— 直接覆蓋舊值
        'skip'            — 保留舊值，跳過

    col_labels: {欄位名稱: 中文標籤}
    """
    conn = get_conn()
    now = datetime.now().isoformat()

    all_ids = set(
        r[0] for r in conn.execute("SELECT metadata_id FROM artifacts").fetchall()
    )

    inserted = updated = skipped = unmatched = 0
    total = len(df)

    for i, (_, row) in enumerate(df.iterrows()):
        if progress_callback:
            progress_callback(i + 1, total, f"補充第 {i+1}/{total} 筆")

        mid = str(row[key_col]).strip()
        if mid not in all_ids:
            unmatched += 1
            continue

        for vc in value_cols:
            field_code = _make_field_code(vc)
            field_label = col_labels.get(vc, vc)
            val = str(row.get(vc, "")).strip()

            # 更新欄位定義表
            conn.execute(
                """INSERT INTO supplement_field_defs (field_code, field_label, created_at)
                   VALUES (?,?,?)
                   ON CONFLICT(field_code) DO UPDATE SET field_label=excluded.field_label""",
                (field_code, field_label, now)
            )

            existing = conn.execute(
                "SELECT id FROM supplement_fields WHERE metadata_id=? AND field_code=?",
                (mid, field_code)
            ).fetchone()

            if existing:
                if on_duplicate == "overwrite":
                    conn.execute(
                        """UPDATE supplement_fields
                           SET field_value=?, field_label=?, batch_label=?, imported_at=?
                           WHERE metadata_id=? AND field_code=?""",
                        (val, field_label, batch_label, now, mid, field_code)
                    )
                    updated += 1
                else:
                    skipped += 1
            else:
                conn.execute(
                    """INSERT INTO supplement_fields
                       (metadata_id, field_code, field_label, field_value, batch_label, imported_at)
                       VALUES (?,?,?,?,?,?)""",
                    (mid, field_code, field_label, val, batch_label, now)
                )
                inserted += 1

    conn.commit()

    conn.execute(
        """INSERT INTO import_log
           (import_type, batch_label, filename, imported_at,
            total_rows, inserted, updated, skipped, unmatched, notes)
           VALUES (?,?,?,?,?,?,?,?,?,?)""",
        ("supplement", batch_label, "（上傳）", now,
         total, inserted, updated, skipped, unmatched,
         f"補充欄位：{', '.join(col_labels.get(v, v) for v in value_cols)}")
    )
    conn.commit()
    conn.close()

    return {
        "success": True,
        "total": total,
        "inserted": inserted,
        "updated": updated,
        "skipped": skipped,
        "unmatched": unmatched,
        "fields_added": [col_labels.get(v, v) for v in value_cols],
    }


def get_supplement_field_defs() -> list[dict]:
    """回傳所有已定義的補充欄位"""
    conn = get_conn()
    rows = conn.execute(
        "SELECT field_code, field_label, created_at FROM supplement_field_defs ORDER BY created_at"
    ).fetchall()
    conn.close()
    return [{"field_code": r[0], "field_label": r[1], "created_at": r[2]} for r in rows]
