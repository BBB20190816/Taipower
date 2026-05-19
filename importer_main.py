"""
importer_main.py — 主清單（metadata Excel）匯入邏輯

流程：
1. 讀取 Excel，自動識別 metadata維護 工作表
2. 解析雙層標頭（Row0=群組, Row1=欄位代碼），以 Row1 為準
3. 過濾出有序號的資料列
4. 執行校驗（metadataID 唯一性）
5. 對 keywords、subjectMatter、placeName 執行 | 拆分
6. 寫入 artifacts 表與 tags 表
7. 更新 FTS 索引
8. 回傳匯入報告 dict
"""
import json
import sqlite3
from datetime import datetime
from typing import Optional

import pandas as pd

from db import get_conn

# 需要做 | 拆分的欄位（欄位代碼前綴比對）
TAG_SPLIT_FIELDS = ["keywords", "subjectMatter", "placeName"]

# 核心全文搜尋欄位（用於 FTS 索引）
FTS_FIELD_MAP = {
    "main_title":    "mainTitle",
    "abstract":      "abstract",
    "significance":  "significance",
    "keywords_raw":  "keywords",
}


def _parse_headers(df_raw: pd.DataFrame) -> dict[str, str]:
    """
    解析雙層標頭，回傳 {欄位代碼: 欄 index} 的映射。
    Row0 是群組名，Row1 是欄位代碼（英文），以 Row1 為準。
    """
    row1 = df_raw.iloc[1]
    mapping = {}  # col_index -> field_code
    for i, val in enumerate(row1):
        if pd.isna(val) or str(val).strip() == "":
            continue
        # 欄位代碼是第一個 \n 前的部分
        code = str(val).split("\n")[0].strip()
        # 去掉尾端的 ?、*、+ 等修飾符
        code = code.rstrip("?*+")
        if code:
            mapping[i] = code
    return mapping


def _find_metadata_sheet(xl: pd.ExcelFile) -> str:
    """自動找主資料工作表名稱"""
    for name in xl.sheet_names:
        if "metadata" in name.lower() or "維護" in name:
            return name
    return xl.sheet_names[0]


def _split_tags(value: str, field_code: str) -> list[str]:
    """以 | 拆分標籤，去除空白"""
    if not value or pd.isna(value):
        return []
    return [t.strip() for t in str(value).split("|") if t.strip()]


def preview_import(file_path: str) -> dict:
    """
    預覽匯入：不寫入資料庫，只回傳統計資訊供使用者確認。
    回傳 dict：{total, columns, sample_rows, validation_warnings}
    """
    xl = pd.ExcelFile(file_path)
    sheet = _find_metadata_sheet(xl)
    df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)

    col_map = _parse_headers(df_raw)

    # 從 Row3 開始是資料（Row0=群組, Row1=代碼, Row2=空白, Row3+=資料）
    df_data = df_raw.iloc[3:].reset_index(drop=True)

    # 找序號欄（第0欄）並過濾有序號的列
    df_data.columns = [col_map.get(i, f"_col{i}") for i in range(len(df_data.columns))]
    seq_col = df_data.columns[0]
    df_valid = df_data[
        df_data[seq_col].notna() &
        (df_data[seq_col].astype(str).str.strip() != "") &
        (df_data[seq_col].astype(str).str.strip() != "nan")
    ].copy()

    # 找 metadataID 欄
    mid_col = next((c for c in df_valid.columns if c.startswith("metadataID")), None)

    warnings = []
    if mid_col is None:
        warnings.append("⚠️ 找不到 metadataID 欄位，請確認欄位名稱")
    else:
        dup = df_valid[mid_col].duplicated(keep=False)
        if dup.sum() > 0:
            warnings.append(f"⚠️ metadataID 有 {dup.sum()} 筆重複")

    # 核心欄位空白率
    core_check = ["mainTitle", "abstract", "significance", "keywords", "conditions"]
    null_rates = {}
    for field in core_check:
        col = next((c for c in df_valid.columns if c.startswith(field)), None)
        if col:
            n = df_valid[col].isna().sum() + (df_valid[col].astype(str).str.strip() == "").sum()
            null_rates[field] = {"null": int(n), "pct": round(n / len(df_valid) * 100)}

    # 取前5筆樣本
    sample_cols = ["metadataID", "mainTitle", "conditions", "dataType"]
    sample_rows = []
    for _, row in df_valid.head(5).iterrows():
        r = {}
        for sc in sample_cols:
            col = next((c for c in df_valid.columns if c.startswith(sc)), None)
            r[sc] = str(row[col])[:60] if col and pd.notna(row[col]) else ""
        sample_rows.append(r)

    return {
        "sheet": sheet,
        "total": len(df_valid),
        "columns_found": len(col_map),
        "null_rates": null_rates,
        "warnings": warnings,
        "sample_rows": sample_rows,
    }


def run_import(
    file_path: str,
    batch_label: str,
    on_duplicate: str = "skip",      # 'skip' | 'overwrite'
    progress_callback=None,
) -> dict:
    """
    執行主清單匯入。

    on_duplicate:
        'skip'      — 遇到已存在的 metadataID 跳過
        'overwrite' — 遇到已存在的 metadataID 更新欄位值

    progress_callback: 可選，接受 (current, total, message) 的函式，供 Streamlit progress bar 使用
    """
    xl = pd.ExcelFile(file_path)
    sheet = _find_metadata_sheet(xl)
    df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
    col_map = _parse_headers(df_raw)

    df_data = df_raw.iloc[3:].reset_index(drop=True)
    df_data.columns = [col_map.get(i, f"_col{i}") for i in range(len(df_data.columns))]

    seq_col = df_data.columns[0]
    df_valid = df_data[
        df_data[seq_col].notna() &
        (df_data[seq_col].astype(str).str.strip() != "") &
        (df_data[seq_col].astype(str).str.strip() != "nan")
    ].copy().reset_index(drop=True)

    mid_col = next((c for c in df_valid.columns if c.startswith("metadataID")), None)
    if not mid_col:
        return {"success": False, "error": "找不到 metadataID 欄位"}

    conn = get_conn()
    now = datetime.now().isoformat()
    inserted = updated = skipped = 0
    total = len(df_valid)

    for idx, row in df_valid.iterrows():
        if progress_callback:
            progress_callback(idx + 1, total, f"匯入第 {idx+1}/{total} 筆")

        metadata_id = str(row[mid_col]).strip()
        if not metadata_id or metadata_id == "nan":
            skipped += 1
            continue

        # 將整列轉為 JSON（排除私有欄位）
        core = {
            k: (None if pd.isna(v) else str(v).strip())
            for k, v in row.items()
            if not k.startswith("_")
        }
        core_json = json.dumps(core, ensure_ascii=False)

        # 檢查是否已存在
        existing = conn.execute(
            "SELECT id FROM artifacts WHERE metadata_id = ?", (metadata_id,)
        ).fetchone()

        if existing:
            if on_duplicate == "overwrite":
                conn.execute(
                    "UPDATE artifacts SET core_fields=?, batch_label=?, imported_at=? WHERE metadata_id=?",
                    (core_json, batch_label, now, metadata_id)
                )
                # 清除舊標籤
                conn.execute("DELETE FROM tags WHERE metadata_id=?", (metadata_id,))
                updated += 1
            else:
                skipped += 1
                continue
        else:
            conn.execute(
                "INSERT INTO artifacts (metadata_id, batch_label, imported_at, core_fields) VALUES (?,?,?,?)",
                (metadata_id, batch_label, now, core_json)
            )
            inserted += 1

        # 寫入標籤索引
        for field_prefix in TAG_SPLIT_FIELDS:
            tag_col = next((c for c in df_valid.columns if c.startswith(field_prefix)), None)
            if tag_col and pd.notna(row.get(tag_col)):
                tags = _split_tags(str(row[tag_col]), field_prefix)
                for tag in tags:
                    conn.execute(
                        "INSERT INTO tags (metadata_id, field_code, tag) VALUES (?,?,?)",
                        (metadata_id, field_prefix, tag)
                    )

        # 更新 FTS 索引
        fts_values = {"metadata_id": metadata_id}
        for fts_col, field_prefix in FTS_FIELD_MAP.items():
            src = next((c for c in df_valid.columns if c.startswith(field_prefix)), None)
            fts_values[fts_col] = str(row[src]).strip() if src and pd.notna(row.get(src)) else ""

        # 刪除舊 FTS 記錄再插入（FTS5 content='' 需手動維護）
        conn.execute("DELETE FROM artifacts_fts WHERE metadata_id=?", (metadata_id,))
        conn.execute(
            "INSERT INTO artifacts_fts(metadata_id, main_title, abstract, significance, keywords_raw) VALUES(?,?,?,?,?)",
            (metadata_id, fts_values["main_title"], fts_values["abstract"],
             fts_values["significance"], fts_values["keywords_raw"])
        )

    conn.commit()

    # 寫入匯入紀錄
    conn.execute(
        """INSERT INTO import_log
           (import_type, batch_label, filename, imported_at, total_rows, inserted, updated, skipped, unmatched)
           VALUES (?,?,?,?,?,?,?,?,?)""",
        ("main", batch_label, str(file_path), now, total, inserted, updated, skipped, 0)
    )
    conn.commit()
    conn.close()

    return {
        "success": True,
        "total": total,
        "inserted": inserted,
        "updated": updated,
        "skipped": skipped,
        "batch_label": batch_label,
    }


def update_artifact(
    metadata_id: str,
    core_fields: dict,
    supp_updates: dict | None = None,
) -> dict:
    """
    更新單筆文物的核心欄位與補充欄位。

    supp_updates: {field_code: (field_label, new_value)}
    """
    conn = get_conn()
    now = datetime.now().isoformat()

    core_json = json.dumps(core_fields, ensure_ascii=False)
    conn.execute(
        "UPDATE artifacts SET core_fields=?, imported_at=? WHERE metadata_id=?",
        (core_json, now, metadata_id),
    )

    # 重建 tags（keywords / subjectMatter / placeName）
    conn.execute("DELETE FROM tags WHERE metadata_id=?", (metadata_id,))
    for field_prefix in TAG_SPLIT_FIELDS:
        col = next((k for k in core_fields if k.startswith(field_prefix)), None)
        if col and core_fields.get(col):
            for tag in _split_tags(str(core_fields[col]), field_prefix):
                conn.execute(
                    "INSERT INTO tags (metadata_id, field_code, tag) VALUES (?,?,?)",
                    (metadata_id, field_prefix, tag),
                )

    # 重建 FTS 索引
    fts_vals = {}
    for fts_col, field_prefix in FTS_FIELD_MAP.items():
        src = next((k for k in core_fields if k.startswith(field_prefix)), None)
        fts_vals[fts_col] = str(core_fields[src]).strip() if src and core_fields.get(src) else ""

    conn.execute("DELETE FROM artifacts_fts WHERE metadata_id=?", (metadata_id,))
    conn.execute(
        "INSERT INTO artifacts_fts(metadata_id, main_title, abstract, significance, keywords_raw) VALUES(?,?,?,?,?)",
        (metadata_id, fts_vals.get("main_title", ""), fts_vals.get("abstract", ""),
         fts_vals.get("significance", ""), fts_vals.get("keywords_raw", "")),
    )

    # 更新補充欄位
    if supp_updates:
        for field_code, (field_label, new_value) in supp_updates.items():
            conn.execute(
                """INSERT INTO supplement_fields
                   (metadata_id, field_code, field_label, field_value, batch_label, imported_at)
                   VALUES (?,?,?,?,?,?)
                   ON CONFLICT(metadata_id, field_code)
                   DO UPDATE SET field_value=excluded.field_value, imported_at=excluded.imported_at""",
                (metadata_id, field_code, field_label, new_value, "手動編輯", now),
            )

    conn.commit()
    conn.close()
    return {"success": True}
