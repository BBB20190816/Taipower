# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案簡介

電業文物清單研究分析工具（內部研究分析專用）。管理台電文物 metadata Excel 清單，支援匯入、搜尋篩選、交叉分析、單筆編輯、匯出。

## 啟動方式

```bash
pip install -r requirements.txt

# 本機需先建立 .streamlit/secrets.toml（參考 .streamlit/secrets.toml.example）
streamlit run app.py
```

連線字串從 `st.secrets["DATABASE_URL"]` 讀取，fallback 為環境變數 `DATABASE_URL`。

## 檔案架構

```
app.py                 — Streamlit 主應用（六頁面路由）
db.py                  — PostgreSQL 連線（_ConnWrapper）與 init_db()
importer_main.py       — 主清單 Excel 匯入＋單筆更新（update_artifact）
importer_supplement.py — 補充欄位 Excel 匯入
.streamlit/
  secrets.toml.example — 連線字串範本（secrets.toml 本身在 .gitignore）
```

## 資料庫（PostgreSQL / Supabase）

**Schema 重點：**
- `artifacts`：每筆文物一列，`core_fields JSONB` 儲存所有動態欄位
- `tags`：keywords / subjectMatter / placeName 以 `|` 拆分後的獨立標籤列
- `supplement_fields`：補充欄位值，UNIQUE(metadata_id, field_code)
- `supplement_field_defs`：補充欄位定義（field_code, field_label）
- `artifacts_fts`：tsvector 全文搜尋，`search_vector` 欄位 + GIN index
- `import_log`：所有匯入歷程記錄

**JSONB 查詢語法：**
```sql
core_fields->>'mainTitle 文物名稱'          -- 取文字值
core_fields->>'conditions 保存狀況' IS NOT NULL
```
欄位名稱含空格直接寫入 `->>'key'`，不需跳脫。

**連線介面（`_ConnWrapper`）：**
`db.get_conn()` 回傳 `_ConnWrapper`，介面與 sqlite3 相同：
- `conn.execute(sql, params)` — 自動將 `?` 轉為 `%s`，回傳 psycopg2 cursor
- `conn.commit()` / `conn.close()`
- JSONB 寫入時需用 `psycopg2.extras.Json(dict)` 包裝

## Excel 格式規範（主清單）

- 工作表名稱含 `metadata` 或 `維護` 即自動識別
- **雙層標頭**：Row0=群組名, Row1=欄位代碼, Row2=空白, Row3+=資料
- 欄位代碼去掉尾端 `?*+`（如 `keywords+` → `keywords`）
- 第 0 欄為序號，非空即為有效資料列

## 補充欄位匯入

- **Auto 模式**：自動偵測 key 欄（候選：`文物登錄號` / `metadataID` / `登錄號` / `archieveID`）
- **Manual 模式**：使用者手動指定 key 欄、補充欄位、中文標籤
- `_make_field_code()`：純英數保持原名，否則加 `supp_` 前綴

## 單筆編輯

`importer_main.update_artifact(metadata_id, core_fields, supp_updates)` 寫回後自動重建 tags 與 tsvector 索引。`supp_updates` 格式：`{field_code: (field_label, new_value)}`。

## 搜尋邏輯

`plainto_tsquery('simple', query)` 全文搜尋 → 無結果時 fallback `ILIKE %keyword%`（mainTitle / abstract / significance）。

## 匯出頁篩選

側邊欄支援保存狀況 / 典藏類型 / 系統別 / 來源批次篩選，WHERE clause 動態組建後帶入 `conn.execute(sql, params)`。

## 讀取 JSONB 欄位

psycopg2 自動將 JSONB 反序列化為 Python dict，**不需要 `json.loads()`**：
```python
row = conn.execute("SELECT core_fields FROM artifacts WHERE metadata_id=?", (mid,)).fetchone()
core = row[0]  # 已是 dict
```
