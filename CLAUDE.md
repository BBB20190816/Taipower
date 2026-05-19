# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案簡介

電業文物清單研究分析工具（內部研究分析專用）。管理台電文物 metadata Excel 清單，支援匯入、搜尋篩選、交叉分析、單筆編輯、匯出。

## 啟動方式

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 檔案架構

```
app.py                 — Streamlit 主應用（六頁面）
db.py                  — SQLite 初始化與 get_conn()
importer_main.py       — 主清單 Excel 匯入＋單筆更新
importer_supplement.py — 補充欄位 Excel 匯入
data/artifacts.db      — SQLite 資料庫（執行後產生）
```

## 資料庫 Schema 重點

- **`artifacts`**：每筆文物一列，所有欄位存為 `core_fields`（JSON 字串）
- **`tags`**：keywords / subjectMatter / placeName 以 `|` 拆分後的獨立標籤
- **`supplement_fields`**：補充欄位值，UNIQUE(metadata_id, field_code)
- **`supplement_field_defs`**：補充欄位定義（field_code, field_label）
- **`artifacts_fts`**：FTS5 虛擬表，content='' 需手動維護（`importer_main.py` 負責）
- **`import_log`**：所有匯入歷程記錄

SQLite JSON 查詢語法：`json_extract(core_fields, '$.\"mainTitle 文物名稱\"')`

## Excel 格式規範（主清單）

- 工作表名稱含 `metadata` 或 `維護` 即自動識別
- **雙層標頭**：Row0=群組名, Row1=欄位代碼, Row2=空白, Row3+=資料
- 欄位代碼去掉尾端 `?*+`（如 `keywords+` → `keywords`）
- 第 0 欄為序號，非空即為有效資料列

## 補充欄位匯入（importer_supplement.py）

- **Auto 模式**：自動偵測 key 欄（候選：`文物登錄號` / `metadataID` / `登錄號` / `archieveID`）
- **Manual 模式**：使用者手動指定 key 欄、補充欄位、中文標籤
- `_make_field_code()` 將欄名轉為合法 field_code；純英數保持原名，否則加 `supp_` 前綴

## 單筆編輯（update_artifact）

`importer_main.update_artifact(metadata_id, core_fields, supp_updates)` 寫回資料庫後自動重建 tags 與 FTS 索引。`supp_updates` 格式：`{field_code: (field_label, new_value)}`。

## 搜尋邏輯

FTS5 精確搜尋（`MATCH "query"`）→ 搜不到時 fallback LIKE（mainTitle / abstract / significance）。

## 匯出頁篩選

匯出頁側邊欄支援保存狀況 / 典藏類型 / 系統別 / 來源批次篩選，WHERE clause 動態組建後帶入 `conn.execute(sql, params)`。
