"""
app.py — 電業文物清單研究分析工具
執行方式：streamlit run app.py
"""
import json
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from db import init_db, get_conn
from importer_main import preview_import, run_import, update_artifact
from importer_supplement import (
    preview_supplement, validate_supplement,
    run_supplement, get_supplement_field_defs,
)

# ── 初始化 ──────────────────────────────────────────────────────────────────
try:
    init_db()
except Exception as _db_err:
    st.error(
        f"**資料庫連線失敗**：{type(_db_err).__name__}: {_db_err}\n\n"
        "請確認 Streamlit Secrets 已正確設定 `DB_HOST` / `DB_PASSWORD`，"
        "或 `DATABASE_URL`（需含 `?sslmode=require`）。"
    )
    st.stop()

st.set_page_config(
    page_title="電業文物清單研究分析工具",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 側邊欄導覽 ───────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📋 文物清單工具")
    st.caption("內部研究分析專用")
    st.divider()
    page = st.radio(
        "功能選單",
        ["🏠 總覽", "📥 主清單匯入", "➕ 補充欄位匯入", "🔍 搜尋與篩選", "📊 交叉分析", "📤 匯出"],
        label_visibility="collapsed",
    )
    st.divider()

    # 顯示資料庫現有筆數
    try:
        conn = get_conn()
        total_artifacts = conn.execute("SELECT COUNT(*) FROM artifacts").fetchone()[0]
        supp_fields = conn.execute("SELECT COUNT(DISTINCT field_code) FROM supplement_field_defs").fetchone()[0]
        conn.close()
        st.metric("資料庫文物筆數", f"{total_artifacts:,}")
        st.metric("已新增補充欄位", f"{supp_fields} 種")
    except Exception:
        st.caption("資料庫尚未初始化")


# ════════════════════════════════════════════════════════════════════════════
# 頁面：總覽
# ════════════════════════════════════════════════════════════════════════════
if page == "🏠 總覽":
    st.title("電業文物清單研究分析工具")
    st.caption("內部研究分析專用・不對外公開")

    conn = get_conn()
    total = conn.execute("SELECT COUNT(*) FROM artifacts").fetchone()[0]

    if total == 0:
        st.info("資料庫目前沒有資料。請先到「主清單匯入」上傳 Excel 清單。")
    else:
        col1, col2, col3, col4 = st.columns(4)

        cond_counts = dict(conn.execute(
            """SELECT core_fields->>'conditions 保存狀況', COUNT(*)
               FROM artifacts GROUP BY 1"""
        ).fetchall())

        with col1:
            st.metric("總文物筆數", f"{total:,}")
        with col2:
            bad = sum(v for k, v in cond_counts.items() if k and "不良" in str(k))
            st.metric("狀況不良", bad, delta_color="inverse")
        with col3:
            serious = sum(v for k, v in cond_counts.items() if k and "嚴重" in str(k))
            st.metric("損傷嚴重", serious, delta_color="inverse")

        supp_defs = get_supplement_field_defs()
        with col4:
            st.metric("補充欄位種類", len(supp_defs))

        st.divider()

        # 匯入紀錄
        st.subheader("匯入紀錄")
        logs = conn.execute(
            """SELECT import_type, batch_label, filename, imported_at,
                      total_rows, inserted, updated, skipped, unmatched, notes
               FROM import_log ORDER BY imported_at DESC LIMIT 20"""
        ).fetchall()

        if logs:
            df_log = pd.DataFrame(logs, columns=[
                "類型", "批次標籤", "來源檔案", "匯入時間",
                "總筆數", "新增", "更新", "跳過", "未對應", "備註"
            ])
            df_log["類型"] = df_log["類型"].map({"main": "主清單", "supplement": "補充欄位"})
            st.dataframe(df_log, use_container_width=True, hide_index=True)
        else:
            st.caption("尚無匯入紀錄")

        # 補充欄位定義
        if supp_defs:
            st.subheader("已新增的補充欄位")
            df_defs = pd.DataFrame(supp_defs)
            df_defs.columns = ["欄位代碼", "中文標籤", "建立時間"]
            st.dataframe(df_defs, use_container_width=True, hide_index=True)

    conn.close()


# ════════════════════════════════════════════════════════════════════════════
# 頁面：主清單匯入
# ════════════════════════════════════════════════════════════════════════════
elif page == "📥 主清單匯入":
    st.title("📥 主清單匯入")
    st.caption("上傳 metadata Excel 檔案，自動解析欄位並寫入資料庫。")

    uploaded_files = st.file_uploader(
        "選擇 Excel 檔案（.xlsx，可多選）", type=["xlsx"], accept_multiple_files=True
    )

    if uploaded_files:
        # 逐檔預覽
        tmp_dir = Path("data")
        tmp_dir.mkdir(exist_ok=True)

        previews = []
        for uf in uploaded_files:
            tmp_path = tmp_dir / f"tmp_{uf.name}"
            tmp_path.write_bytes(uf.read())
            with st.spinner(f"讀取 {uf.name}..."):
                p = preview_import(str(tmp_path))
            p["_tmp_path"] = str(tmp_path)
            p["_filename"] = uf.name
            previews.append(p)

        # 顯示各檔摘要表
        st.subheader(f"共選取 {len(previews)} 個檔案")
        summary_rows = []
        for p in previews:
            warns = "；".join(p["warnings"]) if p["warnings"] else "—"
            summary_rows.append({
                "檔案名稱": p["_filename"],
                "工作表": p["sheet"],
                "文物筆數": p["total"],
                "欄位數": p["columns_found"],
                "警告": warns,
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        # 展開各檔詳細預覽
        for p in previews:
            with st.expander(f"📄 {p['_filename']} — 前5筆預覽"):
                if p["warnings"]:
                    for w in p["warnings"]:
                        st.warning(w)
                st.dataframe(pd.DataFrame(p["sample_rows"]), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("匯入設定")

        col_a, col_b = st.columns(2)
        with col_a:
            default_label = previews[0]["_filename"].replace(".xlsx", "") if len(previews) == 1 else ""
            batch_label = st.text_input(
                "批次標籤（用於識別這次匯入來源）",
                value=default_label,
                help="例如：113年第5期審查",
            )
        with col_b:
            on_dup = st.selectbox(
                "遇到已存在的 metadataID 時",
                options=["skip", "overwrite"],
                format_func=lambda x: "跳過（保留舊資料）" if x == "skip" else "覆蓋（以新資料取代）",
            )

        if st.button("✅ 確認匯入", type="primary", disabled=not batch_label):
            total_inserted = total_updated = total_skipped = 0
            for i, p in enumerate(previews):
                st.caption(f"正在匯入第 {i+1}/{len(previews)} 個檔案：{p['_filename']}")
                bar = st.progress(0, text="準備中...")

                def cb(cur, tot, msg, _bar=bar):
                    _bar.progress(cur / tot, text=msg)

                result = run_import(p["_tmp_path"], batch_label, on_dup, cb)
                bar.progress(1.0, text="完成！")

                if result["success"]:
                    total_inserted += result["inserted"]
                    total_updated  += result["updated"]
                    total_skipped  += result["skipped"]
                else:
                    st.error(f"{p['_filename']} 匯入失敗：{result.get('error')}")

            st.success(
                f"全部匯入完成！共新增 {total_inserted} 筆、"
                f"更新 {total_updated} 筆、跳過 {total_skipped} 筆"
            )


# ════════════════════════════════════════════════════════════════════════════
# 頁面：補充欄位匯入
# ════════════════════════════════════════════════════════════════════════════
elif page == "➕ 補充欄位匯入":
    st.title("➕ 補充欄位匯入")
    st.caption("上傳補充 Excel 檔案，系統自動比對 metadataID 並合併入資料庫。重複匯入直接覆蓋舊值。")

    uploaded = st.file_uploader("選擇補充欄位 Excel 檔案（.xlsx）", type=["xlsx"])

    if uploaded:
        with st.spinner("讀取檔案..."):
            preview = preview_supplement(uploaded)
            df_supp = preview["df"]

        c1, c2, c3 = st.columns(3)
        c1.metric("工作表", preview["sheet"])
        c2.metric("資料筆數", preview["total"])
        c3.metric("欄位數", len(preview["columns"]))
        st.caption("前5筆預覽：")
        st.dataframe(pd.DataFrame(preview["sample_rows"]), use_container_width=True, hide_index=True)
        st.divider()

        # 模式切換
        if preview["auto_mode"]:
            auto_hint = f"✅ 已自動偵測到 Key 欄：「{preview['auto_key_col']}」，補充欄位：{preview['auto_value_cols']}"
            default_mode_idx = 0
        else:
            auto_hint = "⚠️ 未偵測到常見 Key 欄位名稱，請改用自由映射模式手動指定。"
            default_mode_idx = 1

        mode = st.radio(
            "匯入模式",
            ["🔄 固定格式（自動偵測）", "✏️ 自由映射（手動設定）"],
            index=default_mode_idx,
            horizontal=True,
            help="固定格式：自動偵測 Key 欄（文物登錄號/metadataID 等），其餘欄位全部補充。\n自由映射：手動選擇 Key 欄與補充欄位，並自訂中文標籤。",
        )
        st.caption(auto_hint)
        st.divider()

        # 固定格式模式
        if mode == "🔄 固定格式（自動偵測）":
            if not preview["auto_mode"]:
                st.error("無法自動偵測 Key 欄，請改用「自由映射」模式。")
            else:
                key_col    = preview["auto_key_col"]
                value_cols = preview["auto_value_cols"]
                col_labels = {vc: vc for vc in value_cols}
                st.info(f"**Key 欄**：{key_col}　　**補充欄位**：{', '.join(value_cols)}")
                batch_label = st.text_input("批次標籤", value=uploaded.name.replace(".xlsx", ""), key="auto_batch")
                if st.button("🔍 執行比對", type="secondary", key="auto_validate"):
                    with st.spinner("比對中..."):
                        val = validate_supplement(df_supp, key_col, value_cols)
                    st.session_state.update({
                        "supp_validation": val, "supp_df": df_supp,
                        "supp_key_col": key_col, "supp_value_cols": value_cols,
                        "supp_col_labels": col_labels, "supp_batch": batch_label,
                    })

        # 自由映射模式
        else:
            st.subheader("Step 1　選擇 Key 欄位")
            default_key_idx = 0
            for i, c in enumerate(preview["columns"]):
                if c in ("文物登錄號", "metadataID", "登錄號", "archieveID"):
                    default_key_idx = i
                    break
            key_col = st.selectbox("Key 欄位（對應 metadataID）", options=preview["columns"], index=default_key_idx)

            st.subheader("Step 2　選擇補充欄位")
            available_cols = [c for c in preview["columns"] if c != key_col and c not in ("上傳排序", "序號")]
            value_cols = st.multiselect("補充欄位（可多選）", options=available_cols, default=available_cols)

            col_labels = {}
            if value_cols:
                st.subheader("Step 3　設定欄位中文標籤")
                for vc in value_cols:
                    col_labels[vc] = st.text_input(f"「{vc}」的標籤", value=vc, key=f"lbl_{vc}")

            batch_label = st.text_input("批次標籤", value=uploaded.name.replace(".xlsx", ""), key="manual_batch")

            if value_cols and st.button("🔍 執行比對", type="secondary", key="manual_validate"):
                with st.spinner("比對中..."):
                    val = validate_supplement(df_supp, key_col, value_cols)
                st.session_state.update({
                    "supp_validation": val, "supp_df": df_supp,
                    "supp_key_col": key_col, "supp_value_cols": value_cols,
                    "supp_col_labels": col_labels, "supp_batch": batch_label,
                })

        # 比對結果與確認匯入（兩種模式共用）
        if "supp_validation" in st.session_state:
            val   = st.session_state["supp_validation"]
            v_val = st.session_state["supp_value_cols"]
            st.divider()
            st.subheader("比對結果")
            r1, r2, r3 = st.columns(3)
            r1.metric("補充檔總筆數", val["total"])
            r2.metric("✅ 對應到主資料", val["matched"])
            r3.metric("❌ 對應不到（略過）", val["unmatched"])

            if val.get("will_overwrite"):
                overwrite_info = "、".join(f"「{k}」{v} 筆" for k, v in val["will_overwrite"].items())
                st.info(f"以下欄位已有舊值，匯入後直接覆蓋：{overwrite_info}")

            if val["unmatched_ids"]:
                with st.expander(f"查看對應不到的 {val['unmatched']} 個 ID"):
                    st.code("\n".join(val["unmatched_ids"]))

            batch_label_final = st.session_state.get("supp_batch", "")
            can_import = val["matched"] > 0 and batch_label_final

            if st.button(
                f"✅ 確認匯入（{val['matched']} 筆 × {len(v_val)} 個欄位，重複直接覆蓋）",
                type="primary", disabled=not can_import,
            ):
                bar = st.progress(0, text="準備中...")
                def cb(cur, tot, msg):
                    bar.progress(cur / tot, text=msg)

                result = run_supplement(
                    df=st.session_state["supp_df"],
                    key_col=st.session_state["supp_key_col"],
                    value_cols=v_val,
                    col_labels=st.session_state["supp_col_labels"],
                    batch_label=batch_label_final,
                    on_duplicate="overwrite",
                    progress_callback=cb,
                )
                bar.progress(1.0, text="完成！")
                if result["success"]:
                    st.success(
                        f"補充完成！新增 {result['inserted']} 筆、"
                        f"覆蓋更新 {result['updated']} 筆、"
                        f"未對應略過 {result['unmatched']} 筆"
                    )
                    st.caption(f"已補充欄位：{', '.join(result['fields_added'])}")
                    for k in ["supp_validation","supp_df","supp_key_col","supp_value_cols","supp_col_labels","supp_batch"]:
                        st.session_state.pop(k, None)
                else:
                    st.error("匯入失敗")


# ════════════════════════════════════════════════════════════════════════════
# 頁面：搜尋與篩選
# ════════════════════════════════════════════════════════════════════════════
elif page == "🔍 搜尋與篩選":
    st.title("🔍 搜尋與篩選")

    conn = get_conn()

    # 取得所有補充欄位定義（用於顯示）
    supp_defs = {d["field_code"]: d["field_label"] for d in get_supplement_field_defs()}

    # ── 側邊欄篩選器 ──────────────────────────────────────────────────────
    with st.sidebar:
        st.subheader("篩選條件")

        # 保存狀況
        cond_vals = [r[0] for r in conn.execute(
            "SELECT DISTINCT core_fields->>'conditions 保存狀況' FROM artifacts WHERE core_fields->>'conditions 保存狀況' IS NOT NULL ORDER BY 1"
        ).fetchall()]
        sel_cond = st.multiselect("保存狀況", cond_vals)

        # 典藏類型
        dtype_vals = [r[0] for r in conn.execute(
            "SELECT DISTINCT core_fields->>'dataType 典藏類型' FROM artifacts WHERE core_fields->>'dataType 典藏類型' IS NOT NULL ORDER BY 1"
        ).fetchall()]
        sel_dtype = st.multiselect("典藏類型", dtype_vals)

        # 系統別
        energy_vals = [r[0] for r in conn.execute(
            "SELECT DISTINCT core_fields->>'energyType 系統別' FROM artifacts WHERE core_fields->>'energyType 系統別' IS NOT NULL ORDER BY 1"
        ).fetchall()]
        sel_energy = st.multiselect("系統別", energy_vals)

        # 關鍵字標籤（Top 50）
        tag_vals = [r[0] for r in conn.execute(
            "SELECT tag, COUNT(*) as n FROM tags WHERE field_code='keywords' GROUP BY tag ORDER BY n DESC LIMIT 50"
        ).fetchall()]
        sel_tags = st.multiselect("關鍵字標籤", tag_vals)

        # 補充欄位篩選
        supp_filters = {}
        if supp_defs:
            st.divider()
            st.caption("補充欄位篩選")
            for fc, fl in supp_defs.items():
                vals = [r[0] for r in conn.execute(
                    "SELECT DISTINCT field_value FROM supplement_fields WHERE field_code=? AND field_value != '' ORDER BY 1",
                    (fc,)
                ).fetchall()]
                if vals:
                    sel = st.multiselect(fl, vals, key=f"supp_{fc}")
                    if sel:
                        supp_filters[fc] = sel

        # 批次標籤
        batch_vals = [r[0] for r in conn.execute(
            "SELECT DISTINCT batch_label FROM artifacts ORDER BY 1"
        ).fetchall()]
        sel_batch = st.multiselect("來源批次", batch_vals)

    # ── 搜尋框 ────────────────────────────────────────────────────────────
    search_query = st.text_input("🔍 全文搜尋（文物名稱、簡述、文化意義、關鍵詞）", placeholder="輸入關鍵字...")

    # ── 組建查詢 ──────────────────────────────────────────────────────────
    where_clauses = []
    params = []

    if search_query.strip():
        # FTS 搜尋（tsvector）
        fts_ids = [r[0] for r in conn.execute(
            "SELECT metadata_id FROM artifacts_fts WHERE search_vector @@ plainto_tsquery('simple', ?)",
            (search_query,)
        ).fetchall()]
        if not fts_ids:
            # fallback: ILIKE 搜尋（中文友善）
            like = f"%{search_query}%"
            fts_ids = [r[0] for r in conn.execute(
                """SELECT metadata_id FROM artifacts WHERE
                   core_fields->>'mainTitle 文物名稱' ILIKE ? OR
                   core_fields->>'abstract 文物綜合簡述' ILIKE ? OR
                   core_fields->>'significance 文化意義' ILIKE ?""",
                (like, like, like)
            ).fetchall()]
        if fts_ids:
            placeholders = ",".join("?" * len(fts_ids))
            where_clauses.append(f"a.metadata_id IN ({placeholders})")
            params.extend(fts_ids)
        else:
            where_clauses.append("1=0")

    if sel_cond:
        ph = ",".join("?" * len(sel_cond))
        where_clauses.append(f"a.core_fields->>'conditions 保存狀況' IN ({ph})")
        params.extend(sel_cond)

    if sel_dtype:
        ph = ",".join("?" * len(sel_dtype))
        where_clauses.append(f"a.core_fields->>'dataType 典藏類型' IN ({ph})")
        params.extend(sel_dtype)

    if sel_energy:
        ph = ",".join("?" * len(sel_energy))
        where_clauses.append(f"a.core_fields->>'energyType 系統別' IN ({ph})")
        params.extend(sel_energy)

    if sel_tags:
        for tag in sel_tags:
            where_clauses.append(
                "EXISTS (SELECT 1 FROM tags t WHERE t.metadata_id=a.metadata_id AND t.tag=?)"
            )
            params.append(tag)

    if sel_batch:
        ph = ",".join("?" * len(sel_batch))
        where_clauses.append(f"a.batch_label IN ({ph})")
        params.extend(sel_batch)

    # 補充欄位篩選
    for fc, sel_vals in supp_filters.items():
        ph = ",".join("?" * len(sel_vals))
        where_clauses.append(
            f"EXISTS (SELECT 1 FROM supplement_fields sf WHERE sf.metadata_id=a.metadata_id AND sf.field_code=? AND sf.field_value IN ({ph}))"
        )
        params.append(fc)
        params.extend(sel_vals)

    where_sql = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""

    # ── 執行查詢 ──────────────────────────────────────────────────────────
    count_sql = f"SELECT COUNT(*) FROM artifacts a {where_sql}"
    total_results = conn.execute(count_sql, params).fetchone()[0]

    st.caption(f"共找到 **{total_results}** 筆")

    data_sql = f"""
        SELECT a.metadata_id,
               a.core_fields->>'mainTitle 文物名稱' AS 文物名稱,
               a.core_fields->>'dataType 典藏類型'  AS 典藏類型,
               a.core_fields->>'conditions 保存狀況' AS 保存狀況,
               a.core_fields->>'keywords+ 關鍵詞'   AS 關鍵詞,
               LEFT(a.core_fields->>'abstract 文物綜合簡述', 80) AS 簡述,
               a.batch_label AS 來源批次
        FROM artifacts a {where_sql}
        LIMIT 200
    """
    rows = conn.execute(data_sql, params).fetchall()

    if rows:
        df_results = pd.DataFrame(rows, columns=["metadataID", "文物名稱", "典藏類型", "保存狀況", "關鍵詞", "簡述（節錄）", "來源批次"])

        # 若有補充欄位，合併顯示
        if supp_defs:
            for fc, fl in supp_defs.items():
                supp_map = dict(conn.execute(
                    "SELECT metadata_id, field_value FROM supplement_fields WHERE field_code=?", (fc,)
                ).fetchall())
                df_results[fl] = df_results["metadataID"].map(supp_map)

        st.dataframe(df_results, use_container_width=True, hide_index=True)

        # 點擊詳情（顯示選取筆的全欄位，支援編輯）
        st.divider()
        selected_id = st.selectbox("查看完整詳情", ["（不選）"] + list(df_results["metadataID"]))
        if selected_id and selected_id != "（不選）":
            row = conn.execute("SELECT core_fields FROM artifacts WHERE metadata_id=?", (selected_id,)).fetchone()
            if row:
                core = row[0]  # psycopg2 自動將 JSONB 反序列化為 dict
                supp_rows = conn.execute(
                    "SELECT field_code, field_label, field_value FROM supplement_fields WHERE metadata_id=?",
                    (selected_id,)
                ).fetchall()

                edit_key = f"edit_{selected_id}"

                # 儲存成功提示（rerun 後顯示）
                if st.session_state.pop(f"save_ok_{selected_id}", False):
                    st.success("✅ 已儲存！")

                hdr_col, btn_col = st.columns([7, 1])
                hdr_col.caption(f"**{selected_id}**")
                if not st.session_state.get(edit_key):
                    if btn_col.button("✏️ 編輯", key=f"open_{selected_id}"):
                        st.session_state[edit_key] = True
                        st.rerun()

                if st.session_state.get(edit_key):
                    # ── 編輯模式 ──────────────────────────────────────────
                    st.caption("欄位值可直接點擊修改，完成後按「儲存」。")

                    core_rows = [
                        (k, str(v) if v and str(v) != "None" else "")
                        for k, v in core.items()
                        if not k.startswith("_")
                    ]
                    df_edit = pd.DataFrame(core_rows, columns=["欄位", "值"])
                    edited_df = st.data_editor(
                        df_edit,
                        column_config={"欄位": st.column_config.TextColumn(disabled=True)},
                        use_container_width=True,
                        hide_index=True,
                        key=f"de_{selected_id}",
                    )

                    supp_new = {}
                    if supp_rows:
                        st.subheader("補充欄位")
                        for fc, fl, fv in supp_rows:
                            supp_new[fc] = (fl, st.text_input(
                                fl, value=str(fv or ""), key=f"se_{selected_id}_{fc}"
                            ))

                    c_save, c_cancel, _ = st.columns([1, 1, 5])
                    if c_save.button("💾 儲存", type="primary", key=f"save_{selected_id}"):
                        new_core = dict(zip(edited_df["欄位"], edited_df["值"]))
                        update_artifact(selected_id, new_core, supp_new or None)
                        st.session_state[edit_key] = False
                        st.session_state[f"save_ok_{selected_id}"] = True
                        st.rerun()
                    if c_cancel.button("✖️ 取消", key=f"cancel_{selected_id}"):
                        st.session_state[edit_key] = False
                        st.rerun()

                else:
                    # ── 檢視模式 ──────────────────────────────────────────
                    all_fields = {k: v for k, v in core.items() if v and str(v) != "None" and not k.startswith("_")}
                    for _, fl, fv in supp_rows:
                        all_fields[f"[補充] {fl}"] = fv
                    df_detail = pd.DataFrame(
                        list(all_fields.items()),
                        columns=["欄位", "值"]
                    )
                    st.dataframe(df_detail, use_container_width=True, hide_index=True)
    else:
        st.info("沒有符合條件的資料")

    conn.close()


# ════════════════════════════════════════════════════════════════════════════
# 頁面：交叉分析
# ════════════════════════════════════════════════════════════════════════════
elif page == "📊 交叉分析":
    st.title("📊 交叉分析")

    conn = get_conn()
    supp_defs = {d["field_code"]: d["field_label"] for d in get_supplement_field_defs()}

    PIVOT_FIELDS = {
        "典藏類型":  "core_fields->>'dataType 典藏類型'",
        "典藏次類型": "core_fields->>'subType 典藏次類型'",
        "保存狀況":  "core_fields->>'conditions 保存狀況'",
        "系統別":   "core_fields->>'energyType 系統別'",
        "主要材質":  "core_fields->>'material 主要材質'",
        "來源批次":  "batch_label",
    }
    # 加入補充欄位
    for fc, fl in supp_defs.items():
        PIVOT_FIELDS[f"[補充] {fl}"] = None  # 特殊處理

    field_names = list(PIVOT_FIELDS.keys())

    col_a, col_b = st.columns(2)
    with col_a:
        row_field = st.selectbox("列（Row）欄位", field_names, index=0)
    with col_b:
        col_field = st.selectbox("欄（Column）欄位", field_names, index=2)

    if st.button("產生交叉統計表", type="primary"):
        def get_vals(field_name):
            if field_name.startswith("[補充]"):
                fc = next((k for k, v in supp_defs.items() if f"[補充] {v}" == field_name), None)
                if not fc:
                    return {}
                rows = conn.execute(
                    "SELECT metadata_id, field_value FROM supplement_fields WHERE field_code=?", (fc,)
                ).fetchall()
                return {r[0]: r[1] for r in rows}
            else:
                expr = PIVOT_FIELDS[field_name]
                rows = conn.execute(f"SELECT metadata_id, {expr} FROM artifacts").fetchall()
                return {r[0]: r[1] for r in rows}

        row_map = get_vals(row_field)
        col_map = get_vals(col_field)

        all_ids = [r[0] for r in conn.execute("SELECT metadata_id FROM artifacts").fetchall()]

        data = []
        for mid in all_ids:
            rv = row_map.get(mid, "（未填）") or "（未填）"
            cv = col_map.get(mid, "（未填）") or "（未填）"
            data.append({"row": rv, "col": cv})

        df_pivot = pd.DataFrame(data)
        pivot = pd.crosstab(df_pivot["row"], df_pivot["col"], margins=True, margins_name="合計")
        pivot.index.name = f"{row_field} \\ {col_field}"

        st.dataframe(pivot, use_container_width=True)

        # 匯出
        buf = io.BytesIO()
        pivot.to_excel(buf, engine="openpyxl")
        st.download_button(
            "⬇ 下載交叉統計表（Excel）",
            data=buf.getvalue(),
            file_name=f"crossTab_{row_field}x{col_field}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    conn.close()


# ════════════════════════════════════════════════════════════════════════════
# 頁面：匯出
# ════════════════════════════════════════════════════════════════════════════
elif page == "📤 匯出":
    st.title("📤 自訂欄位匯出")
    st.caption("勾選需要的欄位，依目前篩選條件匯出 Excel。")

    conn = get_conn()
    supp_defs = {d["field_code"]: d["field_label"] for d in get_supplement_field_defs()}

    EXPORT_COLS = {
        "metadataID（詮釋資料識別碼）": "metadata_id",
        "文物名稱":   "core_fields->>'mainTitle 文物名稱'",
        "典藏類型":   "core_fields->>'dataType 典藏類型'",
        "典藏次類型": "core_fields->>'subType 典藏次類型'",
        "保存狀況":   "core_fields->>'conditions 保存狀況'",
        "文物綜合簡述": "core_fields->>'abstract 文物綜合簡述'",
        "文化意義":   "core_fields->>'significance 文化意義'",
        "關鍵詞":    "core_fields->>'keywords+ 關鍵詞'",
        "主要材質":  "core_fields->>'material 主要材質'",
        "起始西元年": "core_fields->>'dateNameYearStart 起始西元年'",
        "系統別":    "core_fields->>'energyType 系統別'",
        "來源批次":  "batch_label",
    }

    # ── 側邊欄篩選器 ──────────────────────────────────────────────────────
    with st.sidebar:
        st.subheader("篩選條件")

        cond_vals = [r[0] for r in conn.execute(
            "SELECT DISTINCT core_fields->>'conditions 保存狀況' FROM artifacts WHERE core_fields->>'conditions 保存狀況' IS NOT NULL ORDER BY 1"
        ).fetchall()]
        sel_cond = st.multiselect("保存狀況", cond_vals, key="exp_cond")

        dtype_vals = [r[0] for r in conn.execute(
            "SELECT DISTINCT core_fields->>'dataType 典藏類型' FROM artifacts WHERE core_fields->>'dataType 典藏類型' IS NOT NULL ORDER BY 1"
        ).fetchall()]
        sel_dtype = st.multiselect("典藏類型", dtype_vals, key="exp_dtype")

        energy_vals = [r[0] for r in conn.execute(
            "SELECT DISTINCT core_fields->>'energyType 系統別' FROM artifacts WHERE core_fields->>'energyType 系統別' IS NOT NULL ORDER BY 1"
        ).fetchall()]
        sel_energy = st.multiselect("系統別", energy_vals, key="exp_energy")

        batch_vals = [r[0] for r in conn.execute(
            "SELECT DISTINCT batch_label FROM artifacts ORDER BY 1"
        ).fetchall()]
        sel_batch = st.multiselect("來源批次", batch_vals, key="exp_batch")

    # ── 組建 WHERE clause ─────────────────────────────────────────────────
    where_clauses = []
    params = []

    if sel_cond:
        ph = ",".join("?" * len(sel_cond))
        where_clauses.append(f"core_fields->>'conditions 保存狀況' IN ({ph})")
        params.extend(sel_cond)
    if sel_dtype:
        ph = ",".join("?" * len(sel_dtype))
        where_clauses.append(f"core_fields->>'dataType 典藏類型' IN ({ph})")
        params.extend(sel_dtype)
    if sel_energy:
        ph = ",".join("?" * len(sel_energy))
        where_clauses.append(f"core_fields->>'energyType 系統別' IN ({ph})")
        params.extend(sel_energy)
    if sel_batch:
        ph = ",".join("?" * len(sel_batch))
        where_clauses.append(f"batch_label IN ({ph})")
        params.extend(sel_batch)

    where_sql = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""

    total_export = conn.execute(f"SELECT COUNT(*) FROM artifacts {where_sql}", params).fetchone()[0]
    st.caption(f"目前篩選結果：**{total_export:,}** 筆")

    selected_cols = st.multiselect(
        "選擇要匯出的欄位",
        list(EXPORT_COLS.keys()),
        default=["metadataID（詮釋資料識別碼）", "文物名稱", "典藏類型", "保存狀況", "文物綜合簡述"],
    )

    # 補充欄位選項
    supp_export = []
    if supp_defs:
        st.subheader("補充欄位")
        supp_export = st.multiselect(
            "選擇補充欄位",
            list(supp_defs.values()),
            default=list(supp_defs.values()),
        )

    if selected_cols and st.button("產生匯出檔", type="primary"):
        exprs = [EXPORT_COLS[c] for c in selected_cols]
        sql = f"SELECT {', '.join(exprs)} FROM artifacts {where_sql}"
        rows = conn.execute(sql, params).fetchall()

        df_export = pd.DataFrame(rows, columns=selected_cols)

        # 合併補充欄位
        for fc, fl in supp_defs.items():
            if fl in supp_export:
                smap = dict(conn.execute(
                    "SELECT metadata_id, field_value FROM supplement_fields WHERE field_code=?", (fc,)
                ).fetchall())
                id_col = "metadataID（詮釋資料識別碼）" if "metadataID（詮釋資料識別碼）" in df_export.columns else df_export.columns[0]
                df_export[fl] = df_export[id_col].map(smap)

        buf = io.BytesIO()
        df_export.to_excel(buf, index=False, engine="openpyxl")
        st.download_button(
            f"⬇ 下載 Excel（{len(df_export)} 筆 × {len(df_export.columns)} 欄）",
            data=buf.getvalue(),
            file_name="export_artifacts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.dataframe(df_export.head(20), use_container_width=True, hide_index=True)

    conn.close()
