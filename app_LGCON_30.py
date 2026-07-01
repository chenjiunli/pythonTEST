import streamlit as st
import pandas as pd
import io
import os
import re
from datetime import datetime
import zoneinfo

# 側邊欄強制保持展開
st.set_page_config(
    page_title="迅得 OHT LGCON&LGIO V2.0 (30套) 專案進度查詢", 
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': None
    }
)

# 網頁外觀純淨化與防拷貝機制
st.markdown("""
    <style>
    /* 隱藏數據表格右上角的所有控制按鈕 */
    [data-testid="stDataFrameToolbar"] { display: none !important; }
    /* 只隱藏主選單與頁尾 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    /* 資安防拷貝機制 */
    [data-testid="stDataFrame"], [data-testid="stMetric"], .stMarkdown {
        -webkit-user-select: none !important;
        -moz-user-select: none !important;
        -ms-user-select: none !important;
        user-select: none !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🛠️ 迅得 OHT LGCON&LGIO V2.0 (30套) 專案進度即時查詢面板")
st.write("🤝 本面板資料已自動同步，並依 Item 進行排序。您可於表格內利用滑鼠滾輪或手勢滑動瀏覽。")

# 快取檔案名稱
DATA_FILE = "latest_bom_data_30.pkl"
META_FILE = "bom_meta_info_30.pkl"  
PROG_FILE = "progress_data_30.pkl"  
ANOMALY_FILE = "anomaly_data_30.pkl" 

# 數字安全清洗器
def clean_numeric_values(val):
    if pd.isna(val): return 0
    val_str = str(val).strip().replace(',', '')
    match = re.search(r'-?\d+\.?\d*', val_str)
    if match:
        try:
            num = float(match.group())
            return int(num) if num > 0 else 0
        except:
            return 0
    return 0

# --- 初始化 Session State 防止無限重整迴圈 ---
if 'processed_file_id' not in st.session_state:
    st.session_state['processed_file_id'] = None

with st.sidebar:
    st.markdown("### ⚙️ 內部資料管理")
    password = st.text_input("管理員密碼：", type="password")
    
    if password == "1234":
        st.success("密碼正確！")
        uploaded_file = st.file_uploader("請上傳最新的 Excel 檔案 (包含 成本-30 與 進度 Sheet)：", type=["xlsx"])
        
        if uploaded_file is not None:
            current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
            if st.session_state['processed_file_id'] != current_file_id:
                progress_bar = st.sidebar.progress(0)
                status_text = st.sidebar.empty()
                
                try:
                    status_text.caption("⏳ [10%] 正在開啟 Excel 檔案結構...")
                    progress_bar.progress(10)
                    xls = pd.ExcelFile(uploaded_file)
                    
                    if "成本-30" in xls.sheet_names:
                        status_text.caption("⏳ [30%] 正在讀取 BOM 物料清單 (成本-30)...")
                        progress_bar.progress(30)
                        raw_df = pd.read_excel(xls, sheet_name="成本-30", header=None)
                        
                        # 1. 🚀 動態尋找頂部統計數據 (避免寫死列數)
                        sum_delivery = 0
                        sum_produce = 0
                        sum_stock = 0
                        for idx in range(min(10, len(raw_df))):
                            row_vals = [str(x).strip() for x in raw_df.iloc[idx].values]
                            for c_idx, val in enumerate(row_vals):
                                if "交貨數量" in val and c_idx + 1 < len(row_vals):
                                    sum_delivery = clean_numeric_values(raw_df.iloc[idx, c_idx + 1])
                                elif "生產數量" in val and c_idx + 1 < len(row_vals):
                                    sum_produce = clean_numeric_values(raw_df.iloc[idx, c_idx + 1])
                                elif ("備料數量" in val or "備貨數量" in val) and c_idx + 1 < len(row_vals):
                                    sum_stock = clean_numeric_values(raw_df.iloc[idx, c_idx + 1])

                        # 2. 🚀 智慧搜尋表頭列與欄位位置 (解決 Excel 欄位平移與隱藏問題)
                        header_row_idx = None
                        for idx in range(len(raw_df)):
                            row_str = [str(x).strip().upper() for x in raw_df.iloc[idx].values]
                            if "ITEM" in row_str and ("REFERENCE" in row_str or "PART NUMBER" in row_str):
                                header_row_idx = idx
                                break

                        if header_row_idx is not None:
                            header_vals = [str(x).strip().upper() for x in raw_df.iloc[header_row_idx].values]
                            
                            def get_col_index(keywords, default_val):
                                for kw in keywords:
                                    if kw.upper() in header_vals:
                                        return header_vals.index(kw.upper())
                                return default_val
                                
                            col_item = get_col_index(["Item"], 0)
                            col_part_num = get_col_index(["Part Number"], 1)
                            col_ref = get_col_index(["Reference", "Location", "零件位置"], 2)
                            col_mfr_pn = get_col_index(["Mfr_PN", "料號", "Manufacturer P/N"], 3)
                            col_part_name = get_col_index(["Part Name", "品名規格"], 8)
                            col_mfr_name = get_col_index(["Mfr_Name", "製造商"], 11)
                            col_shortage = get_col_index(["缺料", "缺貨數量"], 16)
                            col_delivery_date = get_col_index(["交期", "預計交期"], 27)
                        else:
                            col_item, col_part_num, col_ref, col_mfr_pn, col_part_name, col_mfr_name, col_shortage, col_delivery_date = 0, 1, 2, 3, 8, 11, 16, 27

                        # --- 🐛 除錯模式透視鏡 ---
                        st.sidebar.markdown("---")
                        st.sidebar.write("### 🐛 智慧欄位定位結果")
                        st.sidebar.code(f"表頭列數：第 {header_row_idx} 列\n"
                                        f"料號欄位索引：{col_mfr_pn}\n"
                                        f"零件位置(Reference)索引：{col_ref}\n"
                                        f"Part Name 索引：{col_part_name}")
                        st.sidebar.markdown("---")

                        # 3. 擷取核心 BOM 表資料
                        clean_df = pd.DataFrame({
                            "Item": raw_df.iloc[:, col_item],               
                            "Manufacturer_P/N": raw_df.iloc[:, col_mfr_pn],   
                            "Manufacture_Name": raw_df.iloc[:, col_mfr_name],  
                            "raw_shortage": raw_df.iloc[:, col_shortage],      
                            "交期": raw_df.iloc[:, col_delivery_date]               
                        })
                        
                        df_data = clean_df.dropna(subset=["Item"]).copy()
                        df_data["Item"] = df_data["Item"].astype(str).str.strip()
                        df_data = df_data[df_data["Item"].str.match(r'^\d+$', na=False)]
                        df_data["缺料"] = df_data["raw_shortage"].apply(clean_numeric_values)
                        
                        df_filtered = df_data[["Item", "Manufacturer_P/N", "Manufacture_Name", "缺料", "交期"]].copy()
                        
                        for col in df_filtered.columns:
                            if col == "交期":
                                df_filtered[col] = df_filtered[col].apply(
                                    lambda x: x.strftime('%Y/%m/%d') if isinstance(x, datetime) or hasattr(x, 'strftime') else str(x).strip()
                                )
                                df_filtered[col] = df_filtered[col].replace({"NaT": "", "nan": "", "None": ""})
                            elif col == "缺料":
                                df_filtered[col] = df_filtered[col].astype(str)
                            else:
                                df_filtered[col] = df_filtered[col].astype(str).str.strip()
                        
                        df_filtered['Item_num'] = pd.to_numeric(df_filtered['Item'], errors='coerce')
                        df_filtered = df_filtered.sort_values(by=['Item_num'], ascending=True).drop(columns=['Item_num'])
                        
                        tz_taibei = zoneinfo.ZoneInfo("Asia/Taipei")
                        current_time = datetime.now(tz_taibei).strftime("%Y-%m-%d %H:%M")
                        
                        new_meta = pd.DataFrame([{
                            "version": uploaded_file.name, 
                            "update_time": current_time,
                            "sum_delivery": sum_delivery,
                            "sum_produce": sum_produce,
                            "sum_stock": sum_stock
                        }])
                        new_meta.to_pickle(META_FILE)
                        df_filtered.to_pickle(DATA_FILE)

                        # --- 階段 3：🛡️ 異常檢測強化細節 ---
                        status_text.caption("⏳ [50%] 正在執行異常料號檢測 (自動對齊 Reference)...")
                        progress_bar.progress(50)
                        
                        check_df = pd.DataFrame({
                            "Item": raw_df.iloc[:, col_item],               
                            "料號 (Mfr_PN)": raw_df.iloc[:, col_mfr_pn],   
                            "Part Number": raw_df.iloc[:, col_part_num],
                            "Reference (零件位置)": raw_df.iloc[:, col_ref],  
                            "Part Name": raw_df.iloc[:, col_part_name]      
                        })
                        
                        check_df = check_df.dropna(subset=["Item"])
                        check_df["Item"] = check_df["Item"].astype(str).str.strip()
                        check_df = check_df[check_df["Item"].str.match(r'^\d+$', na=False)]
                        
                        check_df["料號 (Mfr_PN)"] = check_df["料號 (Mfr_PN)"].astype(str).str.strip()
                        invalid_keywords = ["", "NAN", "NONE", "N/A", "-", "NA"]
                        valid_mask = ~check_df["料號 (Mfr_PN)"].str.upper().isin(invalid_keywords)
                        valid_items = check_df[valid_mask].copy()
                        
                        dup_mask = valid_items.duplicated(subset=["料號 (Mfr_PN)"], keep=False)
                        anomaly_df = valid_items[dup_mask].copy()
                        
                        anomaly_df = anomaly_df.sort_values(by=["料號 (Mfr_PN)", "Item"])
                        anomaly_df = anomaly_df.fillna("").astype(str).replace({"nan": "", "None": ""})
                        anomaly_df.to_pickle(ANOMALY_FILE)

                    # --- 階段 4：處理專案進度 ---
                    prog_sheet_name = next((sn for sn in xls.sheet_names if "進度" in sn), None)
                    if prog_sheet_name:
                        status_text.caption("⏳ [80%] 正在分析排程進度表...")
                        progress_bar.progress(80)
                        raw_prog = pd.read_excel(xls, sheet_name=prog_sheet_name, header=None)
                        
                        prog_data = []
                        target_tasks = [
                            ("PCB", "到料", "PCB 到料完成時間"),
                            ("SMT", "齊料", "SMT料 齊料時間"),
                            ("DIP", "齊料", "DIP料 齊料時間"),
                            ("SMT", "上線", "SMT 安排上線時間"),
                            ("DIP", "上線", "DIP 安排上線時間"),
                            ("交貨", "", "交貨時間")
                        ]
                        
                        found_tasks = {}
                        for r in range(len(raw_prog)):
                            row_vals = raw_prog.iloc[r].values
                            row_str = "".join([str(x).upper() for x in row_vals]).replace(" ", "")
                            
                            for kw1, kw2, display_name in target_tasks:
                                if display_name in found_tasks: continue
                                if kw1.upper() in row_str and kw2.upper() in row_str:
                                    dates = []
                                    for v in row_vals:
                                        if isinstance(v, datetime): dates.append(v.strftime('%Y-%m-%d'))
                                        elif isinstance(v, str) and re.search(r'\d{4}[-/]\d{1,2}', v):
                                            match = re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', v)
                                            if match: dates.append(match.group())
                                    
                                    start, end = "-", "-"
                                    if len(dates) >= 2: start, end = dates[0], dates[1]
                                    elif len(dates) == 1: start = end = dates[0]
                                    
                                    found_tasks[display_name] = {"階段任務": display_name, "開始時間": start, "結束時間": end}
                        
                        for _, _, name in target_tasks:
                            prog_data.append(found_tasks.get(name, {"階段任務": name, "開始時間": "-", "結束時間": "-"}))
                        
                        pd.DataFrame(prog_data).to_pickle(PROG_FILE)

                    progress_bar.progress(100)
                    st.session_state['processed_file_id'] = current_file_id
                    status_text.empty()
                    st.sidebar.success("🎉 全數據同步成功！")
                    st.rerun()
                    
                except Exception as e:
                    st.sidebar.error(f"解析失敗: {e}")
            else:
                st.sidebar.success(f"✅ 檔案 {uploaded_file.name} 已完成同步。")

    st.sidebar.markdown("---")
    st.sidebar.caption("🤖 系統版本：`V6.9-30Set` (優化純淨面板)")
    st.sidebar.caption("⚙️ 引擎：Streamlit x Python")

# --- 主畫面顯示區域 ---
if os.path.exists(DATA_FILE) and os.path.exists(META_FILE):
    meta_df = pd.read_pickle(META_FILE)
    display_df = pd.read_pickle(DATA_FILE)
    
    st.info(f"📌 **資料版本 (檔案名稱)：** {meta_df.loc[0, 'version']}")
    
    c1, c2, c3 = st.columns(3)
    c1.metric("📦 總交貨數量", f"{int(meta_df.loc[0, 'sum_delivery']):,}")
    c2.metric("🏭 總生產數量", f"{int(meta_df.loc[0, 'sum_produce']):,}")
    c3.metric("💾 總備貨數量", f"{int(meta_df.loc[0, 'sum_stock']):,}")
    
    st.markdown("---")
    
    if os.path.exists(PROG_FILE):
        st.markdown("### 📅 專案排程進度")
        prog_df = pd.read_pickle(PROG_FILE)
        st.dataframe(prog_df, use_container_width=True, hide_index=True)
        st.markdown("---")
        
    if os.path.exists(ANOMALY_FILE):
        st.markdown("### 🛡️ 異常檢測 (重複料號分析)")
        anomaly_df = pd.read_pickle(ANOMALY_FILE)
        if len(anomaly_df) > 0:
            st.warning(f"⚠️ 發現 {len(anomaly_df)} 筆重複料號！請對照表格內的 Reference 欄位確認定義是否正確。")
            st.dataframe(anomaly_df, use_container_width=True, hide_index=True)
        else:
            st.success("✅ 系統檢測完畢，已自動對齊標頭並濾除無效空值，目前無重複定義之異常料號。")
        st.markdown("---")
    
    # BOM 表顯示 (已移除非必要提示，簡化排版)
    st.markdown("### 📋 BOM 物料清單狀態")
    st.caption(f"⏱️ 更新時間：{meta_df.loc[0, 'update_time']} ｜ 📊 總計：{len(display_df)} 筆")
    
    st.dataframe(display_df, width='stretch', hide_index=True, height=500)
else:
    st.warning("⏳ 系統資料尚未建立。請管理員上傳包含「成本-30」與「進度」Sheet 的檔案。")