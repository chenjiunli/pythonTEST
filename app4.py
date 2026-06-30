import streamlit as st
import pandas as pd
import io
import os
import re
from datetime import datetime
import zoneinfo

# 側邊欄強制保持展開
st.set_page_config(
    page_title="SAA TCS Controller V2.0 專案進度查詢", 
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

st.title("🛠️ SAA TCS Controller V2.0 專案進度即時查詢面板")
st.write("🤝 本面板資料已自動同步，並依 Item 進行排序。您可於表格內利用滑鼠滾輪或手勢滑動瀏覽。")

DATA_FILE = "latest_bom_data.pkl"
META_FILE = "bom_meta_info.pkl"  
PROG_FILE = "progress_data.pkl"  
ANOMALY_FILE = "anomaly_data.pkl"  # 🚀 新增：異常檢測快取檔案

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

with st.sidebar:
    st.markdown("### ⚙️ 內部資料管理")
    password = st.text_input("管理員密碼：", type="password")
    
    if password == "1234":
        st.success("密碼正確！")
        uploaded_file = st.file_uploader("請上傳最新的 Excel 檔案 (包含 BOM 與進度 Sheet)：", type=["xlsx"])
        
        if uploaded_file is not None:
            file_id = f"{uploaded_file.name}_{uploaded_file.size}"
            
            if st.session_state.get('processed_file') != file_id:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    # --- 階段 1 ---
                    status_text.caption("⏳ [10%] 正在載入並解析 Excel 結構...")
                    progress_bar.progress(10)
                    xls = pd.ExcelFile(uploaded_file)
                    
                    if "成本-15" in xls.sheet_names:
                        # --- 階段 2 ---
                        status_text.caption("⏳ [30%] 正在讀取 BOM 物料清單...")
                        progress_bar.progress(30)
                        raw_df = pd.read_excel(xls, sheet_name="成本-15", header=None)
                        
                        # --- 階段 3 ---
                        status_text.caption("⏳ [50%] 正在過濾並清洗缺料數據...")
                        progress_bar.progress(50)
                        
                        sum_delivery = clean_numeric_values(raw_df.iloc[0, 13])  # N1
                        sum_produce = clean_numeric_values(raw_df.iloc[1, 13])   # N2
                        sum_stock = clean_numeric_values(raw_df.iloc[2, 13])     # N3
                        
                        # 🎯 絕對欄位鎖定 (加入 Location 還有 Part Name 進行比對)
                        clean_df = pd.DataFrame({
                            "Item": raw_df.iloc[:, 0],               # A欄
                            "Manufacturer_P/N": raw_df.iloc[:, 2],   # C欄
                            "Location": raw_df.iloc[:, 5],           # 🚀 F欄 (Location)
                            "Part Name": raw_df.iloc[:, 7],          # 🚀 H欄 (Part Name)
                            "Manufacture_Name": raw_df.iloc[:, 10],  # K欄
                            "raw_shortage": raw_df.iloc[:, 16],      # Q欄 (缺貨數量)
                            "交期": raw_df.iloc[:, 27]               # AB欄
                        })
                        
                        df_data = clean_df.dropna(subset=["Item"]).copy()
                        df_data["Item"] = df_data["Item"].astype(str).str.strip()
                        df_data = df_data[df_data["Item"].str.match(r'^\d+$', na=False)]
                        df_data["缺料"] = df_data["raw_shortage"].apply(clean_numeric_values)
                        
                        df_data["Location"] = df_data["Location"].astype(str).str.strip().replace({"nan": "", "None": ""})
                        df_data["Part Name"] = df_data["Part Name"].astype(str).str.strip().replace({"nan": "", "None": ""})
                        
                        # --- 🚀 階段 3.5：異常檢測 (尋找重複料號) ---
                        status_text.caption("⏳ [60%] 正在比對料號定義與異常檢測...")
                        progress_bar.progress(60)
                        
                        valid_pn_mask = df_data["Manufacturer_P/N"].apply(lambda x: str(x).upper() not in ["", "NAN", "NONE"])
                        valid_pn_df = df_data[valid_pn_mask]
                        # 找出重複出現的料號 (Keep=False 代表所有重複的都會被標記出來)
                        dup_pns = valid_pn_df[valid_pn_df.duplicated(subset=["Manufacturer_P/N"], keep=False)]
                        
                        if not dup_pns.empty:
                            anomaly_df = dup_pns[["Item", "Manufacturer_P/N", "Part Name", "Location"]].copy()
                            anomaly_df['Item_num'] = pd.to_numeric(anomaly_df['Item'], errors='coerce')
                            # 依照料號排序，讓相同的料號排在一起方便肉眼比對
                            anomaly_df = anomaly_df.sort_values(by=["Manufacturer_P/N", "Item_num"]).drop(columns=['Item_num'])
                            anomaly_df.columns = ["項次 (Item)", "製造商料號 (P/N)", "料件名稱 (Part Name)", "位置 (Location)"]
                        else:
                            anomaly_df = pd.DataFrame()
                        
                        anomaly_df.to_pickle(ANOMALY_FILE)
                        
                        # --- 階段 4 ---
                        status_text.caption("⏳ [70%] 正在儲存 BOM 資料...")
                        progress_bar.progress(70)
                        
                        df_filtered = pd.DataFrame()
                        df_filtered["Item"] = df_data["Item"]
                        df_filtered["Manufacturer_P/N"] = df_data["Manufacturer_P/N"]
                        df_filtered["Manufacture_Name"] = df_data["Manufacture_Name"]
                        df_filtered["缺料"] = df_data["缺料"]
                        df_filtered["交期"] = df_data["交期"]
                        
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

                    prog_sheet_name = None
                    for sn in xls.sheet_names:
                        if "進度" in sn:  
                            prog_sheet_name = sn
                            break
                    
                    if prog_sheet_name:
                        # --- 階段 5 ---
                        status_text.caption("⏳ [80%] 正在讀取排程進度表...")
                        progress_bar.progress(80)
                        raw_prog = pd.read_excel(xls, sheet_name=prog_sheet_name, header=None)
                        
                        # --- 階段 6 ---
                        status_text.caption("⏳ [90%] 正在分析排程時間...")
                        progress_bar.progress(90)
                        
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
                            row_str_combined = "".join([str(x).strip().upper() for x in row_vals]).replace(" ", "")
                            
                            for kw1, kw2, display_name in target_tasks:
                                if display_name in found_tasks:
                                    continue
                                
                                if kw1.upper() in row_str_combined and kw2.upper() in row_str_combined:
                                    task_col_idx = -1
                                    for c_idx, val in enumerate(row_vals):
                                        v_str = str(val).strip().upper().replace(" ", "")
                                        if kw1.upper() in v_str:
                                            task_col_idx = c_idx
                                            break
                                    
                                    start_date, end_date = "", ""
                                    next_vals = []
                                    if task_col_idx != -1:
                                        for c_idx in range(task_col_idx + 1, len(row_vals)):
                                            v = row_vals[c_idx]
                                            if pd.notna(v) and str(v).strip() not in ['', 'nan', 'None', 'NaT']:
                                                next_vals.append(v)
                                    
                                    dates = []
                                    for v in next_vals:
                                        if isinstance(v, datetime):
                                            dates.append(v.strftime('%Y-%m-%d'))
                                        elif isinstance(v, str):
                                            d_str = v.strip()
                                            match1 = re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', d_str)
                                            match2 = re.search(r'\d{1,2}[-/]\d{1,2}', d_str)
                                            if match1: dates.append(match1.group())
                                            elif match2: dates.append(match2.group())
                                    
                                    if len(dates) >= 2:
                                        start_date, end_date = dates[0], dates[1]
                                    elif len(dates) == 1:
                                        start_date = dates[0]
                                        end_date = dates[0] 
                                    else:
                                        if len(next_vals) > 0: start_date = str(next_vals[0]).strip()
                                        if len(next_vals) > 1: end_date = str(next_vals[1]).strip()
                                    
                                    found_tasks[display_name] = {
                                        "階段任務": display_name,
                                        "開始時間": start_date,
                                        "結束時間": end_date
                                    }
                        
                        for _, _, display_name in target_tasks:
                            if display_name in found_tasks:
                                prog_data.append(found_tasks[display_name])
                            else:
                                prog_data.append({"階段任務": display_name, "開始時間": "-", "結束時間": "-"})
                        
                        pd.DataFrame(prog_data).to_pickle(PROG_FILE)

                    # --- 階段 7：完成 ---
                    progress_bar.progress(100)
                    status_text.empty()
                    progress_bar.empty()
                    
                    st.session_state['processed_file'] = file_id
                    st.sidebar.success(f"🎉 檔案 {uploaded_file.name} 數據同步成功！")
                    st.rerun()
                    
                except Exception as e:
                    progress_bar.empty()
                    status_text.empty()
                    st.sidebar.error(f"解析失敗。錯誤: {e}")
            else:
                st.sidebar.success(f"✅ 檔案 {uploaded_file.name} 已完成同步。")

    st.sidebar.markdown("---")
    st.sidebar.caption("🤖 系統軟體版本：`V6.4` (料號比對與防呆升級版)")
    st.sidebar.caption("⚙️ 核心引擎：Streamlit x Python")

# --- 主畫面顯示區域 ---
is_data_ready = False

if os.path.exists(DATA_FILE) and os.path.exists(META_FILE):
    try:
        display_df = pd.read_pickle(DATA_FILE)
        is_data_ready = True
    except:
        is_data_ready = False

if is_data_ready:
    meta_df = pd.read_pickle(META_FILE)
    
    version_label = meta_df.loc[0, 'version'] if 'version' in meta_df.columns else "未知版本"
    time_label = meta_df.loc[0, 'update_time'] if 'update_time' in meta_df.columns else "未知時間"
    v_delivery = int(meta_df.loc[0, 'sum_delivery']) if 'sum_delivery' in meta_df.columns else 0
    v_produce = int(meta_df.loc[0, 'sum_produce']) if 'sum_produce' in meta_df.columns else 0
    v_stock = int(meta_df.loc[0, 'sum_stock']) if 'sum_stock' in meta_df.columns else 0
    
    st.info(f"📌 **資料版本 (檔案名稱)：** {version_label}")
    
    # 頂部三大指標卡片
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="📦 總交貨數量", value=f"{v_delivery:,}")
    with col2:
        st.metric(label="🏭 總生產數量", value=f"{v_produce:,}")
    with col3:
        st.metric(label="💾 總備貨數量", value=f"{v_stock:,}")
        
    st.markdown("---")
    
    # === 🚀 新增：BOM 料件異常檢測區塊 ===
    if os.path.exists(ANOMALY_FILE):
        anomaly_df_display = pd.read_pickle(ANOMALY_FILE)
        st.markdown("### 🕵️ BOM 異常檢測 (重複料號衝突)")
        if not anomaly_df_display.empty:
            st.error("⚠️ **警告：發現以下料號在 BOM 表中重複出現！** 請比對 `Location` 與 `Part Name` 確認定義是否產生衝突：")
            st.dataframe(anomaly_df_display, use_container_width=True, hide_index=True)
        else:
            st.success("✅ **目前 BOM 表內定義完美**，未發現任何重複料號衝突！")
        st.markdown("---")

    # === 專案排程進度區塊 ===
    if os.path.exists(PROG_FILE):
        prog_df_display = pd.read_pickle(PROG_FILE)
        st.markdown("### 📅 專案排程進度")
        
        p_col1, p_col2 = st.columns([1.5, 1])
        with p_col1:
            st.dataframe(prog_df_display, use_container_width=True, hide_index=True)
        st.markdown("---")
    
    # === BOM 物料清單區塊 ===
    col_left, col_right = st.columns([2, 1])
    with col_left:
        st.markdown("### 📋 BOM 物料清單狀態")
        st.caption(f"⏱️ **更新時間：** {time_label} ｜ 📊 **總計：** {len(display_df)} 筆")
    with col_right:
        st.caption("💡 *註：缺料欄位已綁定 Q 欄，小於等於 0 之品項顯示為 0。*")
    
    st.dataframe(
        display_df, 
        width='stretch', 
        hide_index=True,
        height=550 
    )
else:
    st.warning("⏳ 系統資料尚未建立。請管理員展開左側選單，輸入密碼並上傳最新的 Excel 檔案。")