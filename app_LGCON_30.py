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

# 快取檔案名稱 (增加 _30 識別)
DATA_FILE = "latest_bom_data_30.pkl"
META_FILE = "bom_meta_info_30.pkl"  
PROG_FILE = "progress_data_30.pkl"  

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
            # 建立檔案唯一識別碼
            current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
            
            # 如果是新檔案才處理
            if st.session_state['processed_file_id'] != current_file_id:
                progress_bar = st.sidebar.progress(0)
                status_text = st.sidebar.empty()
                
                try:
                    # --- 階段 1：讀取檔案 ---
                    status_text.caption("⏳ [10%] 正在開啟 Excel 檔案結構...")
                    progress_bar.progress(10)
                    xls = pd.ExcelFile(uploaded_file)
                    
                    # --- 階段 2：處理 BOM (成本-30) ---
                    if "成本-30" in xls.sheet_names:
                        status_text.caption("⏳ [30%] 正在讀取 BOM 物料清單 (成本-30)...")
                        progress_bar.progress(30)
                        raw_df = pd.read_excel(xls, sheet_name="成本-30", header=None)
                        
                        # 🎯 絕對座標鎖定 (N1, N2, N3)
                        sum_delivery = clean_numeric_values(raw_df.iloc[0, 13])  # N1
                        sum_produce = clean_numeric_values(raw_df.iloc[1, 13])   # N2
                        sum_stock = clean_numeric_values(raw_df.iloc[2, 13])     # N3
                        
                        # 🎯 絕對欄位鎖定 (30套版: A=0, D=3, L=11, Q=16, AB=27)
                        clean_df = pd.DataFrame({
                            "Item": raw_df.iloc[:, 0],               
                            "Manufacturer_P/N": raw_df.iloc[:, 3],   
                            "Manufacture_Name": raw_df.iloc[:, 11],  
                            "raw_shortage": raw_df.iloc[:, 16],      
                            "交期": raw_df.iloc[:, 27]               
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

                    # --- 階段 3：處理專案進度 ---
                    prog_sheet_name = next((sn for sn in xls.sheet_names if "進度" in sn), None)
                    if prog_sheet_name:
                        status_text.caption("⏳ [70%] 正在分析排程進度表...")
                        progress_bar.progress(70)
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
                                    # 找日期
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

                    # --- 階段 4：完成 ---
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
    st.sidebar.caption("🤖 系統版本：`V6.5-30Set` (整合進度表)")
    st.sidebar.caption("⚙️ 引擎：Streamlit x Python")

# --- 主畫面顯示區域 ---
if os.path.exists(DATA_FILE) and os.path.exists(META_FILE):
    meta_df = pd.read_pickle(META_FILE)
    display_df = pd.read_pickle(DATA_FILE)
    
    st.info(f"📌 **資料版本 (檔案名稱)：** {meta_df.loc[0, 'version']}")
    
    # 指標卡片
    c1, c2, c3 = st.columns(3)
    c1.metric("📦 總交貨數量", f"{int(meta_df.loc[0, 'sum_delivery']):,}")
    c2.metric("🏭 總生產數量", f"{int(meta_df.loc[0, 'sum_produce']):,}")
    c3.metric("💾 總備貨數量", f"{int(meta_df.loc[0, 'sum_stock']):,}")
    
    st.markdown("---")
    
    # 排程進度顯示
    if os.path.exists(PROG_FILE):
        st.markdown("### 📅 專案排程進度")
        prog_df = pd.read_pickle(PROG_FILE)
        st.dataframe(prog_df, use_container_width=True, hide_index=True)
        st.markdown("---")
    
    # BOM 表顯示
    st.markdown("### 📋 BOM 物料清單狀態")
    l_col, r_col = st.columns([2, 1])
    l_col.caption(f"⏱️ 更新時間：{meta_df.loc[0, 'update_time']} ｜ 📊 總計：{len(display_df)} 筆")
    r_col.caption("💡 註：缺料綁定 Q 欄，≦ 0 顯示為 0。")
    
    st.dataframe(display_df, width='stretch', hide_index=True, height=500)
else:
    st.warning("⏳ 系統資料尚未建立。請管理員上傳包含「成本-30」與「進度」Sheet 的檔案。")