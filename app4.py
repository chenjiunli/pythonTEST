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

# 數字安全清洗器：處理千分位逗號與不規則字元
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
        uploaded_file = st.file_uploader("請上傳最新的 BOM Excel 檔案：", type=["xlsx"])
        
        if uploaded_file is not None:
            try:
                # 1. 讀取 Excel (最原始狀態，不解析標頭)
                raw_df = pd.read_excel(uploaded_file, sheet_name="成本-15", header=None)
                
                # --- 🐛 除錯模式透視鏡 ---
                st.sidebar.markdown("---")
                st.sidebar.write("### 🐛 除錯模式：看看 Pandas 抓到什麼？")
                st.sidebar.code(f"N1 (0, 13) 值：{raw_df.iloc[0, 13]}")
                st.sidebar.code(f"N2 (1, 13) 值：{raw_df.iloc[1, 13]}")
                st.sidebar.code(f"N3 (2, 13) 值：{raw_df.iloc[2, 13]}")
                st.sidebar.code(f"Q欄前五筆：\n{raw_df.iloc[0:5, 16].tolist()}")
                st.sidebar.markdown("---")
                # -------------------------

                # 2. 🎯【絕對座標鎖定】：直接抓取 N1, N2, N3 的數值 (N欄 = Index 13)
                sum_delivery = clean_numeric_values(raw_df.iloc[0, 13])  # N1
                sum_produce = clean_numeric_values(raw_df.iloc[1, 13])   # N2
                sum_stock = clean_numeric_values(raw_df.iloc[2, 13])     # N3
                
                # 3. 🎯【絕對欄位鎖定】：A=0, C=2, K=10, Q=16, AB=27
                clean_df = pd.DataFrame({
                    "Item": raw_df.iloc[:, 0],               # A欄
                    "Manufacturer_P/N": raw_df.iloc[:, 2],   # C欄
                    "Manufacture_Name": raw_df.iloc[:, 10],  # K欄
                    "raw_shortage": raw_df.iloc[:, 16],      # Q欄 (缺貨數量)
                    "交期": raw_df.iloc[:, 27]               # AB欄
                })
                
                # 4. 行數過濾雜訊 (只保留 Item 欄位是純數字的列)
                df_data = clean_df.dropna(subset=["Item"]).copy()
                df_data["Item"] = df_data["Item"].astype(str).str.strip()
                df_data = df_data[df_data["Item"].str.match(r'^\d+$', na=False)]
                
                # 5. 清洗缺料數據 (小於等於 0 或文字皆轉換為 0)
                df_data["缺料"] = df_data["raw_shortage"].apply(clean_numeric_values)
                
                # 6. 構建最終輸出
                df_filtered = pd.DataFrame()
                df_filtered["Item"] = df_data["Item"]
                df_filtered["Manufacturer_P/N"] = df_data["Manufacturer_P/N"]
                df_filtered["Manufacture_Name"] = df_data["Manufacture_Name"]
                df_filtered["缺料"] = df_data["缺料"]
                df_filtered["交期"] = df_data["交期"]
                
                # --- 資料型態安全轉換 ---
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
                
                # 依 Item 自然數字排序
                df_filtered['Item_num'] = pd.to_numeric(df_filtered['Item'], errors='coerce')
                df_filtered = df_filtered.sort_values(by=['Item_num'], ascending=True).drop(columns=['Item_num'])
                
                # 取得時間
                tz_taibei = zoneinfo.ZoneInfo("Asia/Taipei")
                current_time = datetime.now(tz_taibei).strftime("%Y-%m-%d %H:%M")
                
                # 存檔至快取
                new_meta = pd.DataFrame([{
                    "version": uploaded_file.name, 
                    "update_time": current_time,
                    "sum_delivery": sum_delivery,
                    "sum_produce": sum_produce,
                    "sum_stock": sum_stock
                }])
                new_meta.to_pickle(META_FILE)
                df_filtered.to_pickle(DATA_FILE)
                
                st.sidebar.success(f"🎉 絕對座標同步成功！")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"解析失敗。錯誤: {e}")

    st.sidebar.markdown("---")
    st.sidebar.caption("🤖 系統軟體版本：`V5.5` (含除錯透視鏡)")
    st.sidebar.caption("⚙️ 核心引擎：Streamlit x Python")

# --- 主畫面顯示區域 ---
is_data_ready = False

if os.path.exists(DATA_FILE) and os.path.exists(META_FILE):
    try:
        display_df = pd.read_pickle(DATA_FILE)
        if len(display_df.columns) == 5:
            display_df.columns = ["Item", "Manufacturer_P/N", "Manufacture_Name", "缺料", "交期"]
            is_data_ready = True
    except:
        is_data_ready = False

if is_data_ready:
    meta_df = pd.read_pickle(META_FILE)
    version_label = meta_df.loc[0, 'version']
    time_label = meta_df.loc[0, 'update_time']
    v_delivery = int(meta_df.loc[0, 'sum_delivery'])
    v_produce = int(meta_df.loc[0, 'sum_produce'])
    v_stock = int(meta_df.loc[0, 'sum_stock'])
    
    st.info(f"📌 **資料版本 (BOM 檔名)：** {version_label}")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="📦 總交貨數量", value=f"{v_delivery:,}")
    with col2:
        st.metric(label="🏭 總生產數量", value=f"{v_produce:,}")
    with col3:
        st.metric(label="💾 總備貨數量", value=f"{v_stock:,}")
        
    col_left, col_right = st.columns([2, 1])
    with col_left:
        st.caption(f"⏱️ **更新時間：** {time_label} ｜ 📊 **總計：** {len(display_df)} 筆")
    with col_right:
        st.caption("💡 *註：缺料欄位已綁定 Q 欄，小於等於 0 之品項顯示為 0。*")
    
    st.markdown("---")
    
    st.dataframe(
        display_df, 
        width='stretch', 
        hide_index=True,
        height=550 
    )
else:
    st.warning("⏳ 系統資料尚未建立。請管理員展開左側選單，輸入密碼並上傳最新的 BOM Excel 檔案。")