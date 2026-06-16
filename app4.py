import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
import zoneinfo  # 引入時區控制套件

st.set_page_config(page_title="SAA TCS Controller V2.0 專案進度查詢", layout="wide")

# 1. 變更專案大標題
st.title("🛠️ SAA TCS Controller V2.0 專案進度即時查詢面板")
st.write("🤝 本面板資料已自動同步，並依 Item 進行排序。您可於表格內利用滑鼠滾輪或手勢滑動瀏覽。")

# 設定雲端伺服器儲存最新資料與版本紀錄的檔案名稱
DATA_FILE = "latest_bom_data.pkl"
META_FILE = "bom_meta_info.pkl"  

# 管理員更新區域（隱藏在側邊欄，輸入密碼才能上傳更新檔案）
with st.sidebar:
    st.subheader("⚙️ 內部資料管理")
    password = st.text_input("管理員密碼：", type="password")
    
    if password == "1234":
        st.success("密碼正確！")
        uploaded_file = st.file_uploader("請上傳最新的 BOM Excel 檔案：", type=["xlsx"])
        
        if uploaded_file is not None:
            try:
                # 讀取 成本-15 分頁，第 4 列為標頭 (header=3)
                df = pd.read_excel(uploaded_file, sheet_name="成本-15", header=3)
                df.columns = [str(c).strip() for c in df.columns]
                
                # 精準撈取 A(Item), C(Manufacturer_P/N), K(Manufacture_Name), AB(交期)
                target_cols = ["Item", "Manufacturer_P/N", "Manufacture_Name", "交期"]
                valid_cols = [c for c in target_cols if c in df.columns]
                
                df_filtered = df[valid_cols].dropna(subset=["Item"])
                
                # 依 Item 數字由小到大排序
                df_filtered['Item_num'] = pd.to_numeric(df_filtered['Item'], errors='coerce')
                df_filtered = df_filtered.sort_values(by=['Item_num', 'Item'], ascending=True).drop(columns=['Item_num'])
                
                # 🌟 【修正時區 BUG】：強制指定抓取台灣台北時間 (UTC+8)
                tz_taipei = zoneinfo.ZoneInfo("Asia/Taipei")
                current_time = datetime.now(tz_taipei).strftime("%Y-%m-%d %H:%M")
                
                # 2. 擷取上傳的 Excel 檔案名稱作為資料版本
                file_version_name = uploaded_file.name
                
                # 把新的時間與檔案名稱版本存進 Meta 檔案
                new_meta = pd.DataFrame([{"version": file_version_name, "update_time": current_time}])
                new_meta.to_pickle(META_FILE)
                
                # 永久儲存資料主體到伺服器空間
                df_filtered.to_pickle(DATA_FILE)
                st.sidebar.success(f"🎉 資料同步成功！\n版本：{file_version_name}")
            except Exception as e:
                st.sidebar.error(f"解析失敗，請確認檔案內有「成本-15」分頁。錯誤: {e}")

    # 🌟 【新需求】：在左側邊欄最下方標註目前的軟體版本
    st.sidebar.markdown("---")
    st.sidebar.caption("🤖 系統軟體版本：`V2.1`")
    st.sidebar.caption("⚙️ 核心引擎：Streamlit x Python 3.14")

# --- 主畫面顯示區域（客戶看到的畫面） ---
if os.path.exists(DATA_FILE):
    # 讀取資料
    display_df = pd.read_pickle(DATA_FILE)
    display_df.columns = ["Item", "Manufacturer_P/N", "Manufacture_Name", "交期"]
    
    # 預設歷史紀錄
    version_label = "未命名版本"
    time_label = "未知"
    
    if os.path.exists(META_FILE):
        meta_df = pd.read_pickle(META_FILE)
        version_label = meta_df.loc[0, 'version']
        time_label = meta_df.loc[0, 'update_time']
    
    # 優化直式防跑版排版
    st.markdown(f"""
    <div style="background-color: #f1f5f9; padding: 15px; border-radius: 8px; border-left: 5px solid #0284c7; margin-bottom: 20px;">
        <p style="margin: 0; font-size: 14px; color: #64748b; font-weight: bold;">📌 資料版本 (BOM 檔名)：</p>
        <p style="margin: 2px 0 10px 0; font-size: 16px; color: #0f172a; word-break: break-all;">{version_label}</p>
        <div style="display: flex; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
            <div>
                <span style="font-size: 14px; color: #64748b; font-weight: bold;">⏱️ 更新時間：</span>
                <span style="font-size: 14px; color: #0f172a; font-weight: bold;">{time_label}</span>
            </div>
            <div>
                <span style="font-size: 14px; color: #64748b; font-weight: bold;">📊 總資料筆數：</span>
                <span style="font-size: 14px; color: #0f172a; font-weight: bold;">{len(display_df)} 筆</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 展示乾淨大表格
    st.dataframe(
        display_df, 
        use_container_width=True, 
        hide_index=True,
        height=550 
    )
else:
    st.warning("⏳ 內部管理員尚未上傳初始化 BOM 資料，請聯絡管理員更新。")