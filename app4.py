import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="BOM 專案進度查詢", layout="wide")

# 網頁外觀與標題
st.title("🛠️ 迅得 BOM 專案進度即時查詢面板")
st.write("🤝 本面板資料已自動同步，並依 Item 進行排序。")

# 設定雲端伺服器儲存最新資料的檔案名稱
DATA_FILE = "latest_bom_data.pkl"

# 管理員更新區域（隱藏在側邊欄，輸入密碼才能上傳更新檔案）
with st.sidebar:
    st.subheader("⚙️ 內部資料管理")
    password = st.text_input("管理員密碼：", type="password")
    
    if password == "1234": # 這裡可以改成您自己想設的密碼
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
                
                # 永久儲存到伺服器空間
                df_filtered.to_pickle(DATA_FILE)
                st.sidebar.success("🎉 資料已成功同步至網頁伺服器！")
            except Exception as e:
                st.sidebar.error(f"解析失敗，請確認檔案內有「成本-15」分頁。錯誤: {e}")

# --- 主畫面顯示區域（客戶看到的畫面） ---
if os.path.exists(DATA_FILE):
    # 讀取目前儲存在雲端最新的資料
    display_df = pd.read_pickle(DATA_FILE)
    
    # 重新對齊欄位名稱
    display_df.columns = ["Item", "Manufacturer_P/N", "Manufacture_Name", "交期"]
    
    st.info(f"📊 目前已同步最新排程進度 ｜ 總計：{len(display_df)} 筆資料")
    
    # 隱藏 Index 索引，展示乾淨大表格
    st.dataframe(display_df, use_container_width=True, hide_index=True)
else:
    st.warning("⏳ 內部管理員尚未上傳初始化 BOM 資料，請聯絡管理員更新。")