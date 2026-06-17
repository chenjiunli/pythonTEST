import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
import zoneinfo

# 🌟 方案 A：維持側邊欄預設收合
st.set_page_config(
    page_title="SAA TCS Controller V2.0 專案進度查詢", 
    layout="wide",
    initial_sidebar_state="collapsed" 
)

# 🌟 隱藏所有不必要的元件（包含右上角 GitHub 貓咪圖示、表格工具列、側邊欄標題）
st.markdown("""
    <style>
    /* 🛠️ 取消右上角的 GitHub 連結圖示與主選單 */
    #MainMenu, footer, header {
        visibility: hidden !important;
        height: 0 !important;
    }
    /* 隱藏表格右上角的所有控制按鈕 */
    [data-testid="stDataFrameToolbar"] {
        display: none !important;
    }
    /* 隱藏側邊欄頂部管理標題區塊 */
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
        display: none !important;
    }
    /* 移除側邊欄頂部多餘的內距 */
    [data-testid="stSidebarUserContent"] {
        padding-top: 20px !important;
    }
    /* 禁止使用者反白選取網頁上的文字 */
    body {
        -webkit-user-select: none;
        -moz-user-select: none;
        -ms-user-select: none;
        user-select: none;
    }
    </style>
""", unsafe_allow_html=True)

# 變更專案大標題
st.title("🛠️ SAA TCS Controller V2.0 專案進度即時查詢面板")
st.write("🤝 本面板資料已自動同步，並依 Item 進行排序。您可於表格內利用滑鼠滾輪或手勢滑動瀏覽。")

# 設定雲端伺服器儲存最新資料與版本紀錄的檔案名稱
DATA_FILE = "latest_bom_data.pkl"
META_FILE = "bom_meta_info.pkl"  

# 管理員更新區域
with st.sidebar:
    password = st.text_input("管理員密碼：", type="password")
    
    if password == "1234":
        st.success("密碼正確！")
        uploaded_file = st.file_uploader("請上傳最新的 BOM Excel 檔案：", type=["xlsx"])
        
        if uploaded_file is not None:
            try:
                # 1. 讀取「成本-15」分頁
                # 為了抓取前三行的數量資訊，我們多讀取原始表頭上方
                raw_df = pd.read_excel(uploaded_file, sheet_name="成本-15", header=None, nrows=4)
                
                # 🎯 動態提取前三行的交貨、生產、備貨數量（假設這三個中文字在 A 欄或 B 欄附近）
                # 為了防呆，我們在全表前 4 行搜尋關鍵字並抓取其右方/下方的數字
                qty_info = {"交貨數量": "未知", "生產數量": "未知", "備貨數量": "未知"}
                for i in range(min(4, len(raw_df))):
                    row_str = " ".join([str(val) for val in raw_df.iloc[i].values])
                    for key in qty_info.keys():
                        if key in row_str:
                            # 嘗試撈取該行中非中文字的數字
                            for val in raw_df.iloc[i].values:
                                if isinstance(val, (int, float)) and not pd.isna(val):
                                    qty_info[key] = int(val)
                                    break

                # 2. 正式讀取主要資料，第 4 列為標頭 (header=3)
                df = pd.read_excel(uploaded_file, sheet_name="成本-15", header=3)
                df.columns = [str(c).strip() for c in df.columns]
                
                # 🎯 增加缺料欄位 Q
                target_cols = ["Item", "Manufacturer_P/N", "Manufacture_Name", "缺料", "交期"]
                valid_cols = [c for c in target_cols if c in df.columns]
                
                df_filtered = df[valid_cols].dropna(subset=["Item"])
                
                # 依 Item 數字由小到大排序
                df_filtered['Item_num'] = pd.to_numeric(df_filtered['Item'], errors='coerce')
                df_filtered = df_filtered.sort_values(by=['Item_num', 'Item'], ascending=True).drop(columns=['Item_num'])
                
                # 修正時區
                tz_taipei = zoneinfo.ZoneInfo("Asia/Taipei")
                current_time = datetime.now(tz_taipei).strftime("%Y-%m-%d %H:%M")
                
                # 擷取上傳的 Excel 檔案名稱作為資料版本
                file_version_name = uploaded_file.name
                
                # 把新的時間、檔案名稱版本、以及動態數量存進 Meta 檔案
                new_meta = pd.DataFrame([{
                    "version": file_version_name, 
                    "update_time": current_time,
                    "deliv_qty": qty_info["交貨數量"],
                    "prod_qty": qty_info["生產數量"],
                    "stock_qty": qty_info["備貨數量"]
                }])
                new_meta.to_pickle(META_FILE)
                
                # 永久儲存資料主體到伺服器空間
                df_filtered.to_pickle(DATA_FILE)
                st.sidebar.success(f"🎉 資料同步成功！\n版本：{file_version_name}")
            except Exception as e:
                st.sidebar.error(f"解析失敗，請確認檔案內有「成本-15」分頁。錯誤: {e}")

    # 左側邊欄最下方標註目前的軟體版本
    st.sidebar.markdown("---")
    st.sidebar.caption("🤖 系統軟體版本：`V2.4` (專案客製版)")
    st.sidebar.caption("⚙️ 核心引擎：Streamlit x Python 3.14")

# --- 主畫面顯示區域（客戶看到的畫面） ---
if os.path.exists(DATA_FILE):
    # 讀取資料
    display_df = pd.read_pickle(DATA_FILE)
    
    # 重新對齊欄位名稱（包含新的缺料欄位）
    display_df.columns = ["Item", "Manufacturer_P/N", "Manufacture_Name", "缺料", "交期"]
    
    # 預設歷史與數量紀錄
    version_label = "未命名版本"
    time_label = "未知"
    d_q, p_q, s_q = "隨檔案更動", "隨檔案更動", "隨檔案更動"
    
    if os.path.exists(META_FILE):
        meta_df = pd.read_pickle(META_FILE)
        version_label = meta_df.loc[0, 'version']
        time_label = meta_df.loc[0, 'update_time']
        # 讀取動態數量
        if "deliv_qty" in meta_df.columns:
            d_q = meta_df.loc[0, 'deliv_qty']
            p_q = meta_df.loc[0, 'prod_qty']
            s_q = meta_df.loc[0, 'stock_qty']
    
    # 🌟 資訊看板區優化：加入動態數量顯示（交貨、生產、備貨數量）
    st.markdown(f"""
    <div style="background-color: #f1f5f9; padding: 15px; border-radius: 8px; border-left: 5px solid #0284c7; margin-bottom: 10px;">
        <p style="margin: 0; font-size: 14px; color: #64748b; font-weight: bold;">📌 資料版本 (BOM 檔名)：</p>
        <p style="margin: 2px 0 12px 0; font-size: 16px; color: #0f172a; word-break: break-all; font-weight: 500;">{version_label}</p>
        
        <div style="display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 12px; background: #ffffff; padding: 10px; border-radius: 6px; box-shadow: inset 0 1px 2px rgba(0,0,0,0.05);">
            <div style="font-size: 14px; color: #1e293b;"><span style="color: #475569; font-weight: bold;">📦 交貨數量：</span><span style="color: #0284c7; font-weight: bold; font-size: 15px;">{d_q}</span></div>
            <div style="font-size: 14px; color: #1e293b;"><span style="color: #475569; font-weight: bold;">🏭 生產數量：</span><span style="color: #0284c7; font-weight: bold; font-size: 15px;">{p_q}</span></div>
            <div style="font-size: 14px; color: #1e293b;"><span style="color: #475569; font-weight: bold;">🪵 備貨數量：</span><span style="color: #0284c7; font-weight: bold; font-size: 15px;">{s_q}</span></div>
        </div>

        <div style="display: flex; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
            <div>
                <span style="font-size: 13px; color: #64748b; font-weight: bold;">⏱️ 更新時間：</span>
                <span style="font-size: 13px; color: #0f172a; font-weight: bold;">{time_label}</span>
            </div>
            <div>
                <span style="font-size: 13px; color: #64748b; font-weight: bold;">📊 總資料筆數：</span>
                <span style="font-size: 13px; color: #0f172a; font-weight: bold;">{len(display_df)} 筆</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 🌟 新增：缺料欄位邏輯備註說明字樣
    st.caption("💡 備註：「缺料」欄位之數據係以【備料數量】扣除【庫存量】自動計算得出。")
    
    # 展示乾淨大表格
    st.dataframe(
        display_df, 
        use_container_width=True, 
        hide_index=True,
        height=520 
    )
else:
    st.warning("⏳ 內部管理員尚未上傳初始化 BOM 資料，請聯絡管理員更新。")