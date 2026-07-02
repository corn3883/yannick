import streamlit as st
import pandas as pd
import os
import re
import time
import random
import io
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# 1. 設定 Chrome 參數
chrome_options = Options()
chrome_options.add_argument("--headless")          # 無頭模式（雲端必備）
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# 2. 針對 Streamlit 雲端環境的特殊設定
# 直接指定 Linux 系統內建的 chromium 與 chromedriver 位置
chrome_options.binary_location = "/usr/bin/chromium"
service = Service(executable_path="/usr/bin/chromedriver")

# 3. 啟動瀏覽器（將 service 和 options 帶入）
driver = webdriver.Chrome(service=service, options=chrome_options)

# --- 網頁頁面設定 ---
st.set_page_config(page_title="黑貓宅急便 批次查詢系統", layout="centered")

st.title("📦 黑貓宅急便 全自動批次查詢系統")
st.markdown("""
本系統會自動讀取您上傳的 Excel，並抓取黑貓官網**最新的物流狀態**。
1. Excel **A欄** 標題必須為 **`託運單號`**。
2. 查詢完畢後可直接下載 Excel 結果檔。
""")

# --- 核心爬蟲函數 ---
def query_tcat_logic(tracking_no, driver):
    if not tracking_no or pd.isna(tracking_no):
        return "單號為空", "無時間資料"
    
    tracking_no = str(tracking_no).strip().replace("'", "").replace('-', '').replace(' ', '')
    url = f"https://www.t-cat.com.tw/inquire/TraceDetail.aspx?BillID={tracking_no}"
    
    max_retries = 2
    for attempt in range(max_retries):
        try:
            driver.get(url)
            time.sleep(1.2)  
            
            page_text = driver.find_element(By.TAG_NAME, "body").text
            if "查無資料" in page_text or "無此單號" in page_text:
                return "查無此單號資料", "無時間資料"
            
            date_pattern = re.compile(r'\d{4}[/\-]\d{2}[/\-]\d{2}\s+\d{2}:\d{2}')
            
            # 定位表格列
            rows = driver.find_elements(By.TAG_NAME, "tr")
            valid_rows = []
            for row in rows:
                cells = row.find_elements(By.XPATH, "./td | ./th")
                cell_texts = [c.text.strip() for c in cells if c.text.strip()]
                if len(cell_texts) >= 2 and date_pattern.search(cell_texts[0]):
                    valid_rows.append(cell_texts)
            
            if valid_rows:
                # 抓取第一列：[0]時間, [1]狀態
                return valid_rows[0][1], valid_rows[0][0]
                
            return "狀態未知", "未能擷取網頁時間"
            
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2)
                continue
            else:
                return f"查詢失敗: {str(e)}", "查詢失敗"

# --- 網頁上傳區 ---
uploaded_file = st.file_uploader("📂 請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype={'託運單號': str})
        
        if '託運單號' not in df.columns:
            st.error("❌ 錯誤：Excel 內找不到 '託運單號' 欄位，請檢查標題名稱。")
        else:
            st.success(f"✅ 成功讀取 {len(df)} 筆單號！")
            
            if st.button("🚀 開始全自動批次查詢"):
                # 初始化進度條
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # 初始化瀏覽器
                options = webdriver.ChromeOptions()
                options.add_argument('--headless=new')
                options.add_argument('--disable-gpu')
                options.add_argument('--window-size=1920,1080')
                options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
                
                driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                
                results_status = []
                results_time = []
                
                start_time = time.time()
                
                for i, row in df.iterrows():
                    no = row['託運單號']
                    status_text.text(f"🔍 正在查詢第 {i+1} 筆 / 共 {len(df)} 筆 (單號: {no})")
                    
                    st_val, ti_val = query_tcat_logic(no, driver)
                    
                    results_status.append(st_val)
                    results_time.append(ti_val)
                    
                    # 更新網頁進度條
                    progress_bar.progress((i + 1) / len(df))
                    
                    # 隨機延遲預防封鎖
                    time.sleep(random.uniform(2.1, 3.8))
                
                driver.quit()
                
                # 更新 Dataframe
                df['目前狀態'] = results_status
                df['資料登入時間'] = results_time
                
                # 排序欄位
                cols = list(df.columns)
                fixed = ['託運單號', '目前狀態', '資料登入時間']
                for f in fixed: 
                    if f in cols: cols.remove(f)
                df = df[fixed + cols]
                
                end_time = time.time()
                st.balloons() # 查詢成功的小特效
                st.success(f"🎉 查詢完成！總耗時：{round(end_time - start_time, 1)} 秒")
                
                # --- 下載區 ---
                # 將結果轉為 Excel 緩存
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                processed_data = output.getvalue()
                
                st.markdown("### 💾 下載結果")
                st.download_button(
                    label="點擊下載查詢結果 Excel",
                    data=processed_data,
                    file_name=f"黑貓查詢結果_{time.strftime('%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 網頁預覽
                st.write("📋 結果預覽：", df.head(10))

    except Exception as e:
        st.error(f"發生意外錯誤: {e}")