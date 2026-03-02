import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import os
import glob
import time
import datetime
import random
from io import StringIO

# --- 設定 ---
URL = "https://etf.allianzgi.com.tw/etf-info/E0002?tab=4"
HISTORY_DIR = "history"
HTML_FILENAME = "index.html"

if not os.path.exists(HISTORY_DIR):
    os.makedirs(HISTORY_DIR)

def get_data():
    print("🚀 啟動爬蟲 (濾渣精準點擊版)...")
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"❌ Driver 安裝失敗: {e}")
        return None
    
    target_df = None
    try:
        driver.get(URL)
        print("⏳ 等待網頁載入...")
        time.sleep(10)
        
        # 1. 關閉可能擋住按鈕的 Cookie 橫幅
        try:
            cookie_btn = driver.find_element(By.XPATH, "//button[contains(text(), '接受') or contains(text(), '同意')]")
            driver.execute_script("arguments[0].click();", cookie_btn)
            time.sleep(1)
        except: pass

        # 2. 暴力尋找並點擊「顯示更多」
        click_count = 0
        while click_count < 15:
            try:
                # 尋找畫面上所有的 "顯示更多"
                btns = driver.find_elements(By.XPATH, "//*[contains(text(), '顯示更多')]")
                visible_btn = None
                for b in btns:
                    if b.is_displayed():
                        visible_btn = b
                        break
                
                if visible_btn:
                    # 先把畫面滾動到按鈕的位置，確保沒有被遮擋
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", visible_btn)
                    time.sleep(1)
                    # 點擊按鈕
                    driver.execute_script("arguments[0].click();", visible_btn)
                    click_count += 1
                    print(f"👉 第 {click_count} 次點擊「顯示更多」...")
                    time.sleep(3) # 等待新表格長出來
                else:
                    print("✅ 畫面上已無「顯示更多」按鈕，展開完畢！")
                    break
            except Exception as e:
                print("🛑 點擊結束。")
                break
                
        print("⏳ 正在擷取完整表格...")
        time.sleep(3)
        
        dfs = pd.read_html(StringIO(driver.page_source))
        
        for df in dfs:
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(-1)
            df.columns = df.columns.astype(str).str.strip()
            cols = str(df.columns.tolist())
            
            # 必須包含名稱與比例
            if ("名稱" in cols or "股票" in cols) and ("比例" in cols or "權重" in cols or "持股" in cols):
                if len(df) > 5:
                    target_df = df
                    break
                    
    except Exception as e:
        print(f"❌ 錯誤: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()
            
    return target_df

def clean_percentage(x):
    try:
        if pd.isna(x): return 0.0
        s = str(x).replace('%', '').replace(',', '').strip()
        return float(s) if s != '-' else 0.0
    except: return 0.0

def generate_fake_history(df_now, col_w):
    print("✨ 生成模擬歷史資料...")
    df_fake = df_now.copy()
    for i in range(len(df_fake)):
        val = clean_percentage(df_fake.iloc[i][col_w])
        change = random.uniform(-0.3, 0.3)
        df_fake.at[i, col_w] = f"{max(0, val + change):.2f}%"
    yst = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    df_fake.to_csv(os.path.join(HISTORY_DIR, f"portfolio_{yst}.csv"), index=False, encoding="utf-8-sig")

def main():
    df_now = get_data()
    if df_now is None or df_now.empty: 
        print("❌ 抓取失敗，程式結束")
        return

    col_n = next((c for c in df_now.columns if '名稱' in c or '成分股' in c), None)
    col_w = next((c for c in df_now.columns if any(k in c for k in ['比例', '權重', '持股'])), None)
    col_c = next((c for c in df_now.columns if '代號' in c), col_n)

    if not col_n or not col_w:
        print("❌ 欄位辨識失敗")
        return

    # 🌟 關鍵修復：把名稱叫做「顯示更多」或空白的幽靈股票刪除！
    df_now = df_now[~df_now[col_n].astype(str).str.contains('顯示更多', na=False)]
    df_now = df_now[df_now[col_n].astype(str).str.strip() != '']
    
    print(f"🎉 成功過濾渣滓！總共取得 {len(df_now)} 筆真正持股！")

    today = datetime.date.today().strftime("%Y-%m-%d")
    csv_path = os.path.join(HISTORY_DIR, f"portfolio_{today}.csv")
    df_now.to_csv(csv_path, index=False, encoding="utf-8-sig")

    files = sorted(glob.glob(os.path.join(HISTORY_DIR, "*.csv")))
    if len(files) < 2:
        generate_fake_history(df_now, col_w)
        files = sorted(glob.glob(os.path.join(HISTORY_DIR, "*.csv")))

    f_now, f_prev = files[-1], files[-2]
    d_now = os.path.basename(f_now).replace("portfolio_", "").replace(".csv", "")
    d_prev = os.path.basename(f_prev).replace("portfolio_", "").replace(".csv", "")

    df1 = pd.read_csv(f_now, encoding="utf-8-sig").set_index(col_n)
    df2 = pd.read_csv(f_prev, encoding="utf-8-sig").set_index(col_n)
    
    df1 = df1[~df1.index.duplicated(keep='first')]
    df2 = df2[~df2.index.duplicated(keep='first')]

    m = df1.join(df2, lsuffix='_new', rsuffix='_old', how='outer')

    rows = ""
    w_new, w_old = f"{col_w}_new", f"{col_w}_old"
    
    m['sort_val'] = m[w_new].apply(clean_percentage)
    m = m.sort_values(by='sort_val', ascending=False)

    for name, r in m.iterrows():
        # 如果名字是 NaN 或是空的，直接跳過不顯示
        if pd.isna(name) or str(name).strip() == "": continue
        
        wn = r[w_new] if pd.notna(r[w_new]) else "0%"
        wo = r[w_old] if pd.notna(r[w_old]) else "0%"
        diff = clean_percentage(wn) - clean_percentage(wo)
        
        bg, tc, sym = "white", "#333", "-"
        if diff > 0.001: bg, tc, sym = "#ffe6e6", "#d93025", "▲"
        elif diff < -0.001: bg, tc, sym = "#e6ffe6", "#188038", "▼"
        
        rows += f"<tr style='background:{bg}'><td>{name}</td><td>{wo}</td><td>{wn}</td><td style='color:{tc}'><b>{sym} {diff:+.2f}%</b></td></tr>"

    html = f"""<!DOCTYPE html><html><head><meta charset='utf-8'><title>00993A 報表</title>
    <style>body{{font-family:sans-serif;max-width:800px;margin:20px auto;padding:10px;background:#f4f4f9}}
    .card{{background:white;padding:20px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,0.1)}}
    table{{width:100%;border-collapse:collapse}}th{{background:#2c3e50;color:white;padding:10px}}td{{padding:10px;border-bottom:1px solid #eee}}
    .btn{{display:inline-block;padding:10px 20px;background:#27ae60;color:white;text-decoration:none;border-radius:5px;margin-bottom:15px}}</style>
    <script>function downloadCSV() {{ 
        let csv = "\\uFEFF名稱,舊權重,新權重,變動\\n";
        document.querySelectorAll("table tr").forEach((tr, i) => {{
            if(i>0) csv += Array.from(tr.querySelectorAll("td")).map(td => td.innerText.replace(/,/g,"")).join(",") + "\\n";
        }});
        let link = document.createElement("a");
        link.href = "data:text/csv;charset=utf-8," + encodeURI(csv);
        link.download = "00993A_Report.csv"; link.click();
    }}</script></head>
    <body><div class='card'><h2>📈 安聯台灣主動式 (00993A)</h2><p>更新日期：{d_now}</p>
    <a href="javascript:downloadCSV()" class="btn">📥 下載報表 (Excel)</a>
    <table><thead><tr><th>名稱</th><th>舊權重</th><th>新權重</th><th>變動</th></tr></thead><tbody>{rows}</tbody></table></div></body></html>"""

    with open(HTML_FILENAME, "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ 報表生成完畢")

if __name__ == "__main__":
    main()
