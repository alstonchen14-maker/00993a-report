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
    print("🚀 啟動爬蟲 (終極破解版)...")
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
        print("⏳ 等待網頁初始載入...")
        time.sleep(10) 
        
        # --- 🌟 終極破解：模擬真人連續向下滾動 ---
        print("🖱️ 開始模擬滾動以加載完整持股...")
        last_height = driver.execute_script("return document.body.scrollHeight")
        for i in range(10):  # 最多滾動 10 次
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)  # 每次滾動停 2 秒讓資料讀取
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
            print(f"  > 已滾動第 {i+1} 次")

        # 針對安聯網頁，有些表格需要點擊「載入更多」或切換顯示筆數
        # 我們直接用 JS 強制把所有隱藏的列顯示出來 (如果有的話)
        driver.execute_script("""
            let tables = document.querySelectorAll('table');
            tables.forEach(t => t.style.display = 'block');
        """)
        
        print("⏳ 滾動完成，正在提取表格...")
        time.sleep(3)

        try:
            # 抓取目前的 HTML
            dfs = pd.read_html(StringIO(driver.page_source))
            print(f"🔍 找到 {len(dfs)} 個表格")
        except:
            print("❌ 找不到表格")
            return None

        for df in dfs:
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(-1)
            df.columns = df.columns.astype(str).str.strip()
            cols = str(df.columns.tolist())
            
            # 增加關鍵字「比例」與「成分股」
            if any(k in cols for k in ["權重", "持股", "比例", "成分股"]) and any(k in cols for k in ["名稱", "股票"]):
                # 檢查抓到的筆數是否大於 10
                if len(df) > 5: # 只要不是空表就先收
                    target_df = df
                    print(f"✅ 成功抓取表格，目前筆數: {len(target_df)}")
                    # 如果筆數還是太少，我們繼續找下一個表格 (有些網頁會有好幾個表)
                    if len(target_df) > 10:
                        break
    except Exception as e:
        print(f"❌ 錯誤: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()
    return target_df

# ... (下方 clean_percentage, generate_fake_history, main 函數保持不變) ...

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
    df_fake.to_csv(os.path.join(HISTORY_DIR, f"portfolio_{yst}.csv"), index=False)

def main():
    df_now = get_data()
    if df_now is None: 
        print("❌ 抓取失敗，程式結束")
        return

    col_w = next((c for c in df_now.columns if any(k in c for k in ['權重', '比例', '持股'])), None)
    col_n = next((c for c in df_now.columns if '名稱' in c), None)
    col_c = next((c for c in df_now.columns if '代號' in c), col_n)

    if col_w and col_n:
        today = datetime.date.today().strftime("%Y-%m-%d")
        csv_path = os.path.join(HISTORY_DIR, f"portfolio_{today}.csv")
        df_now.to_csv(csv_path, index=False)
        print(f"✅ 今日資料已儲存: {csv_path}")

        files = sorted(glob.glob(os.path.join(HISTORY_DIR, "*.csv")))
        if len(files) < 2:
            generate_fake_history(df_now, col_w)
            files = sorted(glob.glob(os.path.join(HISTORY_DIR, "*.csv")))

        f_now, f_prev = files[-1], files[-2]
        d_now = os.path.basename(f_now).replace("portfolio_", "").replace(".csv", "")
        d_prev = os.path.basename(f_prev).replace("portfolio_", "").replace(".csv", "")

        df1 = pd.read_csv(f_now).drop_duplicates(subset=[col_c]).set_index(col_c)
        df2 = pd.read_csv(f_prev).drop_duplicates(subset=[col_c]).set_index(col_c)
        m = df1.join(df2, lsuffix='_new', rsuffix='_old', how='outer')

        rows = ""
        m['sort'] = m[f"{col_w}_new"].apply(clean_percentage)
        m = m.sort_values(by='sort', ascending=False)

        for i, r in m.iterrows():
            nm = r[f"{col_n}_new"] if pd.notna(r[f"{col_n}_new"]) else r[f"{col_n}_old"]
            wn = r[f"{col_w}_new"] if pd.notna(r[f"{col_w}_new"]) else "0%"
            wo = r[f"{col_w}_old"] if pd.notna(r[f"{col_w}_old"]) else "0%"
            diff = clean_percentage(wn) - clean_percentage(wo)
            bg, tc, sym = "white", "#333", "-"
            if diff > 0.001: bg, tc, sym = "#ffe6e6", "#d93025", "▲"
            elif diff < -0.001: bg, tc, sym = "#e6ffe6", "#188038", "▼"
            rows += f"<tr style='background:{bg}'><td>{nm}</td><td>{wo}</td><td>{wn}</td><td style='color:{tc}'><b>{sym} {diff:+.2f}%</b></td></tr>"

        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>00993A ETF 追蹤報表</title>
<style>
  body{{font-family:"Microsoft JhengHei",sans-serif;max-width:800px;margin:20px auto;padding:10px;background:#f4f4f9}}
  .card{{background:white;padding:20px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,0.1)}}
  table{{width:100%;border-collapse:collapse;margin-top:20px}}
  th{{background:#2c3e50;color:white;padding:10px;text-align:left}}
  td{{padding:10px;border-bottom:1px solid #eee}}
  .btn{{
    display:inline-block; padding:10px 20px; 
    background-color:#27ae60; color:white; 
    text-decoration:none; border-radius:5px; 
    font-weight:bold; cursor:pointer; border:none; margin-bottom:15px;
  }}
  .btn:hover{{background-color:#219150;}}
</style>
<script>
function downloadCSV() {{
    var csv = [];
    var rows = document.querySelectorAll("table tr");
    for (var i = 0; i < rows.length; i++) {{
        var row = [], cols = rows[i].querySelectorAll("td, th");
        for (var j = 0; j < cols.length; j++) 
            row.push(cols[j].innerText.replace(/,/g, ""));
        csv.push(row.join(","));
    }}
    var csvFile = new Blob(["\\uFEFF" + csv.join("\\n")], {{type: "text/csv"}});
    var downloadLink = document.createElement("a");
    downloadLink.download = "00993A_Report_{d_now}.csv";
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = "none";
    document.body.appendChild(downloadLink);
    downloadLink.click();
}}
</script>
</head>
<body>
<div class='card'>
  <h2>📈 安聯台灣主動式 (00993A)</h2>
  <p style='color:#666'>更新日期：{d_now} (比較對象: {d_prev})</p>
  <button onclick="downloadCSV()" class="btn">📥 下載報表 (Excel)</button>
  <table>
    <thead><tr><th>名稱</th><th>舊權重</th><th>新權重</th><th>變動</th></tr></thead>
    <tbody>{rows}</tbody>
  </table>
</div>
</body>
</html>"""

        with open(HTML_FILENAME, "w", encoding="utf-8") as f:
            f.write(html)
        print("✅ HTML 報表生成完畢")
    else:
        print("❌ 找不到符合的欄位名稱")

if __name__ == "__main__":
    main()
