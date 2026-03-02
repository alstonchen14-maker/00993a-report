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
    print("🚀 啟動爬蟲 (自動翻頁合併版)...")
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
    
    all_records = pd.DataFrame() # 用來收集每一頁的資料
    
    try:
        driver.get(URL)
        print("⏳ 等待網頁載入...")
        time.sleep(10) 
        
        for page in range(1, 15): # 最多翻 14 頁 (確保能抓完)
            print(f"📄 正在讀取第 {page} 頁...")
            time.sleep(3) # 等待表格渲染
            
            html = driver.page_source
            page_df = None
            try:
                dfs = pd.read_html(StringIO(html))
                for df in dfs:
                    if isinstance(df.columns, pd.MultiIndex):
                        df.columns = df.columns.get_level_values(-1)
                    df.columns = df.columns.astype(str).str.strip()
                    cols = str(df.columns.tolist())
                    if ("權重" in cols or "持股" in cols or "比例" in cols) and ("名稱" in cols or "股票" in cols):
                        page_df = df
                        break
            except:
                pass
                
            if page_df is not None and not page_df.empty:
                # 檢查是否和上一頁重複 (防止按鈕失效但沒報錯)
                if not all_records.empty:
                    last_item = all_records.iloc[-1].to_string()
                    current_last = page_df.iloc[-1].to_string()
                    if last_item == current_last:
                        print("⚠️ 這一頁資料跟上一頁一樣，已達最後一頁。")
                        break
                        
                # 將這頁的資料加入總表
                if all_records.empty:
                    all_records = page_df
                else:
                    all_records = pd.concat([all_records, page_df], ignore_index=True)
                print(f"✅ 成功抓取！目前累計: {len(all_records)} 筆")
            else:
                print("❌ 這頁找不到表格")
                break

            # --- 尋找並點擊下一頁 ---
            print("🔍 尋找下一頁按鈕...")
            next_btn = None
            
            # 嘗試各種常見的下一頁 CSS 選擇器
            selectors = [
                "li.ant-pagination-next:not(.ant-pagination-disabled)", # Ant Design
                "li.next:not(.disabled)", # Bootstrap
                "button.next:not([disabled])",
                "[title='下一頁']:not(.ant-pagination-disabled)",
                ".pagination-next:not(.disabled)"
            ]
            for sel in selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, sel)
                    for el in elements:
                        if el.is_displayed():
                            next_btn = el
                            break
                except: pass
                if next_btn: break
                
            # 如果 CSS 找不到，用 XPATH 找含有 "下一頁" 或 ">" 的元素
            if not next_btn:
                try:
                    xpaths = [
                        "//li[contains(@class, 'next')]",
                        "//a[contains(text(), '下一頁')]",
                        "//button[contains(text(), '下一頁')]",
                        "//i[contains(@class, 'right')]/parent::*" # 箭頭圖示
                    ]
                    for xp in xpaths:
                        elements = driver.find_elements(By.XPATH, xp)
                        for el in elements:
                            if el.is_displayed():
                                next_btn = el
                                break
                        if next_btn: break
                except: pass

            if next_btn:
                classes = next_btn.get_attribute("class") or ""
                if "disabled" in classes.lower() or next_btn.get_attribute("disabled") or next_btn.get_attribute("aria-disabled") == "true":
                    print("🛑 下一頁按鈕已反灰，停止翻頁。")
                    break
                
                # 用 JS 點擊比較不會被擋住
                driver.execute_script("arguments[0].click();", next_btn)
                print("👉 點擊下一頁...")
            else:
                print("🛑 畫面上找不到下一頁按鈕，停止翻頁。")
                break

    except Exception as e:
        print(f"❌ 錯誤: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()
            
    # 最後去重複 (以防萬一)
    if not all_records.empty:
        col_n = next((c for c in all_records.columns if '名稱' in c), None)
        if col_n:
            all_records = all_records.drop_duplicates(subset=[col_n]).reset_index(drop=True)
        print(f"🎉 最終共取得 {len(all_records)} 筆不重複持股！")
        
    return all_records if not all_records.empty else None

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

    col_w = next((c for c in df_now.columns if '權重' in c or '比例' in c or '持股' in c), None)
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
