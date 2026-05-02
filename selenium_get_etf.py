import time
import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import re

def clean_numeric(val):
    if val is None or val == "" or val == "-":
        return None
    try:
        s = str(val).replace(',', '').replace('%', '').replace('+', '').replace(' ', '')
        return float(s)
    except:
        return val

def get_browser():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def scrape_page(driver, url, scroll_times=10):
    print(f"正在開啟網頁: {url}")
    driver.get(url)
    time.sleep(5)
    for i in range(scroll_times):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.2)
    return driver.page_source

def parse_rows(html_content, mode="search"):
    name_symbol_pattern = re.compile(r'Ell\">(.*?)</div>.*?Ell\">(.*?)</span>', re.DOTALL)
    matches = list(name_symbol_pattern.finditer(html_content))
    
    data_list = []
    for match in matches:
        name = match.group(1).strip()
        symbol = match.group(2).strip()
        if not re.match(r'^[A-Z0-9]+\.TW', symbol): continue
        if any(k in name for k in ['自選股', '登入', '瀏覽']): continue
        
        start_pos = match.end()
        next_match_pos = html_content.find('Ell">', start_pos)
        end_pos = next_match_pos if next_match_pos != -1 else start_pos + 2000
        
        row_content = html_content[start_pos:end_pos]
        div_pattern = re.compile(r'<div[^>]*Miw\(\d+px\)[^>]*>(.*?)</div>', re.DOTALL)
        divs = div_pattern.findall(row_content)
        cols = [re.sub(r'<[^>]*>', '', d).strip() for d in divs]

        if mode == "search" and len(cols) >= 14:
            data_list.append({
                '代號': symbol, '股票名稱': name,
                '股價': clean_numeric(cols[0]), '漲跌幅(%)': clean_numeric(cols[1]),
                '成交量(張)': clean_numeric(cols[2]), '殖利率(%)': clean_numeric(cols[3]),
                '管理費(%)': clean_numeric(cols[4]), '資產規模(億)': clean_numeric(cols[5]),
                '成立年數': clean_numeric(cols[6]), '年初至今(%)': clean_numeric(cols[7]),
                '3個月績效(%)': clean_numeric(cols[8]), '六個月績效(%)': clean_numeric(cols[9]),
                '一年績效(%)': clean_numeric(cols[10]), '二年(年化%)': clean_numeric(cols[11]),
                '三年(年化%)': clean_numeric(cols[12]), '五年(年化%)': clean_numeric(cols[13]),
                '十年(年化%)': clean_numeric(cols[14]) if len(cols) > 14 else None
            })
        elif mode == "performance" and len(cols) >= 4:
            data_list.append({
                '代號': symbol,
                '1週績效(%)': clean_numeric(cols[3]) if len(cols) > 3 else None,
                '1個月績效(%)': clean_numeric(cols[4]) if len(cols) > 4 else None
            })
    return data_list

def run_master_scraper():
    print("--- Yahoo ETF 大數據合體工具 (300+筆 + 1週/1月績效) ---")
    driver = get_browser()
    try:
        search_html = scrape_page(driver, "https://tw.stock.yahoo.com/tw-etf/search", 15)
        search_data = parse_rows(search_html, "search")
        df_all = pd.DataFrame(search_data).drop_duplicates(subset=['代號'])
        print(f"✅ 已取得主列表: {len(df_all)} 筆詳細資料。")

        perf_html = scrape_page(driver, "https://tw.stock.yahoo.com/tw-etf/performance", 8)
        perf_data = parse_rows(perf_html, "performance")
        df_perf = pd.DataFrame(perf_data).drop_duplicates(subset=['代號'])
        print(f"✅ 已取得短期績效: {len(df_perf)} 筆 (1週/1月)。")

        if df_all.empty: return

        final_df = pd.merge(df_all, df_perf, on='代號', how='left')

        # --- 自動排序：先按 1週績效 降序，再按 年初至今 降序 ---
        # 這樣績效好的就會在最上面
        sort_cols = []
        if '1週績效(%)' in final_df.columns: sort_cols.append('1週績效(%)')
        if '年初至今(%)' in final_df.columns: sort_cols.append('年初至今(%)')
        final_df = final_df.sort_values(by=sort_cols, ascending=False)

        # 整理欄位順序
        new_order = ['股票名稱', '代號', '股價', '漲跌幅(%)', '1週績效(%)', '1個月績效(%)', '成交量(張)', '殖利率(%)', '管理費(%)', '資產規模(億)', '成立年數', '年初至今(%)', '3個月績效(%)', '六個月績效(%)', '一年績效(%)']
        cols = final_df.columns.tolist()
        final_df = final_df[new_order + [c for c in cols if c not in new_order]]

        # --- 儲存並自動添加 Excel 篩選按鈕 ---
        filename = f"Yahoo_ETF_大合體_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        
        # 使用 ExcelWriter 來加入進階格式
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='ETF績效排行')
            # 獲取工作表對象
            worksheet = writer.sheets['ETF績效排行']
            # 自動設定篩選範圍 (從 A1 到最後一個儲存格)
            full_range = f"A1:{chr(64 + len(final_df.columns))}{len(final_df) + 1}"
            worksheet.auto_filter.ref = full_range
            # 自動調整欄寬
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except: pass
                worksheet.column_dimensions[column_letter].width = max_length + 2

        print(f"✅ 任務圓滿完成！")
        print(f"🎉 最終產出檔案：{filename} (共 {len(final_df)} 筆，已自動排序)")

    except Exception as e:
        print(f"❌ 發生錯誤: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    run_master_scraper()
