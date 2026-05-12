from pathlib import Path
import glob
import os
import re
import sys
import time

from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


URL = "https://www.ezmoney.com.tw/ETF/Fund/Info?fundCode=49YTW"
SNAPSHOT_DIR = Path("snapshots")
OUTPUT_XLSX = Path("每日異動紀錄.xlsx")


def fetch_html():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=options)
    try:
        driver.get(URL)
        time.sleep(3)

        tab = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="#asset"]'))
        )
        driver.execute_script("arguments[0].click();", tab)

        WebDriverWait(driver, 15).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, "#asset table td")) > 0
        )
        time.sleep(3)
        return driver.page_source
    finally:
        driver.quit()


def parse_snapshot(html):
    soup = BeautifulSoup(html, "html.parser")

    date_text = ""
    date_str = ""
    h5 = soup.find("h5")
    if h5:
        date_text = h5.get_text(strip=True)
        match = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", date_text)
        if match:
            date_str = f"{match.group(1)}{int(match.group(2)):02d}{int(match.group(3)):02d}"

    if not date_str:
        raise RuntimeError("找不到資料日期，請確認網站頁面格式是否改變。")

    table = None
    for candidate in soup.find_all("table", class_="table table-bordered middle table-striped mt-15"):
        th = candidate.find("th", attrs={"colspan": "4"})
        if th and "股票" in th.get_text(strip=True):
            table = candidate
            break

    if table is None:
        raise RuntimeError("找不到持股表格，請確認網站頁面格式是否改變。")

    rows = []
    for tr in table.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) == 4:
            rows.append([
                tds[0].get_text(strip=True),
                tds[1].get_text(strip=True),
                tds[2].get_text(strip=True),
                tds[3].get_text(strip=True),
            ])

    if not rows:
        raise RuntimeError("持股表格沒有資料。")

    df = pd.DataFrame(rows, columns=["股票代號", "股票名稱", "股數", "持股權重"])
    df["股數"] = df["股數"].str.replace(",", "", regex=False).astype(int)
    df["持股權重"] = df["持股權重"].str.replace("%", "", regex=False).astype(float)
    return date_text, date_str, df


def list_snapshot_files():
    dated_files = []
    for file_name in glob.glob(str(SNAPSHOT_DIR / "*.csv")):
        name = os.path.splitext(os.path.basename(file_name))[0]
        if re.fullmatch(r"\d{8}", name):
            dated_files.append((name, file_name))
    return sorted(dated_files, key=lambda item: item[0], reverse=True)


def build_diff(latest_date, latest_path, prev_date, prev_path):
    df_latest = pd.read_csv(latest_path, encoding="utf-8-sig", dtype={"股票代號": str})
    df_prev = pd.read_csv(prev_path, encoding="utf-8-sig", dtype={"股票代號": str})

    merged = pd.merge(
        df_latest,
        df_prev,
        on="股票代號",
        how="outer",
        suffixes=("_新", "_舊"),
    )

    records = []
    for _, row in merged.iterrows():
        code = row["股票代號"]
        name_new = row.get("股票名稱_新")
        name_old = row.get("股票名稱_舊")
        name = name_new if pd.notna(name_new) else name_old

        shares_new = row.get("股數_新")
        shares_old = row.get("股數_舊")
        weight_new = row.get("持股權重_新")
        weight_old = row.get("持股權重_舊")

        if pd.isna(shares_old) and pd.notna(shares_new):
            status = "新增"
            delta_shares = int(shares_new)
            delta_weight = round(float(weight_new), 4)
        elif pd.notna(shares_old) and pd.isna(shares_new):
            status = "刪除"
            delta_shares = -int(shares_old)
            delta_weight = round(-float(weight_old), 4)
        else:
            delta_shares = int(shares_new) - int(shares_old)
            delta_weight = round(float(weight_new) - float(weight_old), 4)
            if delta_shares > 0:
                status = "增加"
            elif delta_shares < 0:
                status = "減少"
            else:
                continue

        records.append([latest_date, code, name, status, delta_shares, delta_weight])

    df_diff = pd.DataFrame(
        records,
        columns=["日期", "股票代號", "股票名稱", "異動狀況", "異動股數", "異動持股權重"],
    )

    if df_diff.empty:
        return df_diff

    status_order = {"新增": 0, "增加": 1, "減少": 2, "刪除": 3}
    df_diff["_sort"] = df_diff["異動狀況"].map(status_order)
    return (
        df_diff.sort_values(by=["_sort", "股票代號"])
        .drop(columns=["_sort"])
        .reset_index(drop=True)
    )


def update_change_log(latest_date, df_diff):
    if OUTPUT_XLSX.exists():
        df_old = pd.read_excel(OUTPUT_XLSX, dtype={"日期": str, "股票代號": str})
        df_old = df_old[df_old["日期"] != latest_date]
        df_final = pd.concat([df_old, df_diff], ignore_index=True)
    else:
        df_final = df_diff.copy()

    df_final.to_excel(OUTPUT_XLSX, index=False)


def main():
    SNAPSHOT_DIR.mkdir(exist_ok=True)

    html = fetch_html()
    date_text, date_str, df = parse_snapshot(html)
    csv_path = SNAPSHOT_DIR / f"{date_str}.csv"

    print(date_text)
    if csv_path.exists():
        print(f"Latest data date {date_str} already exists. Skip update.")
        return 0

    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"Saved: {csv_path}")

    dated_files = list_snapshot_files()
    if len(dated_files) < 2:
        print("Only one snapshot exists. Skip diff for now.")
        return 0

    latest_date, latest_path = dated_files[0]
    prev_date, prev_path = dated_files[1]
    df_diff = build_diff(latest_date, latest_path, prev_date, prev_path)

    print(f"Compare: {prev_date} -> {latest_date}")
    print(df_diff.to_string(index=False))

    update_change_log(latest_date, df_diff)
    print(f"Updated: {OUTPUT_XLSX}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
