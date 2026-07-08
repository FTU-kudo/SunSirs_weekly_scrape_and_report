"""
╔══════════════════════════════════════════════════════════════════╗
║   SunSirs – Incremental Scraper  (VS Code / GitHub Actions)      ║
║                                                                  ║
║  Chạy: python src/commodities/sunsirs_scraper.py                 ║
║                                                                  ║
║  Lần đầu  → scrape từ 2018-01-01 đến hôm nay (~45 phút)          ║
║  Lần sau  → chỉ scrape ngày mới kể từ lần lưu cuối               ║
║                                                                  ║
║  Output:                                                         ║
║    data/raw/commodities/sunsirs/sunsirs_raw.parquet              ║
║    data/processed/commodities/selected_commodity_prices.csv      ║
╚══════════════════════════════════════════════════════════════════╝
"""

# ──────────────────────────────────────────────────────────────────────────────
# Imports
# ──────────────────────────────────────────────────────────────────────────────
import os
import time
import random
import logging
from datetime import date, timedelta
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import pandas as pd
from tqdm import tqdm  # terminal-style (không dùng tqdm.notebook như Colab)

# ──────────────────────────────────────────────────────────────────────────────
# Đường dẫn — dùng config/paths.py nếu có, fallback về relative path
# ──────────────────────────────────────────────────────────────────────────────
try:
    from config.paths import RAW_SUNSIRS, PROCESSED_COMMOD
except ImportError:
    # Fallback khi chạy trực tiếp từ thư mục src/
    BASE_DIR      = Path(__file__).resolve().parent.parent.parent
    RAW_SUNSIRS   = BASE_DIR / "data" / "raw"       / "commodities" / "sunsirs"
    PROCESSED_COMMOD = BASE_DIR / "data" / "processed" / "commodities"

# Tạo thư mục nếu chưa tồn tại
RAW_SUNSIRS.mkdir(parents=True, exist_ok=True)
PROCESSED_COMMOD.mkdir(parents=True, exist_ok=True)

# Tên file output
RAW_PARQUET  = RAW_SUNSIRS   / "sunsirs_raw.parquet"          # thay thế .xlsx
SELECTED_CSV = PROCESSED_COMMOD / "selected_commodity_prices.csv"  # thay thế .xlsx

# ──────────────────────────────────────────────────────────────────────────────
# Cấu hình — chỉnh sửa tại đây
# ──────────────────────────────────────────────────────────────────────────────
SERIES_START = date(2018, 1, 1)    # ngày bắt đầu scrape lần đầu
TODAY        = date.today()

DELAY_MIN    = 1.5     # giây chờ giữa các request (tối thiểu)
DELAY_MAX    = 3.5     # giây chờ giữa các request (tối đa)
MAX_RETRIES  = 3
RETRY_WAIT   = 10      # giây chờ trước khi retry

# Hàng hoá cần lọc vào file selected
# Key = tên cần tìm (khớp một phần, không phân biệt hoa/thường)
# Value = tên cột trong file output
COMMODITIES_WANTED = {
    "Urea"              : "Urea",
    "Phosphorus yellow" : "Phosphorus yellow",
    "Phosphoric acid"   : "Phosphoric acid",
    "Hydrochloric acid" : "Hydrochloric acid",
    "Sulfuric acid"     : "Sulfuric acid",
}

# ──────────────────────────────────────────────────────────────────────────────
# Logging — hiển thị log trong terminal VS Code / GitHub Actions
# ──────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# HTTP Session
# ──────────────────────────────────────────────────────────────────────────────
BASE_URL = "https://sunsirs.com/uk/sdetail-day-{yyyy}-{mmdd}.html"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept"         : "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer"        : "https://sunsirs.com/uk/",
})

# ──────────────────────────────────────────────────────────────────────────────
# Helper: xây URL theo ngày
# ──────────────────────────────────────────────────────────────────────────────
def build_url(d: date) -> str:
    return BASE_URL.format(yyyy=d.strftime("%Y"), mmdd=d.strftime("%m%d"))


# ──────────────────────────────────────────────────────────────────────────────
# Helper: fetch trang HTML với retry
# ──────────────────────────────────────────────────────────────────────────────
def fetch_page(url: str) -> str | None:
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = SESSION.get(url, timeout=20)
            if r.status_code == 200:
                return r.text
            if r.status_code == 404:
                return None   # ngày lễ / cuối tuần → không có data
            log.warning("HTTP %s (lần %d): %s", r.status_code, attempt, url)
        except requests.RequestException as e:
            log.warning("Lỗi request (lần %d): %s", attempt, e)
        if attempt < MAX_RETRIES:
            time.sleep(RETRY_WAIT)
    return None


# ──────────────────────────────────────────────────────────────────────────────
# Helper: parse bảng giá từ HTML
# ──────────────────────────────────────────────────────────────────────────────
def parse_prices(html: str, d: date) -> list[dict]:
    soup  = BeautifulSoup(html, "lxml")
    table = (
        soup.find("table", class_=lambda c: c and "com"   in c.lower())
        or soup.find("table", class_=lambda c: c and "price" in c.lower())
        or soup.find("table")
    )
    if not table:
        return []

    rows = table.find_all("tr")
    headers, data_rows = [], []
    for row in rows:
        cells = row.find_all(["th", "td"])
        texts = [c.get_text(strip=True) for c in cells]
        if not texts:
            continue
        if not headers and row.find("th"):
            headers = texts
        else:
            data_rows.append(texts)
    if not headers and data_rows:
        headers = data_rows.pop(0)

    # Đổi tên cột giá cho nhất quán
    if len(headers) > 3:
        headers[2] = "Previous day price"
        headers[3] = "Current day price"

    records = []
    for row in data_rows:
        if not any(row):
            continue
        row = row[:len(headers)] + [""] * max(0, len(headers) - len(row))
        rec = {"date": d.isoformat()}
        for h, v in zip(headers, row):
            rec[h] = v
        records.append(rec)
    return records


# ──────────────────────────────────────────────────────────────────────────────
# Helper: tự động nhận diện cột tên hàng hoá
# ──────────────────────────────────────────────────────────────────────────────
def detect_name_col(df: pd.DataFrame) -> str | None:
    for kw in ["tên hàng", "hàng hóa", "commodity", "name"]:
        for col in df.columns:
            if kw.lower() in col.lower():
                return col
    return None


# ──────────────────────────────────────────────────────────────────────────────
# Parquet I/O — thay thế hoàn toàn Excel cho raw data
# ──────────────────────────────────────────────────────────────────────────────
def get_last_saved_date(parquet_path: Path) -> date | None:
    """Đọc parquet, trả về ngày gần nhất đã lưu."""
    if not parquet_path.exists():
        return None
    try:
        df = pd.read_parquet(parquet_path, columns=["date"])
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.dropna(subset=["date"])
        return df["date"].max().date() if not df.empty else None
    except Exception as e:
        log.warning("Không đọc được file hiện tại: %s", e)
        return None


def append_rows_to_parquet(new_records: list[dict], parquet_path: Path) -> pd.DataFrame:
    """
    Merge dữ liệu mới vào parquet hiện có (hoặc tạo mới).
    Tự động deduplicate theo (date, tên hàng hoá).
    """
    new_df = pd.DataFrame(new_records)
    new_df["date"] = pd.to_datetime(new_df["date"]).dt.date.astype(str)

    if parquet_path.exists():
        existing = pd.read_parquet(parquet_path)
        combined = pd.concat([existing, new_df], ignore_index=True)
    else:
        combined = new_df

    name_col = detect_name_col(combined)
    if name_col:
        combined = combined.drop_duplicates(subset=["date", name_col], keep="last")

    combined = combined.sort_values("date").reset_index(drop=True)
    combined.to_parquet(parquet_path, index=False, engine="pyarrow")
    return combined


# ──────────────────────────────────────────────────────────────────────────────
# Migration: chuyển xlsx cũ → parquet (chạy 1 lần nếu đã có xlsx)
# ──────────────────────────────────────────────────────────────────────────────
def migrate_xlsx_to_parquet():
    """
    Nếu đã có file .xlsx từ thời Colab mà chưa có .parquet,
    tự động convert sang parquet.
    """
    if RAW_PARQUET.exists():
        return  # đã migrate rồi, bỏ qua

    # Tìm file xlsx bất kỳ trong thư mục raw
    xlsx_files = sorted(RAW_SUNSIRS.glob("*.xlsx"))
    if not xlsx_files:
        return

    xlsx_path = xlsx_files[-1]  # lấy file mới nhất (theo tên)
    log.info("⏳ Tìm thấy file cũ: %s → đang convert sang parquet…", xlsx_path.name)

    try:
        df = pd.read_excel(xlsx_path, dtype=str, engine="openpyxl")
        df.to_parquet(RAW_PARQUET, index=False, engine="pyarrow")
        log.info("✅ Migration xong: %s → sunsirs_raw.parquet (%d rows)",
                 xlsx_path.name, len(df))
    except Exception as e:
        log.warning("❌ Migration thất bại: %s", e)


# ──────────────────────────────────────────────────────────────────────────────
# Build selected commodities → CSV (sep=';')
# ──────────────────────────────────────────────────────────────────────────────
def clean_price(series: pd.Series) -> pd.Series:
    return (
        series
        .str.replace(r"[^\d.\-]", "", regex=True)
        .replace("", pd.NA)
        .pipe(pd.to_numeric, errors="coerce")
    )


def build_selected_csv(raw_path: Path, out_path: Path):
    """
    Đọc parquet raw → lọc các hàng hoá trong COMMODITIES_WANTED
    → pivot theo ngày → lưu CSV (sep=';') cho Google Sheets / Excel.
    """
    if not raw_path.exists():
        log.warning("⚠  Raw parquet không tồn tại — bỏ qua bước export CSV.")
        return

    raw = pd.read_parquet(raw_path)
    raw.columns = raw.columns.str.strip()

    name_col  = detect_name_col(raw)
    price_col = next(
        (c for c in raw.columns
         if any(k in c.lower() for k in ["current day price", "price"])),
        None,
    )

    if not name_col or not price_col:
        log.warning("⚠  Không nhận diện được cột tên/giá. Các cột: %s", list(raw.columns))
        return

    # Fuzzy match tên hàng hoá
    available = raw[name_col].dropna().unique()
    matched   = {}
    for search, label in COMMODITIES_WANTED.items():
        hits = [n for n in available if search.lower() in n.lower()]
        for h in hits:
            matched[h] = label

    if not matched:
        log.warning("⚠  Không khớp được hàng hoá nào. Kiểm tra COMMODITIES_WANTED.")
        return

    log.info("🎯 Đã khớp: %s", list(matched.values()))

    # Lọc, làm sạch, pivot
    filtered = raw[raw[name_col].isin(matched)].copy()
    filtered["_price"] = clean_price(filtered[price_col])
    filtered["date"]   = pd.to_datetime(filtered["date"], errors="coerce")
    filtered["_label"] = filtered[name_col].map(matched)
    filtered = filtered.dropna(subset=["date", "_price"])

    pivoted = (
        filtered
        .pivot_table(index="date", columns="_label",
                     values="_price", aggfunc="mean")
        .sort_index()
        .reset_index()
    )
    pivoted.columns.name = None

    # Sắp xếp cột theo thứ tự COMMODITIES_WANTED
    ordered = ["date"] + [v for v in COMMODITIES_WANTED.values() if v in pivoted.columns]
    pivoted = pivoted[ordered]
    pivoted["date"] = pivoted["date"].dt.strftime("%Y-%m-%d")

    # Lưu CSV với sep=';' — tương thích Excel locale Việt Nam & Google Sheets
    pivoted.to_csv(out_path, sep=";", index=False, encoding="utf-8-sig")

    log.info("✅ CSV đã lưu → %s  (%d ngày × %d hàng hoá)",
             out_path.name, len(pivoted), len(pivoted.columns) - 1)


# ──────────────────────────────────────────────────────────────────────────────
# Helper: sinh dãy ngày
# ──────────────────────────────────────────────────────────────────────────────
def date_range(start: date, end: date):
    cur = start
    while cur <= end:
        yield cur
        cur += timedelta(days=1)


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────
def main():
    log.info("=" * 60)
    log.info("⏳ SunSirs Scraper — %s", TODAY)
    log.info("=" * 60)
    log.info("   Raw    → %s", RAW_PARQUET)
    log.info("   Output → %s", SELECTED_CSV)

    # Bước 0: migrate xlsx cũ (từ Colab) sang parquet nếu cần
    migrate_xlsx_to_parquet()

    # Bước 1: xác định khoảng ngày cần scrape
    last_saved = get_last_saved_date(RAW_PARQUET)

    if last_saved is None:
        scrape_from = SERIES_START
        log.info("📂 Chưa có dữ liệu → scrape toàn bộ: %s → %s", SERIES_START, TODAY)
    else:
        scrape_from = last_saved + timedelta(days=1)
        days_behind = (TODAY - last_saved).days
        log.info("📂 Dữ liệu hiện có đến: %s", last_saved)
        log.info("   Chỉ scrape ngày mới: %s → %s (%d ngày)",
                 scrape_from, TODAY, days_behind)

    if scrape_from > TODAY:
        log.info("✅ Đã up-to-date — không cần scrape thêm.")
        build_selected_csv(RAW_PARQUET, SELECTED_CSV)
        return

    # Bước 2: scrape
    dates_to_scrape  = list(date_range(scrape_from, TODAY))
    all_new_records  = []
    empty_days       = []
    failed_days      = []

    log.info("\n⏳ Đang scrape %d ngày…\n", len(dates_to_scrape))

    for i, d in enumerate(tqdm(dates_to_scrape, desc="Fetching", unit="day")):
        url  = build_url(d)
        html = fetch_page(url)

        if html is None:
            failed_days.append(str(d))
        else:
            records = parse_prices(html, d)
            if records:
                all_new_records.extend(records)
            else:
                empty_days.append(str(d))

        # Delay ngẫu nhiên để tránh bị block
        if i < len(dates_to_scrape) - 1:
            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    # Bước 3: lưu dữ liệu mới vào parquet
    if all_new_records:
        combined_df = append_rows_to_parquet(all_new_records, RAW_PARQUET)
        log.info("\n✅ Parquet đã cập nhật → %s  (tổng %d rows)",
                 RAW_PARQUET.name, len(combined_df))
    else:
        log.info("\nKhông có dữ liệu mới (có thể toàn cuối tuần / ngày lễ).")

    # Báo cáo kết quả scraping
    if empty_days:
        log.info("   Ngày không có data (nghỉ/lễ): %d ngày", len(empty_days))
    if failed_days:
        log.warning("   ⚠  Ngày thất bại: %s", ", ".join(failed_days))
        failed_log = RAW_SUNSIRS / "failed_days.txt"
        failed_log.write_text("\n".join(failed_days), encoding="utf-8")
        log.warning("   Đã lưu danh sách → %s", failed_log.name)

    # Bước 4: build file selected CSV
    build_selected_csv(RAW_PARQUET, SELECTED_CSV)

    log.info("=" * 60)
    log.info("🎉 Hoàn tất!")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
