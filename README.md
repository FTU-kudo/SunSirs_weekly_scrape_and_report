# 📊 SunSirs Commodity Price Tracker

Automated weekly pipeline that scrapes commodity prices from [SunSirs](https://sunsirs.com/uk/) and delivers a formatted Excel report every Sunday at 20:30 (Vietnam time, UTC+7).

---

## ✨ Features

- **Incremental scraping** — only fetches new data since the last run (first run scrapes full history from 2024-01-01)
- **Fault-tolerant** — automatic retry with exponential backoff on network errors; failed days are logged to `failed_days.txt` and retried next run
- **Compact storage** — raw data stored as `.parquet` (~5× smaller than Excel)
- **Auto-delivery** — formatted `.xlsx` report emailed to recipients every Sunday at 20:30
- **Fully automated** — runs on GitHub Actions, no local machine required

---

## 🗂️ Project Structure

```
├── .github/
│   └── workflows/
│       └── sunsirs_weekly.yml      # GitHub Actions: scrape + email
│
├── config/
│   ├── __init__.py
│   └── paths.py                    # Centralised path configuration
│
├── src/
│   ├── commodities/
│   │   ├── __init__.py
│   │   └── sunsirs_scraper.py      # Core scraper (incremental, batch mode)
│   └── reports/
│       ├── __init__.py
│       └── send_report.py          # Build XLSX + send via Gmail
│
├── data/                           # Auto-created by workflow
│   ├── raw/commodities/sunsirs/
│   │   └── sunsirs_raw.parquet     # Master dataset (all commodities)
│   └── processed/commodities/
│       └── selected_commodity_prices.csv
│
├── .gitignore
├── requirements.txt
└── README.md
```

---

## 📦 Commodities Tracked

| Commodity | Unit |
|---|---|
| Urea | CNY/ton |
| Phosphorus yellow | CNY/ton |
| Phosphoric acid | CNY/ton |
| Hydrochloric acid | CNY/ton |
| Sulfuric acid | CNY/ton |

---

## ⚙️ How It Works

```
Every Sunday 20:00 (UTC+7)
        │
        ▼
┌─────────────────────────────┐
│  Job 1: scrape-sunsirs      │
│  • Checkout repo            │
│  • Read existing parquet    │
│  • Scrape only new days     │  ← ~30 seconds after first run
│  • Commit data back to repo │
└─────────────────────────────┘
        │
        ▼ (on success)
┌─────────────────────────────┐
│  Job 2: send-email          │
│  • Download parquet         │
│  • Build formatted XLSX     │
│  • Send via Gmail SMTP      │  ← recipients receive at ~20:30
└─────────────────────────────┘
```

---

## 🚀 Setup

### 1. Fork or clone this repo

```bash
git clone https://github.com/your-username/sunsirs-scraping-weekly.git
```

### 2. Add GitHub Secrets

Go to **Settings → Secrets and variables → Actions** and add:

| Secret | Description |
|---|---|
| `GMAIL_USER` | Gmail address used to send reports |
| `GMAIL_APP_PASSWORD` | [Gmail App Password](https://myaccount.google.com/apppasswords) (16 characters) |
| `RECIPIENT_EMAILS` | Comma-separated list: `a@gmail.com,b@company.com` |

### 3. Enable GitHub Actions

Go to the **Actions** tab → enable workflows if prompted.

### 4. Test manually

Go to **Actions → SunSirs Weekly Scrape & Report → Run workflow**.

---

## 📋 Requirements

```
requests
beautifulsoup4
pandas
openpyxl
tqdm
lxml
pyarrow
```

---

## 📅 Schedule

| Time (UTC+7) | Job |
|---|---|
| Sunday 20:00 | Scrape new data + commit to repo |
| Sunday ~20:30 | Send Excel report via email |

---

## ⚠️ Notes

- First run scrapes full history (~900+ days, ~1 hour). Subsequent runs take under 1 minute.
- SunSirs data covers Chinese commodity markets. Weekends and Chinese public holidays return no data (expected behaviour).
- Never commit `.env` or credential files — all secrets are managed via GitHub Secrets.

---

## 📄 Data Source

Data sourced from [SunSirs — China Commodity Data Group](https://sunsirs.com/).
