# config/paths.py
from pathlib import Path

BASE_DIR         = Path(__file__).resolve().parent.parent

# Đường dẫn data
RAW_SUNSIRS      = BASE_DIR / "data" / "raw"       / "commodities" / "sunsirs"
PROCESSED_COMMOD = BASE_DIR / "data" / "processed" / "commodities"
REPORTS_COMMOD   = BASE_DIR / "reports"             / "commodities"

# Tạo thư mục nếu chưa tồn tại
for _dir in [RAW_SUNSIRS, PROCESSED_COMMOD, REPORTS_COMMOD]:
    _dir.mkdir(parents=True, exist_ok=True)
