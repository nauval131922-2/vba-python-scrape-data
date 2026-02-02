# ===== KONFIGURASI SCRAPER =====
from datetime import datetime

# ========================================
# 1. KREDENSIAL LOGIN
# ========================================
USERNAME = "nauval"
PASSWORD = "312admin2"

# ========================================
# 2. URL & ENDPOINT
# ========================================
LOGIN_URL = "https://buyapercetakan.mdthoster.com/login.html"
API_BASE = "https://buyapercetakan.mdthoster.com/il/v1/cfg/usr_log/grid"

# ========================================
# 3. SETTING SCRAPING
# ========================================
# Limit data per request (maksimal data yang diambil per hari)
LIMIT = 10000

# Nama file output (tanpa ekstensi, otomatis jadi .xlsx)
FILE_NAME = "temp"

# ========================================
# 4. TANGGAL (FALLBACK)
# ========================================
# Tanggal ini hanya dipakai jika tidak input manual via menu
# Kalau pakai run_scraper_menu.bat, tanggal ini tidak terpakai
TODAY = datetime.now().strftime("%Y-%m-%d")
START_DATE = "2026-01-01"
END_DATE = "2026-01-31"

# ========================================
# 5. SORTING
# ========================================
SORT_BY = "Datetime"  # Kolom untuk sorting: "Datetime", "Level", "User", dll
SORT_ORDER = "desc"  # Urutan: "asc" (naik) atau "desc" (turun)

# ========================================
# 6. STYLING EXCEL
# ========================================
# Warna zebra style (hex kode tanpa #)
ROW_COLOR_ODD = "FFFFFF"  # Putih
ROW_COLOR_EVEN = "F9F9F9"  # Abu-abu muda
HEADER_COLOR = "DDDDDD"  # Abu-abu header

# Maksimal lebar kolom Excel
MAX_COL_WIDTH = 30
