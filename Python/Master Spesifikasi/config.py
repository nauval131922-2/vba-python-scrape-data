# ===== KONFIGURASI SCRAPER (VERSI SIMPLE) =====

# ========================================
# 1. KREDENSIAL LOGIN (ENV)
# ========================================
import os
from dotenv import load_dotenv

load_dotenv()  # baca file .env

# USERNAME = os.getenv("SCRAPER_USERNAME")
USERNAME = "nauval"
# PASSWORD = os.getenv("SCRAPER_PASSWORD")
PASSWORD = "312admin2"

# if not USERNAME or not PASSWORD:
#     raise ValueError("USERNAME atau PASSWORD belum diset di .env")


# ========================================
# 2. URL & ENDPOINT
# ========================================
LOGIN_URL = "https://buyapercetakan.mdthoster.com/login.html"
API_BASE  = "https://buyapercetakan.mdthoster.com/il/v1/gnr/mspsfk/gr1"

# ========================================
# 3. SETTING SCRAPING
# ========================================
# Limit data per request
LIMIT = 10000

# Nama file output (tanpa ekstensi, otomatis jadi .xlsx)
FILE_NAME = "temp"

# ========================================
# 4. REQUEST PAYLOAD (SIMPLE)
# ========================================
# Karena tidak perlu filter tanggal/cabang/status,
# payload jadi sangat sederhana
EXTRA_REQUEST = {
    "bsearch": {}  # Kosong = ambil semua data tanpa filter
}

# ========================================
# 5. STYLING EXCEL
# ========================================
# Warna zebra style (hex kode tanpa #)
ROW_COLOR_ODD  = "FFFFFF"  # Putih
ROW_COLOR_EVEN = "F9F9F9"  # Abu-abu muda
HEADER_COLOR   = "DDDDDD"  # Abu-abu header

# Maksimal lebar kolom Excel
MAX_COL_WIDTH = 30