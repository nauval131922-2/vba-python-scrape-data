# ===== KONFIGURASI SCRAPER (VERSI SIMPLE) =====

# ========================================
# 1. KREDENSIAL LOGIN
# ========================================
# ===== KREDENSIAL (OBFUSCATED) =====
import base64

_ENC_USER = "bmF1dmFs"
_ENC_PASS = "MzEyYWRtaW4y"

def _decode(val):
    return base64.b64decode(val).decode("utf-8")

USERNAME = _decode(_ENC_USER)
PASSWORD = _decode(_ENC_PASS)


# ========================================
# 2. URL & ENDPOINT
# ========================================
LOGIN_URL = "https://buyapercetakan.mdthoster.com/login.html"
API_BASE  = "https://buyapercetakan.mdthoster.com/il/v1/akt/mrek/gr1"

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