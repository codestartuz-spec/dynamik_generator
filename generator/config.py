# -*- coding: utf-8 -*-
"""Loyiha sozlamalari: papkalar, fayl nomlari, qoidalar."""
from pathlib import Path

# C:\PythonLoyhalar\Report\Generator
BASE_DIR = Path(__file__).resolve().parent.parent
FILE_DIR = BASE_DIR / "file"        # CAS dan yuklab olingan hisobotlar
UPLOAD_DIR = BASE_DIR / "upload"    # web ilovaga yuklanadigan tayyor fayllar

# --- Kiruvchi fayllar -------------------------------------------------------
# Shu papkalardagi BARCHA .xlsx fayllar olinadi (fayllar soni cheklanmagan: 5 ta ham, 30 ta ham bo'lishi mumkin)
ARCHIVE_DIR = FILE_DIR / "Meter Archive Details"     # Meter Archive Details hisobotlari
METERING_DIR = FILE_DIR / "MeteringDevicesReport"    # MeteringDevicesReport bo'laklari
# Bitta fayl bo'lgan hisobotlar to'g'ridan-to'g'ri FILE_DIR ichida turadi:
DIRECT_FILE = "DirectMeter.xlsx"

# --- Chiquvchi fayl ---------------------------------------------------------
OUTPUT_NAME = "Details_Qashqadaryo HETK-{date}.xlsx"
OUTPUT_DATE_FORMAT = "%d-%m-%Y"
OUTPUT_SHEET = "Details"

# --- CSV + gzip (web ilovaga yuklash uchun, Excel bilan bir xil ma'lumot) ----
CSV_NAME = "Details_Qashqadaryo HETK-{date}.csv.gz"
CSV_EMPTY = ""                              # bo'sh qiymat CSV da bo'sh qoldiriladi (Excelda '-')
CSV_DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"   # obxod vaqti formati
CSV_GZIP_LEVEL = 6

# --- 1-qoida: olib tashlanadigan ustunlar (Meter Archive Details) -----------
DROP_COLUMNS = [
    "Баланс ТП",
    "Версия прошивки",
    "дата установки",
    "Дата регистрации",
    "Статус Реле",
    "Время обновления",
]

# --- Ustun nomlari (manba fayllarda) ----------------------------------------
ARCHIVE_METER_COL = "Номер счетчика"
METERING_METER_COL = "Meter Number"
METERING_POLL_COL = "Poll time (last)"
DIRECT_METER_COL = "Meter No."

# --- 2-qoida: yangi ustunlar ------------------------------------------------
POLL_COL_NAME = "Oxirgi obxod vaqti"
DAYS_COL_NAME = "Obxod kuni"   # joriy kun 16:00 dan obxodgacha o'tgan to'liq kunlar
DIRECT_COL_NAME = "DirectMeter"

# Hisoblash vaqti: bugungi kun soat REFERENCE_HOUR:00
REFERENCE_HOUR = 23

# (ustun nomi, pastki chegara kun, yuqori chegara kun)  -> lo < lag <= hi
# Birinchi oraliq: lag <= 2 kun (48 soat), manfiy (16:00 dan keyingi) ham shunga kiradi.
# Oxirgisi: lag > 90 kun.
BUCKETS = [
    ("48 soat",   None, 2),
    ("2-7 kun",   2,    7),
    ("7-30 kun",  7,    30),
    ("30-90 kun", 30,   90),
    ("90+ kun",   90,   None),
]
NO_VALUE = "-"   # obxod vaqti umuman bo'lmasa shu belgi qo'yiladi

# --- Koordinatalar (широта / долгота) tekshiruvi -----------------------------
REGION_GEOJSON = BASE_DIR / "generator" / "data" / "qashqadaryo.geojson"
LAT_COL = "широта"
LON_COL = "долгота"
BORDER_TOLERANCE_KM = 2.0   # chegara yaqinidagi nuqtalarga yo'l qo'yiladigan xato (km)

EXCEL_MAX_ROWS = 1_048_576

# ============================================================================
# 2-QISM: Area (transformatorlar) hisoboti
# ============================================================================
AREA_GLOB = "AreaInformationReport*.xlsx"
COMMISSIONING_GLOB = "CommissioningStatus*.xlsx"
ONLINE_GLOB = "Online_Device_Monitroing_Report*.xlsx"   # bir nechta bo'lsa eng yangisi olinadi

AREA_OUTPUT_NAME = "Area_Qashqadaryo HETK-{date}.xlsx"
AREA_CSV_NAME = "Area_Qashqadaryo HETK-{date}.csv.gz"
AREA_SHEET = "Area"

# AreaInformationReport dan qoldiriladigan ustunlar (qolganlari olib tashlanadi)
AREA_KEEP_COLUMNS = ["ESP Code", "ESP name", "Substation Code", "Substation",
                     "Feeder Code", "Feeder", "TR Code", "TR", "Manager Name"]
AREA_ESP_COL = "ESP name"
AREA_TR_COL = "TR"
AREA_TRCODE_COL = "TR Code"
AREA_MANAGER_COL = "Manager Name"

# CommissioningStatus ustunlari
COMM_ESP_COL = "ESP Name"
COMM_TR_COL = "TR No."
COMM_DCU_COL = "DCU No."
COMM_INSTALLED_COL = "Installed(pcs)"       # konsentratorga iste'molchi biriktirilganligi (>0 -> 1)
COMM_DAILY_COL = "Daily Data %"

# Online_Device_Monitroing_Report ustunlari
ONLINE_DCU_COL = "№ Устройства"
ONLINE_STATUS_COL = "Статус"
ONLINE_TIME_COL = "Время онлайна/офлайна"
ONLINE_TYPE_COL = "Тип устройства"
ONLINE_TYPE_DCU = "DCU"
ONLINE_TIME_FORMAT = "%d/%m/%Y %H:%M:%S"
STATUS_OFFLINE = "offline"                   # solishtirish katta-kichik harfga bog'liq emas

# Yangi ustun nomlari
AREA_DCU_COL = "DCU No."
AREA_CONSUMER_COL = "Iste'molchi biriktirilgan"
AREA_LEVEL_COL = "Obxod darajasi"
AREA_STATE_COL = "Konsentrator holati"
AREA_OFFLINE_COL = "Offline"
AREA_OFFLINE_LONG_COL = "3 kundan ortiq offline"

# "3 kundan ortiq offline": bugun 00:00 dan shuncha soatdan oldin offline bo'lgan
OFFLINE_LONG_HOURS = 48

# Daily Data % oraliqlari: (ustun nomi, tur, pastki, yuqori)
#   "eq": qiymat == pastki;   "range": pastki < qiymat <= yuqori (yuqori None -> cheksiz)
DAILY_BINS = [
    ("Daily 0%",       "eq",    0,  None),
    ("Daily 1-30%",    "range", 0,  30),
    ("Daily 30-50%",   "range", 30, 50),
    ("Daily 50-80%",   "range", 50, 80),
    ("Daily 80-100%",  "range", 80, None),
]
