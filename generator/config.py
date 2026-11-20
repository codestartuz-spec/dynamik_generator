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
# Bitta fayl bo'lgan hisobotlar to'g'ridan-to'g'ri FILE_DIR ichida turadi. Ular NOMI SHU SO'Z BILAN
# BOSHLANADIGAN .xlsx fayl sifatida topiladi (DirectMeter.xlsx, DirectMeter(2).xlsx, DirectMeter (1).xlsx ...).
# Bir nechta mos fayl bo'lsa eng yangisi (o'zgargan vaqti bo'yicha) olinadi.
DIRECT_PREFIXES = ["DirectMeter"]

# --- Chiquvchi fayl ---------------------------------------------------------
OUTPUT_NAME = "Details_Qashqadaryo HETK-{date}.xlsx"
OUTPUT_DATE_FORMAT = "%d-%m-%Y"
OUTPUT_SHEET = "Details"

# --- CSV + gzip (web ilovaga yuklash uchun, Excel bilan bir xil ma'lumot) ----
CSV_NAME = "Details_Qashqadaryo HETK-{date}.csv.gz"
CSV_EMPTY = ""                              # bo'sh qiymat CSV da bo'sh qoldiriladi (Excelda '-')
CSV_DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"   # obxod vaqti formati
CSV_GZIP_LEVEL = 6

# --- Ustunlar tilga bog'liq emas ---------------------------------------------
# Manba fayllardagi ustunlar nomi CAS tiliga qarab o'zgarishi mumkin, shuning uchun ular nomi bilan emas,
# specs.py dagi tavsif (ko'p tilli sinonimlar + ustun joylashuvi + ma'lumot turi tekshiruvi) bo'yicha topiladi.
# 1-qoida (olib tashlanadigan 6 ustun: Баланс ТП, Версия прошивки, дата установки, Дата регистрации,
# Статус Реле, Время обновления) ham specs.ARCHIVE da drop=True bilan belgilangan.
# Quyidagilar - CHIQISH fayldagi ustun nomlari (manba tilidan qat'i nazar bir xil chiqadi):
ARCHIVE_METER_COL = "Номер счетчика"

# --- 2-qoida: yangi ustunlar ------------------------------------------------
POLL_COL_NAME = "Oxirgi obxod vaqti"
DAYS_COL_NAME = "Obxod kuni"   # joriy kun 16:00 dan obxodgacha o'tgan to'liq kunlar
DIRECT_COL_NAME = "DirectMeter"

# Hisoblash vaqti: bugungi sana + shu vaqt. Erkin belgilanadi, masalan:
#   REFERENCE_HOUR = 16        -> 16:00
#   REFERENCE_HOUR = 9         -> 09:00
#   REFERENCE_HOUR = "09:30"   -> 09:30   ("S:DD" yoki "SS:DD" ko'rinishida)
REFERENCE_HOUR = 16


def reference_time():
    """REFERENCE_HOUR dan (soat, daqiqa) ni qaytaradi."""
    v = REFERENCE_HOUR
    if isinstance(v, str):
        parts = v.strip().replace(".", ":").split(":")
        h, m = int(parts[0]), int(parts[1]) if len(parts) > 1 and parts[1] else 0
    else:
        h, m = int(v), int(round((float(v) - int(v)) * 60))
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ValueError("REFERENCE_HOUR noto'g'ri: %r (0:00 - 23:59 oralig'ida bo'lishi kerak)" % (v,))
    return h, m


# --- Fayl hajmini kamaytirish: matnli qiymatlar raqamga almashtiriladi (Details, ikkala til uchun) --------
# Katta-kichik harf va tinish belgilari e'tiborga olinmaydi. Ro'yxatda yo'q qiymat o'zgarishsiz qoladi
# (va logda ogohlantiriladi) - yangi qiymat chiqsa, shu yerga qo'shing.
CUSTOMER_TYPE_COL = "Тип потребителя"
CUSTOMER_TYPE_MAP = {
    "Industrial": 1, "Юридические": 1,
    "Residential": 2, "Бытовые": 2,
    "Balance Meter": 3,
}
EQUIPMENT_STATUS_COL = "Статус оборудования"
EQUIPMENT_STATUS_MAP = {
    "Commissioned": 1, "Введено в эксплуатацию": 1,
    "Commissioning": -1, "Запускается в работу": -1,
    "Installed": 0,
    "Initial": 0, "Начальный": 0,
}

# Obxodgacha o'tgan vaqt (hisoblash vaqti - oxirgi obxod vaqti) soat/soniya aniqligida olinadi.
# (ustun nomi, pastki chegara kun, yuqori chegara kun)  ->  pastki*24 soat <= o'tgan vaqt < yuqori*24 soat
#   48 soat   : o'tgan vaqt < 48 soat            (manfiy, ya'ni hisoblash vaqtidan keyingi obxod ham shunga kiradi)
#   2-7 kun   : 48 soat <= o'tgan vaqt < 168 soat
#   7-30 kun  : 168 soat <= o'tgan vaqt < 720 soat
#   30-90 kun : 720 soat <= o'tgan vaqt < 2160 soat
#   90+ kun   : o'tgan vaqt >= 2160 soat
# "Obxod kuni" = o'tgan to'liq sutkalar: 24 soatdan kam -> 0, 24-48 soat -> 1, 48-72 soat -> 2 ...
BUCKETS = [
    ("48 soat",   None, 2),
    ("2-7 kun",   2,    7),
    ("7-30 kun",  7,    30),
    ("30-90 kun", 30,   90),
    ("90+ kun",   90,   None),
]
NO_VALUE = "-"   # (ishlatilmaydi: obxod vaqti yo'q bo'lsa katak bo'sh qoladi)
# "Obxod kuni" ustunida 0 (obxod bugun bo'lgan) ham bo'sh qoldiriladi. 0 ni ko'rsatish uchun False qiling.
BLANK_ZERO_DAYS = True

# --- Koordinatalar (широта / долгота) tekshiruvi -----------------------------
REGION_GEOJSON = BASE_DIR / "generator" / "data" / "qashqadaryo.geojson"
LAT_COL = "широта"
LON_COL = "долгота"
BORDER_TOLERANCE_KM = 2.0   # chegara yaqinidagi nuqtalarga yo'l qo'yiladigan xato (km)

EXCEL_MAX_ROWS = 1_048_576

# ============================================================================
# 2-QISM: Area (transformatorlar) hisoboti
# ============================================================================
# Fayl nomi shu so'zlar bilan boshlansa bo'ldi (oxirida (1), (2), sana, vaqt... bo'lishi mumkin).
# Bir nechta mos fayl bo'lsa eng yangisi olinadi.
AREA_PREFIXES = ["AreaInformationReport"]
COMMISSIONING_PREFIXES = ["CommissioningStatus"]
ONLINE_PREFIXES = ["Online_Device_Monitroing_Report", "Online_Device_Monitoring_Report"]

AREA_OUTPUT_NAME = "Area_Qashqadaryo HETK-{date}.xlsx"
AREA_CSV_NAME = "Area_Qashqadaryo HETK-{date}.csv.gz"
AREA_SHEET = "Area"

# AreaInformationReport dan qoldiriladigan ustunlar (qolganlari olib tashlanadi) - CHIQISH nomlari.
# (Manba ustunlari specs.AREA da tavsiflangan; tartib va nomlar u yerdagi bilan bir xil bo'lsin.)
AREA_KEEP_COLUMNS = ["ESP Code", "ESP name", "Substation Code", "Substation",
                     "Feeder Code", "Feeder", "TR Code", "TR", "Manager Name"]
AREA_ESP_COL = "ESP name"
AREA_TR_COL = "TR"
AREA_TRCODE_COL = "TR Code"
AREA_MANAGER_COL = "Manager Name"

# Konsentrator (DCU) turi va holati (til farqidan qat'i nazar ishlashi uchun bir nechta yozilishi)
ONLINE_TYPE_DCU = "DCU"
# Holat qiymati shu boshlanishlardan biri bilan boshlansa "offline" hisoblanadi
# (Offline / Оффлайн / Офлайн / Oflayn ...). Online / Онлайн / Onlayn bunga kirmaydi.
STATUS_OFFLINE_PREFIXES = ("offline", "оффлайн", "офлайн", "oflayn", "offlayn", "off", "офф", "офл")

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
