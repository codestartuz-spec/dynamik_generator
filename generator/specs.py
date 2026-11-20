# -*- coding: utf-8 -*-
"""
Har bir CAS hisoboti uchun ustunlar tavsifi (schema.Field): ichki nom, chiqish nomi, joylashuv (0 dan),
ru / en / uz sinonimlar va ma'lumot turi. Yangi til qo'shish uchun faqat sinonimlar ro'yxatini to'ldirish kifoya;
sinonim topilmasa ham ustun joylashuvi bo'yicha aniqlanadi.
"""
from schema import Field, Spec

# --------------------------------------------------- Meter Archive Details ----
# Chiqish nomlari (birinchi argument) - oldingi (rus tilidagi) nomlar, shuning uchun til o'zgarsa ham
# natijadagi sarlavhalar bir xil bo'lib qoladi. drop=True - 1-qoida bo'yicha olib tashlanadigan ustunlar.
ARCHIVE = Spec("Meter Archive Details", [
    Field("esp", "ЭСП", 0, ["ESP", "ESP name", "Район", "Tuman"], "text"),
    Field("tr", "ТП", 1, ["TR", "TP", "TR name", "Transformator"], None),
    Field("tr_balance", "Баланс ТП", 2, ["TR Balance", "TP Balance", "Balans TP", "Баланс"], drop=True),
    Field("firmware", "Версия прошивки", 3, ["Firmware Version", "Firmware", "Proshivka versiyasi",
                                              "Dasturiy ta'minot versiyasi"], drop=True),
    Field("account", "Лицевой счет потребителя", 4, ["Лицевой счет", "Account ID", "Account No", "Account",
                                                      "Hisob raqami", "Shaxsiy hisob raqami"]),
    Field("meter", "Номер счетчика", 5, ["Meter Number", "Meter No.", "Meter No", "Hisoblagich raqami",
                                          "Schetchik raqami", "Счетчик №", "№ счетчика"], "id", required=True),
    Field("meter_type", "Тип счетчика", 6, ["Meter Type", "Hisoblagich turi"]),
    Field("comm_type", "Тип связи", 7, ["Communication Type", "Comm Type", "Aloqa turi"]),
    Field("cust_type", "Тип потребителя", 8, ["Customer Type", "Iste'molchi turi"]),
    Field("installed", "дата установки", 9, ["Installed Date", "Installation Date", "Date Installed",
                                              "O'rnatilgan sana", "O'rnatish sanasi"], drop=True),
    Field("registered", "Дата регистрации", 10, ["Registration Date", "Registered Date", "Date Registered",
                                                  "Ro'yxatdan o'tgan sana", "Ro'yxatga olingan sana"], drop=True),
    Field("ct", "Коэффициент Т/Т", 11, ["CT Ratio", "CT", "TT koeffitsiyenti"]),
    Field("pt", "Коэффициент Т/Н", 12, ["PT Ratio", "PT", "TN koeffitsiyenti"]),
    Field("lat", "широта", 13, ["Latitude", "Lat", "Kenglik"]),
    Field("lon", "долгота", 14, ["Longitude", "Lon", "Long", "Uzunlik"]),
    Field("status", "Статус оборудования", 15, ["Equipment Status", "Device Status", "Meter Status",
                                                 "Uskuna holati", "Qurilma holati"]),
    Field("relay", "Статус Реле", 16, ["Relay Status", "Rele holati", "Relay"], drop=True),
    Field("updated", "Время обновления", 17, ["Update Time", "Updated Time", "Last Update", "Update Date",
                                               "Yangilanish vaqti"], drop=True),
    Field("cust_name", "Имя потребителя", 18, ["Customer Name", "Name", "Iste'molchi nomi",
                                                "Iste'molchi ismi"]),
    Field("address", "Адрес установки", 19, ["Installation Address", "Address", "Customer Address",
                                              "O'rnatish manzili", "Manzil"]),
], min_cells=15)

# --------------------------------------------------- MeteringDevicesReport ---
METERING = Spec("MeteringDevicesReport", [
    Field("meter", "Meter Number", 10, ["Meter No.", "Meter No", "Номер счетчика", "№ счетчика",
                                         "Hisoblagich raqami"], "id", required=True),
    Field("poll", "Poll time (last)", 15, ["Last Poll Time", "Poll Time", "Время последнего опроса",
                                            "Время опроса", "Oxirgi so'rov vaqti", "So'rov vaqti"],
          "datetime", required=True),
], min_cells=12)

# --------------------------------------------------- DirectMeter -------------
DIRECT = Spec("DirectMeter", [
    Field("esp", "ESP Name", 1, ["ESP", "ЭСП", "ESP nomi"]),
    Field("meter", "Meter No.", 6, ["Meter No", "Meter Number", "Номер счетчика", "№ счетчика",
                                     "Hisoblagich raqami"], "id", required=True),
], min_cells=10)

# --------------------------------------------------- AreaInformationReport ---
# Chiqish nomlari config.AREA_KEEP_COLUMNS bilan bir xil bo'lishi kerak.
AREA = Spec("AreaInformationReport", [
    Field("esp_code", "ESP Code", 3, ["Код ЭСП", "ESP kodi"], "text"),
    Field("esp", "ESP name", 4, ["ESP Name", "ЭСП", "Название ЭСП", "ESP nomi"], "text", required=True),
    Field("sub_code", "Substation Code", 5, ["Код подстанции", "Podstansiya kodi"]),
    Field("sub", "Substation", 6, ["Подстанция", "Podstansiya"]),
    Field("feeder_code", "Feeder Code", 7, ["Код фидера", "Fider kodi"]),
    Field("feeder", "Feeder", 8, ["Фидер", "Fider"]),
    Field("tr_code", "TR Code", 9, ["Код ТП", "TP kodi"]),
    Field("tr", "TR", 10, ["ТП", "TP", "TR name"], required=True),
    Field("manager", "Manager Name", 12, ["Manager", "Имя менеджера", "Менеджер", "Menejer ismi", "Menejer"]),
], min_cells=12)

# --------------------------------------------------- CommissioningStatus -----
COMMISSIONING = Spec("CommissioningStatus", [
    Field("esp", "ESP Name", 1, ["ESP", "ЭСП", "ESP nomi"], "text", required=True),
    Field("tr", "TR No.", 2, ["TR No", "TR", "ТП", "№ ТП", "ТП №", "TP raqami"], required=True),
    Field("dcu", "DCU No.", 3, ["DCU No", "DCU", "№ DCU", "DCU №", "Konsentrator raqami", "Концентратор"],
          required=True),
    Field("installed", "Installed(pcs)", 7, ["Installed (pcs)", "Installed", "Установлено(шт)", "Установлено",
                                              "O'rnatilgan(dona)", "O'rnatilgan"], "number"),
    Field("daily", "Daily Data %", 10, ["Daily Data", "Ежедневные данные %", "Ежедневные данные",
                                         "Kunlik ma'lumot %", "Kunlik ma'lumot"], "number"),
], min_cells=9)

# --------------------------------------------------- Online_Device_Monitroing_Report
ONLINE = Spec("Online_Device_Monitroing_Report", [
    Field("type", "Тип устройства", 4, ["Device Type", "Qurilma turi"]),
    Field("dcu", "№ Устройства", 5, ["Device No.", "Device No", "Device Number", "Device ID", "№ устройства",
                                      "Qurilma raqami", "Qurilma №"], "id", required=True),
    Field("status", "Статус", 6, ["Status", "Holat", "Holati"], "text", required=True),
    Field("time", "Время онлайна/офлайна", 7, ["Online/Offline Time", "Online Offline Time", "Status Time",
                                                 "Onlayn/oflayn vaqti"], "datetime", required=True),
], min_cells=12)
