@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ==================================================
echo   CAS hisobot generatori  (Qashqadaryo HETK)
echo ==================================================
echo.
echo   Qaysi kompyuterdan ishlayapsiz?
echo.
echo     1 - Uy kompyuteri      (internet proksisiz)
echo     2 - Office kompyuteri  (proksi orqali)
echo.
choice /c 12 /n /m "  Tanlang (1 yoki 2): "
if errorlevel 2 goto :office
goto :home

:home
set "GEN_NETWORK=uy kompyuteri (proksisiz)"
set "HTTP_PROXY="
set "HTTPS_PROXY="
set "ALL_PROXY="
set "http_proxy="
set "https_proxy="
set "all_proxy="
set "NO_PROXY=*"
set "no_proxy=*"
echo.
echo   Tanlandi: Uy kompyuteri. Internet proksisiz ishlaydi.
goto :python

:office
set "PROXY="
if exist "proxy.txt" set /p PROXY=<"proxy.txt"
if defined PROXY goto :office_set
echo.
echo   Office uchun proksi manzili hali saqlanmagan.
echo   Format: manzil:port   (masalan proksi.example:3128)
set /p PROXY="  Proksi manzilini kiriting: "
if not defined PROXY goto :office_bad
> "proxy.txt" echo %PROXY%
echo   Saqlandi: proxy.txt  (bu fayl GitHub ga yuklanmaydi)

:office_set
if /i "%PROXY:~0,7%"=="http://" set "PROXY=%PROXY:~7%"
set "GEN_NETWORK=office kompyuteri (proksi %PROXY%)"
set "HTTP_PROXY=http://%PROXY%"
set "HTTPS_PROXY=http://%PROXY%"
set "ALL_PROXY=http://%PROXY%"
set "http_proxy=http://%PROXY%"
set "https_proxy=http://%PROXY%"
set "all_proxy=http://%PROXY%"
set "NO_PROXY=localhost,127.0.0.1"
set "no_proxy=localhost,127.0.0.1"
echo.
echo   Tanlandi: Office kompyuteri. Proksi: %PROXY%
echo   (Proksini o'zgartirish uchun proxy.txt faylini o'chiring)
goto :python

:office_bad
echo.
echo   Proksi manzili kiritilmadi. Qaytadan ishga tushiring.
pause
exit /b 1

:python
echo.
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo   XATO: Python topilmadi. Python 3.8+ ni o'rnating va "Add to PATH" ni belgilang.
    pause
    exit /b 1
)

rem --- Virtual muhit (birinchi ishga tushirishda) ---
if not exist ".venv\Scripts\python.exe" (
    echo   Virtual muhit yaratilmoqda...
    %PY% -m venv .venv
    if errorlevel 1 goto :fail
)

rem --- Kutilgan kutubxonalar o'rnatilgan bo'lsa, internet kerak bo'lmaydi ---
".venv\Scripts\python.exe" -m pip install -r generator\requirements.txt --disable-pip-version-check -q
if errorlevel 1 (
    echo.
    echo   XATO: kutubxonalarni o'rnatib bo'lmadi. Internet va proksi sozlamasini tekshiring.
    echo   Agar xato proksiga bog'liq bo'lsa, proxy.txt ni o'chirib, qayta tanlang.
    goto :fail
)

".venv\Scripts\python.exe" generator\main.py %*
set "RC=%errorlevel%"
echo.
if not "%RC%"=="0" echo   Dastur xato bilan tugadi (kod %RC%).
pause
exit /b %RC%

:fail
pause
exit /b 1
