@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ==================================================
echo   CAS Generator.exe  ni yig'ish (bir marta kerak)
echo ==================================================
echo.

rem --- Proksi (office): proxy.txt bo'lsa ishlatiladi; uy kompyuterida fayl bo'lmaydi ---
if exist "proxy.txt" (
    set /p PROXY=<"proxy.txt"
)
if defined PROXY (
    if /i "%PROXY:~0,7%"=="http://" set "PROXY=%PROXY:~7%"
    set "HTTP_PROXY=http://%PROXY%"
    set "HTTPS_PROXY=http://%PROXY%"
    echo   Proksi: %PROXY%
) else (
    echo   Proksisiz ishlaydi.
)
echo.

set "PY="
py -3.12 --version >nul 2>nul && set "PY=py -3.12"
if not defined PY py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY python --version >nul 2>nul && set "PY=python"
if not defined PY goto :no_python

if exist ".venv_oyna\Scripts\python.exe" goto :venv_ok
echo   Virtual muhit yaratilmoqda...
%PY% -m venv .venv_oyna
if errorlevel 1 goto :fail
:venv_ok

echo   Kutubxonalar o'rnatilmoqda (customtkinter, pyinstaller)...
".venv_oyna\Scripts\python.exe" -m pip install -r generator\requirements-gui.txt --disable-pip-version-check -q
if errorlevel 1 (
    echo.
    echo   XATO: kutubxonalarni o'rnatib bo'lmadi. Internet va proxy.txt ni tekshiring.
    goto :fail
)

echo   .exe yig'ilmoqda (1-3 daqiqa)...
if exist "CAS Generator" rmdir /s /q "CAS Generator"
".venv_oyna\Scripts\python.exe" -m PyInstaller --noconfirm --clean --windowed --name "CAS Generator" ^
    --paths "%CD%\generator" ^
    --add-data "%CD%\generator\data;data" ^
    --collect-all customtkinter ^
    --distpath "%CD%" --workpath "%CD%\build_tmp" --specpath "%CD%\build_tmp" ^
    "%CD%\generator\app.py"
if errorlevel 1 goto :fail

if exist "build_tmp" rmdir /s /q "build_tmp"
echo.
echo ==================================================
echo   TAYYOR:  %CD%\CAS Generator\CAS Generator.exe
echo.
echo   Uni ishga tushiring (ish stoliga yorliq qo'yish mumkin).
echo   file\ va upload\ papkalari shu papkalar yonida turishi kerak.
echo ==================================================
pause
exit /b 0

:no_python
echo   XATO: Python topilmadi. Python 3.8 yoki yangisini o'rnating va "Add to PATH" ni belgilang.

:fail
echo.
echo   Yig'ish to'xtadi.
pause
exit /b 1
