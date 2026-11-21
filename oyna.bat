@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"
title CAS generatori - oyna

echo ==================================================
echo   CAS generatori - oynali ilova
echo ==================================================
echo.

if exist "proxy.txt" set /p PROXY=<"proxy.txt"
if not defined PROXY goto :find_python
if /i "%PROXY:~0,7%"=="http://" set "PROXY=%PROXY:~7%"
set "HTTP_PROXY=http://%PROXY%"
set "HTTPS_PROXY=http://%PROXY%"
echo   Proksi: %PROXY%

:find_python
rem Oyna uchun alohida muhit (.venv_oyna): generator.bat ishlatadigan .venv ga tegilmaydi.
rem Python 3.12 bo'lsa shuni olamiz (eng barqaror), bo'lmasa mavjud Python.
set "PY="
py -3.12 --version >nul 2>nul && set "PY=py -3.12"
if not defined PY py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY python --version >nul 2>nul && set "PY=python"
if not defined PY goto :no_python
echo   Python: %PY%

if exist ".venv_oyna\Scripts\python.exe" goto :install
echo   Virtual muhit yaratilmoqda .venv_oyna ...
%PY% -m venv .venv_oyna
if errorlevel 1 goto :fail

:install
echo   Kutubxonalar tekshirilmoqda ...
".venv_oyna\Scripts\python.exe" -m pip install xlsxwriter customtkinter --disable-pip-version-check
if errorlevel 1 goto :pip_fail

echo.
echo   Oyna ochilmoqda. Bu qora oynani oyna ishlab turganda yopmang.
echo.
".venv_oyna\Scripts\python.exe" "generator\app.py"
if errorlevel 1 goto :app_fail
exit /b 0

:no_python
echo   XATO: Python topilmadi. Python 3.8 yoki yangisini o'rnating va "Add to PATH" ni belgilang.
goto :fail

:pip_fail
echo.
echo   XATO: kutubxonalarni o'rnatib bo'lmadi. Internet va proxy.txt ni tekshiring.
echo   Yuqoridagi xato matnini yuboring.
goto :fail

:app_fail
echo.
echo   Oyna xato bilan yopildi. Yuqoridagi xato matnini yuboring.
echo   Qo'shimcha: logs\oyna_xato.log

:fail
echo.
pause
exit /b 1
