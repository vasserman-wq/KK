@echo off
rem Сборка KK.exe. Запускать из папки проекта (рядом с KK.spec).
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if not exist kk\resources mkdir kk\resources
if exist "..\Настройки_ОК.xlsx" copy /y "..\Настройки_ОК.xlsx" kk\resources\ >nul
if exist "Настройки_ОК.xlsx" copy /y "Настройки_ОК.xlsx" kk\resources\ >nul
if exist "..\passport_macro\assets\app.ico" copy /y "..\passport_macro\assets\app.ico" kk\resources\ >nul
if not exist "kk\resources\Настройки_ОК.xlsx" echo ВНИМАНИЕ: нет Настройки_ОК.xlsx — встроенных настроек в exe не будет.
py -m pip install -r requirements.txt || goto :error
py -m pytest -q tests || echo Тесты не прошли (см. выше). Сборка продолжается.
py -m PyInstaller --noconfirm --clean KK.spec || goto :error
echo.
echo Готово: dist\KK.exe
goto :eof
:error
echo Ошибка сборки.
exit /b 1
