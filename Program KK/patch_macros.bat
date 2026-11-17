@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "MACRO1=..\SOLIDWORKS_KC_v16.5.15_review (1).txt"
set "MACRO2=..\2 Макрос (2).txt"
if not exist "%MACRO1%" goto :miss1
if not exist "%MACRO2%" goto :miss2
if not exist "tools\patch_macros.py" goto :misstool
where py >nul 2>nul
if errorlevel 1 (set "PY=python") else (set "PY=py")
%PY% tools\patch_macros.py "%MACRO1%" "%MACRO2%" --out vba\patched
if errorlevel 1 goto :fail
echo Готово. Результат в папке vba\patched
pause
exit /b 0
:miss1
echo Не найден файл: %MACRO1%
goto :fail
:miss2
echo Не найден файл: %MACRO2%
goto :fail
:misstool
echo Не найден файл tools\patch_macros.py
goto :fail
:fail
echo Остановлено, ничего не записано.
pause
exit /b 1