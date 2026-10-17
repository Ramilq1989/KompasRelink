@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
echo Замена путей ссылок в файлах КОМПАС-3D (все подпапки)
echo.
set /p ROOT=Папка со сборками (например G:\Проекты):
set /p OLD=Старый путь (например //metall/Проекты):
set /p NEW=Новый путь (например G:/Проекты):
echo.
echo ===== ПРОБНЫЙ ПРОГОН (ничего не меняется) =====
python relink.py "%ROOT%" --map "%OLD%=%NEW%"
echo.
set /p OK=Записать изменения? (резервные копии будут сделаны) [y/N]:
if /i "%OK%"=="y" python relink.py "%ROOT%" --map "%OLD%=%NEW%" --apply
pause
