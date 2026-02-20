@echo off
echo ========================================
echo Starting the build process...
echo ========================================

:: 1. Удаляем старые папки сборки, чтобы не было конфликтов
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

:: 2. Запускаем сборку (режим папки, без консоли, с иконкой)
:: Мы используем GUI.pyw как точку входа
pyinstaller --noconfirm --onedir --windowed ^
 --icon="app_icon.ico" ^
 --name "Excel-Translator" ^
 --add-data "app_icon.ico;." ^
 GUI.pyw

echo.
echo ========================================
echo Build Finished! Check the "dist" folder.
echo ========================================
pause