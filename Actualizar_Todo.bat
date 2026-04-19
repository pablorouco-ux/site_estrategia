@echo off
cd /d "%~dp0"
echo ==========================================
echo   Sincronizando Showroom (Excel + Fotos)
echo ==========================================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\actualizar_site.ps1"
echo.
echo ==========================================
echo   Proceso terminado.
echo ==========================================
pause