@echo off
setlocal
cd /d "%~dp0"
set "REPO=https://github.com/Kochleroguz/kochler-fiyat-otomasyonu/archive/refs/heads/main.zip"
set "WORK=%TEMP%\kochler-fiyat-guncelleme-%RANDOM%-%RANDOM%"
echo Kochler Fiyat Otomasyonu guncelleniyor...
timeout /t 3 /nobreak >nul
mkdir "%WORK%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri '%REPO%' -OutFile '%WORK%\update.zip'; Expand-Archive -Path '%WORK%\update.zip' -DestinationPath '%WORK%' -Force } catch { Write-Error $_; exit 1 }"
if errorlevel 1 goto :error
robocopy "%WORK%\kochler-fiyat-otomasyonu-main" "%~dp0" /E /NFL /NDL /NJH /NJS /NP /XD data __pycache__ /XF fiyatlar.db Guncelle.cmd >nul
if errorlevel 8 goto :error
start "Kochler Fiyat Otomasyonu" cmd /c "cd /d \"%~dp0\" && call Baslat.cmd"
exit /b 0
:error
echo Guncelleme basarisiz oldu. Internet baglantinizi kontrol edip tekrar deneyin.
pause
exit /b 1
