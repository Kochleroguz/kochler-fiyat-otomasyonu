@echo off
cd /d "%~dp0"
title Kochler Fiyat Otomasyonu v0.6
echo Kochler Fiyat Otomasyonu v0.6 baslatiliyor...
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 app.py
) else (
  python app.py
)
pause
