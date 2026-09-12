@echo off
chcp 65001 >nul
title BAN BANG CHOT PHAT DAILY N+1 DANG ANH LEN GTALK
cd /d "%~dp0"
echo.
echo ============================================================
echo   📊 BẮN BẢNG TỔNG KẾT PHẠT DAILY N+1 (DẠNG ẢNH) LÊN GTALK
echo ============================================================
echo.
python render_daily_recap_card.py
if %ERRORLEVEL% NEQ 0 (
    "C:\Users\lap4all\AppData\Local\Python\pythoncore-3.14-64\python.exe" render_daily_recap_card.py
)
echo.
echo Hoàn tất. Nhấn phím bất kỳ để thoát...
pause >nul
