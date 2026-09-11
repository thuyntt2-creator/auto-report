@echo off
chcp 65001 >nul
title BAN BANG CHOT PHAT DAILY N+1 LEN GTALK
cd /d "%~dp0"
echo.
echo ============================================================
echo   📊 BẮN BẢNG TỔNG KẾT PHẠT DAILY N+1 LÊN GTALK TEST GROUP
echo ============================================================
echo.
"C:\Users\lap4all\AppData\Local\Python\bin\python.exe" diem_danh_bot.py --daily
if %ERRORLEVEL% NEQ 0 (
    python diem_danh_bot.py --daily
)
echo.
echo Hoàn tất. Nhấn phím bất kỳ để thoát...
pause >nul
