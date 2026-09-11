@echo off
chcp 65001 >nul
title BOT DIEM DANH & TINH PHAT CUT-OFF AM (VUNG NTB)
cd /d "%~dp0"
echo.
echo ============================================================
echo   🚀 BOT ĐIỂM DANH & TÍNH PHẠT CUT-OFF BÁO CÁO AM - VÙNG NTB
echo ============================================================
echo.
echo Đang khởi động Bot Server và Bộ hẹn giờ Scheduler...
echo.
"C:\Users\lap4all\AppData\Local\Python\bin\python.exe" diem_danh_bot.py
if %ERRORLEVEL% NEQ 0 (
    python diem_danh_bot.py
)
echo.
echo Bot đã dừng. Nhấn phím bất kỳ để thoát...
pause >nul
