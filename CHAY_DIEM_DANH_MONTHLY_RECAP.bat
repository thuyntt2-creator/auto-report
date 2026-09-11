@echo off
chcp 65001 > nul
title CHỐT SỔ ĐIỂM DANH & THU PHẠT THÁNG - VÙNG NTB
echo =================================================================
echo 📊 ĐANG TẠO BÁO CÁO CHỐT THU PHẠT THÁNG (VÙNG NTB)...
echo =================================================================
python print_monthly_recap.py
echo.
pause
