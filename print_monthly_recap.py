# -*- coding: utf-8 -*-
"""
In bảng tổng hợp chốt tháng từ Google Sheet
"""
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import gspread
from google.oauth2.credentials import Credentials

AUTH_USER_FILE = r'C:\Users\lap4all\Documents\Auto report\authorized_user.json'
SPREADSHEET_ID = '147nvGXc2D7UJNJGsWaFjaIZ6FkJDD7Zmevn3lBs-Bl0'

def print_recap():
    creds = Credentials.from_authorized_user_file(AUTH_USER_FILE, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet('Tổng Hợp Tháng')
    rows = ws.get_all_values()

    month_str = rows[1][1]
    print("\n" + "=" * 110)
    print(f"🏆 BẢNG TỔNG HỢP & CHỐT THU TIỀN PHẠT AM THÁNG {month_str} - VÙNG NTB")
    print("=" * 110)
    header = rows[2]
    print(f"{header[0]:<4} | {header[1]:<8} | {header[2]:<22} | {header[3]:<10} | {header[4]:<10} | {header[6]:<12} | {header[7]:<10} | {header[10]:<14} | {header[13]}")
    print("-" * 110)
    for r in rows[3:-1]:
        print(f"{r[0]:<4} | {r[1]:<8} | {r[2]:<22} | {r[3]:<10} | {r[4]:<10} | {r[6]:<12} | {r[7]:<10} | {r[10]:<14} | {r[13]}")
    print("-" * 110)
    total = rows[-1]
    print(f"     |          | {total[2]:<22} | {total[3]:<10} | {total[4]:<10} | {total[6]:<12} | {total[7]:<10} | {total[10]:<14} |")
    print("=" * 110)
    print(f"\n🔗 Link Google Sheet: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={ws.id}\n")

if __name__ == '__main__':
    print_recap()
