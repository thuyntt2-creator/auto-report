# -*- coding: utf-8 -*-
"""
Module Đồng Bộ Dữ Liệu Điểm Danh & Thu Phạt Lên Google Sheet (Vùng NTB) - Batch Mode Tốc Độ Cao
Hiển thị tách biệt rõ ràng đơn giá phạt:
- Phạt Nộp Trễ (50k/lần)
- Phạt Chưa Nộp (100k/lần)
- Phạt Sai Định Dạng (200k/lần)
"""

import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import json
import sqlite3
from datetime import datetime, date
import gspread
from google.oauth2.credentials import Credentials

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "am_config.json")
DB_PATH = os.path.join(BASE_DIR, "diem_danh.db")
AUTH_USER_FILE = r'C:\Users\lap4all\Documents\Auto report\authorized_user.json'
SPREADSHEET_ID = '147nvGXc2D7UJNJGsWaFjaIZ6FkJDD7Zmevn3lBs-Bl0'
SHEET_TITLE = 'Theo Dõi Phạt AM'

HEADERS = [
    "Ngày", 
    "Mã NV", 
    "Tên AM", 
    "Mốc 1: Đầu ngày (08:00)", 
    "Mốc 2: TTS Ca 1 (11:00)", 
    "Mốc 3: TTS Ca 2 (16:00)", 
    "Mốc 4: LTC TTS (20:00)", 
    "Mốc 5: Điểm nóng (10:00)", 
    "Số Lần Trễ",
    "Số Lần Chưa Nộp",
    "Auto Xin Phép Trễ",
    "Tiền Phạt Trễ (x50k)",
    "Tiền Chưa Nộp (x100k)",
    "Tổng Phạt Gốc (VNĐ)", 
    "Admin Loại Trừ / Nghỉ Phép", 
    "Thực Thu (VNĐ)", 
    "Trạng Thái Thu", 
    "Ghi Chú Lý Do"
]

def get_sheet_client():
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    env_json = os.environ.get('GOOGLE_AUTH_JSON')
    if env_json:
        info = json.loads(env_json)
        creds = Credentials.from_authorized_user_info(info, scopes=scopes)
    elif os.path.exists(os.path.join(BASE_DIR, 'authorized_user.json')):
        creds = Credentials.from_authorized_user_file(os.path.join(BASE_DIR, 'authorized_user.json'), scopes=scopes)
    else:
        creds = Credentials.from_authorized_user_file(AUTH_USER_FILE, scopes=scopes)
    return gspread.authorize(creds)

def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    return conn

def sync_daily_to_sheet(target_date: date = None):
    """
    Đồng bộ dữ liệu điểm danh và tiền phạt của ngày target_date lên Google Sheet bằng 1 batch call siêu nhanh.
    """
    if target_date is None:
        target_date = date.today()

    date_str = target_date.strftime("%Y-%m-%d")
    date_display = target_date.strftime("%d/%m/%Y")
    
    config = load_config()
    active_ams = [am for am in config["ams"] if am.get("is_active", True)]
    
    # Lấy dữ liệu điểm danh từ DB
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT am_id, milestone_id, status, late_minutes, penalty_amount,
               strftime('%H:%M', submitted_at) as submit_time
        FROM attendance_records
        WHERE date = ?
        """, (date_str,))
        records = {}
        for row in cur.fetchall():
            records[(row["am_id"], row["milestone_id"])] = dict(row)

        cur.execute("""
        SELECT am_id, milestone_id, reason, raw_text, created_at
        FROM excuses
        WHERE date = ?
        """, (date_str,))
        excuses_by_am = {}
        for row in cur.fetchall():
            aid = row["am_id"]
            if aid not in excuses_by_am:
                excuses_by_am[aid] = []
            excuses_by_am[aid].append(dict(row))

    gc = get_sheet_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(SHEET_TITLE)
    except Exception:
        ws = sh.add_worksheet(title=SHEET_TITLE, rows=1000, cols=20)

    all_values = ws.get_all_values()
    sheet_id = ws.id

    # Đảm bảo Header chuẩn 18 cột
    if not all_values or len(all_values[0]) < 18 or all_values[0][11] != "Tiền Phạt Trễ (x50k)":
        ws.clear()
        ws.update(range_name='A1:R1', values=[HEADERS])
        all_values = [HEADERS]

    existing_rows = {}
    for idx, row in enumerate(all_values[1:], start=2):
        if len(row) >= 2:
            r_date = row[0].strip()
            r_emp_id = row[1].strip()
            existing_rows[(r_date, r_emp_id)] = idx

    try:
        from diem_danh_bot import get_m5_required_ams
        m5_required_map = get_m5_required_ams()
    except Exception:
        m5_required_map = {}

    batch_updates = []
    rows_to_append = []

    for idx, am in enumerate(active_ams, 1):
        am_id = am["id"]
        emp_id = str(am.get("employee_id", ""))
        am_name = am["full_name"]

        m_texts = []
        count_late = 0
        count_missing = 0
        fine_late = 0
        fine_missing = 0

        # Kiểm tra xin phép
        am_excuses = excuses_by_am.get(am_id, [])
        if am_excuses:
            reasons = [e['reason'] for e in am_excuses if e.get('reason')]
            if reasons:
                auto_excuse_text = f"📝 Có xin ({', '.join(reasons)})"
            else:
                auto_excuse_text = "📝 Có xin phép"
        else:
            auto_excuse_text = "-"

        for m_id in (1, 2, 3, 4):
            rec = records.get((am_id, m_id))
            excuse_m = next((e for e in am_excuses if e.get("milestone_id") == m_id or e.get("milestone_id") is None), None)

            if not rec:
                if excuse_m:
                    m_texts.append("⏳ Có xin phép")
                else:
                    fine_missing += config["fines"]["not_submitted"]
                    count_missing += 1
                    m_texts.append("❌ Chưa nộp (100k)")
            elif rec["status"] == "ON_TIME":
                m_texts.append(f"✅ {rec['submit_time']}")
            elif rec["status"] == "LATE":
                if excuse_m or am_excuses:
                    # Đã có xin phép trễ -> TỰ ĐỘNG MIỄN PHẠT 50K (Tính 0đ)!
                    m_texts.append(f"⚠️ Trễ {rec['late_minutes']}p (Đã xin - 0đ)")
                else:
                    fine_late += rec.get("penalty_amount", 50000)
                    count_late += 1
                    m_texts.append(f"⚠️ Trễ {rec['late_minutes']}p ({rec['submit_time']})")
            elif rec["status"] == "INVALID":
                fine_late += config["fines"]["invalid"]
                m_texts.append(f"🚫 Sai định dạng ({rec['submit_time']})")
            else:
                m_texts.append(f"{rec['status']}")

        # Mốc 5
        rec_m5 = records.get((am_id, 5))
        if rec_m5:
            if rec_m5["status"] == "LATE":
                if am_excuses:
                    m_texts.append(f"⚠️ Trễ (Đã xin - 0đ)")
                else:
                    fine_late += rec_m5.get("penalty_amount", 50000)
                    count_late += 1
                    m_texts.append(f"⚠️ Trễ ({rec_m5['submit_time']})")
            elif rec_m5["status"] == "INVALID":
                fine_late += config["fines"]["invalid"]
                m_texts.append("🚫 Sai định dạng")
            else:
                m_texts.append(f"✅ {rec_m5['submit_time']}")
        else:
            if am_id in m5_required_map:
                count_missing += 1
                fine_missing += config["fines"]["not_submitted"]
                m_texts.append("❌ Chưa nộp (100k)")
            else:
                m_texts.append("— (Không có BC <50%)")

        am_fine_total = fine_late + fine_missing

        row_key = (date_display, emp_id)
        if row_key in existing_rows:
            row_idx = existing_rows[row_key]
            batch_updates.append({
                'range': f'A{row_idx}:N{row_idx}',
                'values': [[
                    date_display, emp_id, am_name,
                    m_texts[0], m_texts[1], m_texts[2], m_texts[3], m_texts[4],
                    count_late, count_missing, auto_excuse_text,
                    fine_late, fine_missing, am_fine_total
                ]]
            })
            batch_updates.append({
                'range': f'P{row_idx}',
                'values': [[f'=IF(O{row_idx}=TRUE; 0; N{row_idx})']]
            })
        else:
            target_row_num = len(all_values) + len(rows_to_append) + 1
            formula_p = f'=IF(O{target_row_num}=TRUE; 0; N{target_row_num})'
            new_row = [
                date_display, emp_id, am_name,
                m_texts[0], m_texts[1], m_texts[2], m_texts[3], m_texts[4],
                count_late, count_missing, auto_excuse_text,
                fine_late, fine_missing, am_fine_total,
                False, # Checkbox Cột O
                formula_p, # Công thức Cột P
                "Chưa thu" if am_fine_total > 0 else "Miễn phạt", # Cột Q
                "" # Ghi chú Cột R
            ]
            rows_to_append.append(new_row)

    if batch_updates:
        ws.batch_update(batch_updates, value_input_option='USER_ENTERED')

    if rows_to_append:
        start_row = len(all_values) + 1
        end_row = start_row + len(rows_to_append) - 1
        ws.append_rows(rows_to_append, value_input_option='USER_ENTERED')
        
        # Checkbox cột O (index 14, startColIndex 14, endColIndex 15)
        rule_checkbox = {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row - 1,
                    "endRowIndex": end_row,
                    "startColumnIndex": 14,
                    "endColumnIndex": 15
                },
                "rule": {
                    "condition": {"type": "BOOLEAN"},
                    "showCustomUi": True
                }
            }
        }
        # Dropdown cột Q (index 16, startColIndex 16, endColIndex 17)
        rule_dropdown = {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row - 1,
                    "endRowIndex": end_row,
                    "startColumnIndex": 16,
                    "endColumnIndex": 17
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [
                            {"userEnteredValue": "Chưa thu"},
                            {"userEnteredValue": "Đã thu"},
                            {"userEnteredValue": "Miễn phạt"}
                        ]
                    },
                    "showCustomUi": True
                }
            }
        }
        sh.batch_update({"requests": [rule_checkbox, rule_dropdown]})

    total_rows = max(len(all_values), len(existing_rows) + 1)
    if total_rows > 1:
        ws.format("A1:R1", {
            "backgroundColor": {"red": 0.08, "green": 0.28, "blue": 0.45},
            "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True, "fontSize": 10},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "wrapStrategy": "WRAP"
        })
        ws.format(f"A2:B{total_rows}", {"horizontalAlignment": "CENTER"})
        ws.format(f"D2:H{total_rows}", {"horizontalAlignment": "CENTER"})
        ws.format(f"I2:J{total_rows}", {
            "numberFormat": {"type": "NUMBER", "pattern": "0"},
            "horizontalAlignment": "CENTER", "textFormat": {"bold": True}
        })
        ws.format(f"K2:K{total_rows}", {"horizontalAlignment": "CENTER"})
        ws.format(f"L2:N{total_rows}", {"numberFormat": {"type": "CURRENCY", "pattern": "#,##0\" đ\""}, "horizontalAlignment": "RIGHT"})
        ws.format(f"O2:O{total_rows}", {"horizontalAlignment": "CENTER"})
        ws.format(f"P2:P{total_rows}", {"numberFormat": {"type": "CURRENCY", "pattern": "#,##0\" đ\""}, "horizontalAlignment": "RIGHT", "textFormat": {"bold": True}})
        ws.format(f"Q2:Q{total_rows}", {"horizontalAlignment": "CENTER"})

    print(f"✅ Đã đồng bộ batch thành công ngày {date_display} lên Google Sheet '{SHEET_TITLE}'!")


def get_admin_excused_ams(target_date: date = None):
    if target_date is None:
        target_date = date.today()
    date_display = target_date.strftime("%d/%m/%Y")

    try:
        gc = get_sheet_client()
        sh = gc.open_by_key(SPREADSHEET_ID)
        ws = sh.worksheet(SHEET_TITLE)
        rows = ws.get_all_values()
    except Exception as e:
        print(f"Không thể đọc Google Sheet để lấy danh sách miễn phạt: {e}")
        return set()

    excused_emp_ids = set()
    for row in rows[1:]:
        if len(row) >= 15:
            r_date = row[0].strip()
            r_emp_id = row[1].strip()
            # Cột O là index 14
            is_checked = str(row[14]).strip().upper() == "TRUE"
            if r_date == date_display and is_checked:
                excused_emp_ids.add(r_emp_id)

    return excused_emp_ids

if __name__ == '__main__':
    sync_daily_to_sheet()
