# -*- coding: utf-8 -*-
"""
Module: checkin_sup_bot.py
Xử lý Điểm danh & Check-in đầu ngày dành cho SUP/AM Vùng NTB qua G-Talk Webhook.
- Nhận tin nhắn / ảnh check-in TimestampCam từ Group G-Talk: 2095921878551764992
- Tự động phân tích SUP, kho/bưu cục, thời gian gửi và ghi vào Google Sheet
- Nhắc nhở lúc 07:55 (trước 8h 5 phút)
- Cảnh báo trễ & yêu cầu gửi bù lúc 08:00
- Báo cáo tổng hợp lúc 11:00
- Hỗ trợ xin nghỉ phép qua cú pháp #off hoặc tick tay trực tiếp trên Google Sheet.
"""

import os
import sys
import re
import json
import time
import unicodedata
from datetime import datetime, date, timedelta, timezone
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import gspread
from google.oauth2.credentials import Credentials as UserCredentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials

VN_TZ = timezone(timedelta(hours=7))

def get_vn_now() -> datetime:
    return datetime.now(VN_TZ).replace(tzinfo=None)

def get_vn_today() -> date:
    return get_vn_now().date()

# ─── CẤU HÌNH ───────────────────────────────────────────────
GTALK_OA_TOKEN   = "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
GTALK_CHANNEL_ID = "2095921878551764992"
GTALK_API_URL    = "https://mbff.ghn.vn/api/gtalk/send-message"

SPREADSHEET_ID   = "1sR4bqfatBj7bI2KWzwAPDWeifsfg2FICOZYauuzqBeE"
SHEET_TAB_NAME   = "Sheet1"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

AUTH_CANDIDATES = [
    os.path.join(BASE_DIR, "authorized_user.json"),
    r"C:\Users\lap4all\Documents\Auto report\authorized_user.json",
    r"C:\Users\lap4all\Desktop\auto-report\authorized_user.json",
    r"C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json",
    "authorized_user.json",
]
SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, "credentials.json"),
    r"C:\Users\lap4all\Documents\Auto report\credentials.json",
    r"C:\Users\lap4all\Desktop\auto-report\credentials.json",
    "credentials.json",
]

# ─── DANH SÁCH 3 SUP ─────────────────────────────────────────
SUPS = {
    "am_luc_nt": {
        "id": "am_luc_nt",
        "full_name": "Nguyễn Tiến Lực",
        "role": "SUP/AM",
        "provinces": ["Khánh Hòa", "Đắk Nông"],
        "primary_locations": ["Kho Trung Chuyển Khánh Hòa", "Kho Chuyển Tiếp Đắk Nông"],
        "aliases": ["nguyễn tiến lực", "nguyen tien luc", "tiến lực", "tien luc", "lực", "luc", "lucnt"],
        "location_keywords": ["khánh hòa", "khanh hoa", "ktckh", "đắk nông", "dak nong", "dno", "kh", "dn"]
    },
    "am_hoang_nm": {
        "id": "am_hoang_nm",
        "full_name": "Nguyễn Minh Hoàng",
        "role": "SUP/AM",
        "provinces": ["Lâm Đồng"],
        "primary_locations": ["Kho Chuyển Tiếp Đức Trọng-Lâm Đồng", "Kho Chuyển Tiếp Bảo Lộc-Lâm Đồng"],
        "aliases": ["nguyễn minh hoàng", "nguyen minh hoang", "minh hoàng", "minh hoang", "hoàng", "hoang", "hoangnm"],
        "location_keywords": ["đức trọng", "duc trong", "bảo lộc", "bao loc", "lâm đồng", "lam dong", "dt", "bl", "ldo"]
    },
    "am_khanh_nn": {
        "id": "am_khanh_nn",
        "full_name": "Nguyễn Ngọc Khánh",
        "role": "SUP/AM",
        "provinces": ["Bình Thuận"],
        "primary_locations": ["Kho Chuyển Tiếp Bình Thuận", "Bưu cục Bình Thuận"],
        "aliases": ["nguyễn ngọc khánh", "nguyen ngoc khanh", "ngọc khánh", "ngoc khanh", "khánh", "khanh", "khanhnn"],
        "location_keywords": ["bình thuận", "binh thuan", "kctbt", "bt", "phan thiết", "phan thiet", "la gi", "hàm thuận", "ham thuan", "hàm tân", "ham tan", "tuy phong", "bắc bình", "bac binh"]
    }
}

# ─── GOOGLE SHEETS HELPER ────────────────────────────────────

def get_gspread_client():
    env_json = os.environ.get('GOOGLE_AUTH_JSON')
    if env_json:
        try:
            info = json.loads(env_json)
            creds = UserCredentials.from_authorized_user_info(info, scopes=SCOPES)
            return gspread.authorize(creds)
        except Exception:
            pass
    for auth_file in AUTH_CANDIDATES:
        if os.path.exists(auth_file):
            try:
                creds = UserCredentials.from_authorized_user_file(auth_file, scopes=SCOPES)
                return gspread.authorize(creds)
            except Exception:
                pass
    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = ServiceAccountCredentials.from_service_account_file(cred_path, scopes=SCOPES)
                return gspread.authorize(creds)
            except Exception:
                pass
    raise PermissionError("Không thể xác thực Google Sheets!")

def get_worksheet():
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(SHEET_TAB_NAME)
    except Exception:
        ws = sh.sheet1
    return ws

def ensure_today_rows(ws, target_date_str=None):
    """Đảm bảo có đủ 3 dòng cho 3 SUP trong ngày target_date_str (DD/MM/YYYY)."""
    if not target_date_str:
        target_date_str = get_vn_today().strftime("%d/%m/%Y")

    rows = ws.get_all_values()
    if not rows:
        headers = ['Ngày', 'Tên SUP', 'Trạng thái', 'Giờ check-in', 'Kho / Bưu cục check-in', 'Ghi chú', 'Chi tiết / Ảnh', 'Thời gian cập nhật']
        ws.append_row(headers)
        rows = [headers]

    # Kiểm tra xem SUP nào đã có dòng hôm nay
    existing_sups = set()
    for idx, r in enumerate(rows[1:], 2):
        if len(r) >= 2 and r[0].strip() == target_date_str:
            name = r[1].strip()
            for s_id, s_info in SUPS.items():
                if s_info["full_name"].lower() in name.lower():
                    existing_sups.add(s_id)

    # Thêm dòng cho SUP còn thiếu
    appended = False
    for s_id, s_info in SUPS.items():
        if s_id not in existing_sups:
            ws.append_row([
                target_date_str,
                s_info["full_name"],
                "Chưa check-in",
                "",
                "",
                "",
                "",
                get_vn_now().strftime("%d/%m/%Y %H:%M:%S")
            ])
            appended = True

    if appended:
        time.sleep(1)
    return target_date_str

def get_today_records(target_date_str=None):
    """Lấy dữ liệu check-in của 3 SUP trong ngày từ Google Sheet."""
    ws = get_worksheet()
    target_date_str = ensure_today_rows(ws, target_date_str)
    rows = ws.get_all_values()

    records = {}
    for s_id, s_info in SUPS.items():
        records[s_id] = {
            "row_idx": None,
            "sup_id": s_id,
            "name": s_info["full_name"],
            "date": target_date_str,
            "status": "Chưa check-in",
            "time": "",
            "location": "",
            "note": "",
            "details": "",
            "updated_at": ""
        }

    for idx, r in enumerate(rows[1:], 2):
        if len(r) >= 2 and r[0].strip() == target_date_str:
            sup_name_in_row = r[1].strip()
            for s_id, s_info in SUPS.items():
                if s_info["full_name"].lower() in sup_name_in_row.lower():
                    records[s_id]["row_idx"] = idx
                    records[s_id]["status"] = r[2].strip() if len(r) > 2 else "Chưa check-in"
                    records[s_id]["time"] = r[3].strip() if len(r) > 3 else ""
                    records[s_id]["location"] = r[4].strip() if len(r) > 4 else ""
                    records[s_id]["note"] = r[5].strip() if len(r) > 5 else ""
                    records[s_id]["details"] = r[6].strip() if len(r) > 6 else ""
                    records[s_id]["updated_at"] = r[7].strip() if len(r) > 7 else ""

    return ws, records

def update_sup_record(ws, row_idx, status, checkin_time="", location="", note="", details=""):
    """Cập nhật dòng dữ liệu của SUP trên Google Sheet."""
    updated_at = get_vn_now().strftime("%d/%m/%Y %H:%M:%S")
    ws.update(
        values=[[status, checkin_time, location, note, details, updated_at]],
        range_name=f"C{row_idx}:H{row_idx}"
    )

# ─── GỬI TIN NHẮN GTALK ─────────────────────────────────────

def send_gtalk_message(text: str, channel_id: str = GTALK_CHANNEL_ID, oa_token: str = GTALK_OA_TOKEN):
    """Gửi tin nhắn định dạng HTML vào nhóm GTalk."""
    payload = {
        "channelId": str(channel_id),
        "clientMsgId": str(int(get_vn_now().timestamp() * 1000)),
        "content": {"parseMode": "HTML", "text": text},
        "oaToken": oa_token,
    }
    try:
        r = requests.post(GTALK_API_URL, json=payload, timeout=15, verify=False)
        res = r.json() if r.status_code == 200 else {}
        if r.status_code == 200 and res.get("errorCode") == "success":
            return True, "OK"
        return False, f"HTTP {r.status_code} - {r.text[:200]}"
    except Exception as e:
        return False, str(e)

# ─── NHẬN DIỆN SUP & NỘI DUNG CHECK-IN ───────────────────────

def remove_accents(input_str: str) -> str:
    if not input_str:
        return ""
    nfkd = unicodedata.normalize('NFKD', input_str)
    return ''.join([c for c in nfkd if not unicodedata.combining(c)]).lower()

def detect_target_sup(sender_name: str, text: str):
    """Nhận diện SUP nào dựa trên tên người gửi và nội dung tin nhắn."""
    norm_sender = remove_accents(sender_name)
    norm_text = remove_accents(text)

    # 1. So khớp người gửi trước (Display Name trên GTalk)
    for s_id, s_info in SUPS.items():
        for alias in s_info["aliases"]:
            if remove_accents(alias) in norm_sender:
                return s_id, s_info

    # 2. So khớp từ khóa tên SUP trong text (ví dụ: 'SUP Khánh gửi checkin...', 'Khánh checkin')
    for s_id, s_info in SUPS.items():
        for alias in s_info["aliases"]:
            if re.search(r'\b' + re.escape(remove_accents(alias)) + r'\b', norm_text):
                return s_id, s_info

    # 3. So khớp từ khóa kho đặc trưng của SUP
    for s_id, s_info in SUPS.items():
        for loc_kw in s_info["location_keywords"]:
            if re.search(r'\b' + re.escape(remove_accents(loc_kw)) + r'\b', norm_text):
                return s_id, s_info

    return None, None

def detect_location(text: str, sup_info: dict) -> str:
    """Trích xuất kho hoặc bưu cục từ tin nhắn check-in."""
    norm_txt = remove_accents(text)

    # Loại bỏ các từ khóa cú pháp
    clean = re.sub(r'(?i)#?checkin|#?check\s*in|diem\s*danh|gui\s*checkin', '', text).strip()
    clean = re.sub(r'^[\s\-\:\,\.]+', '', clean).strip()

    # Nhận diện theo các bưu cục / kho đặc thù
    if "phan thiet" in norm_txt:
        return "BC Phan Thiết"
    if "la gi" in norm_txt:
        return "BC La Gi"
    if "ham thuan bac" in norm_txt:
        return "BC Hàm Thuận Bắc"
    if "ham thuan nam" in norm_txt:
        return "BC Hàm Thuận Nam"
    if "ham tan" in norm_txt:
        return "BC Hàm Tân"
    if "duc trong" in norm_txt:
        return "Kho Chuyển Tiếp Đức Trọng"
    if "bao loc" in norm_txt:
        return "Kho Chuyển Tiếp Bảo Lộc"
    if "khanh hoa" in norm_txt or "ktckh" in norm_txt:
        return "Kho Trung Chuyển Khánh Hòa"
    if "dak nong" in norm_txt or "dno" in norm_txt:
        return "Kho Chuyển Tiếp Đắk Nông"
    if "binh thuan" in norm_txt or "kctbt" in norm_txt:
        return "Kho Chuyển Tiếp Bình Thuận"

    # Nếu clean còn text cụ thể (ví dụ: 'KCT Bình Thuận', 'BC Phan Rí')
    if len(clean) > 2 and not any(kw in clean.lower() for kw in ["sup", "chup", "anh", "hinh", "camera"]):
        return clean

    # Mặc định lấy kho chính đầu tiên của SUP
    return sup_info["primary_locations"][0]

# ─── XỬ LÝ WEBHOOK TỪ GTALK ──────────────────────────────────

def handle_sup_webhook(data: dict):
    """
    Xử lý payload khi có tin nhắn/ảnh gửi vào Group SUP (2095921878551764992).
    """
    now = get_vn_now()
    now_hm = now.strftime("%H:%M:%S")
    today_display = now.strftime("%d/%m/%Y")

    sender_obj = data.get("sender") or {}
    sender_name = sender_obj.get("displayName") or sender_obj.get("name") or data.get("senderName") or ""

    # Trích xuất text
    msg_obj = data.get("message") or {}
    content_obj = data.get("content") or msg_obj.get("content") or {}
    text = ""
    if isinstance(content_obj, dict):
        text = content_obj.get("text", "")
    elif isinstance(content_obj, str):
        text = content_obj

    if not text:
        text = str(data.get("text") or msg_obj.get("text") or "")

    # Kiểm tra có đính kèm ảnh không
    content_type = str(data.get("contentType") or msg_obj.get("contentType") or "").lower()
    has_photo = bool("image" in content_type or "photo" in content_type or "attachment" in str(data))

    norm_txt = remove_accents(text)

    # Bỏ qua nếu là tin do chính bot gửi
    if any(kw in text for kw in ["XÁC NHẬN CHECK-IN", "NHẮC NHỞ", "CẢNH BÁO 08:00", "BÁO CÁO CHECK-IN SUP"]):
        return {"status": "skipped", "reason": "bot message"}

    # ── 1. KIỂM TRA TIN BÁO NGHỈ / CÔNG TÁC (#off, xin nghỉ, nghỉ phép) ──
    is_excuse = bool(re.search(r'(?i)#?off\b|xin\s*nghi|nghi\s*phep|bao\s*vang|cong\s*tac|di\s*hop', text))
    if is_excuse:
        sup_id, sup_info = detect_target_sup(sender_name, text)
        if not sup_id:
            return {"status": "ignored", "reason": "cannot identify sup for excuse"}

        # Trích xuất lý do
        reason = re.sub(r'(?i)#?off[:\s]*|xin\s*nghi[:\s]*|nghi\s*phep[:\s]*', '', text).strip()
        if not reason:
            reason = "Nghỉ phép / Công tác"

        try:
            ws, records = get_today_records(today_display)
            row_idx = records[sup_id]["row_idx"]
            if row_idx:
                update_sup_record(ws, row_idx, status="🏖️ Nghỉ phép", note=reason, details=f"Báo lúc {now_hm}")

            reply = (
                f"📝 <b>XÁC NHẬN GHI NHẬN NGHỈ PHÉP</b>\n"
                f"👤 SUP: <b>{sup_info['full_name']}</b>\n"
                f"📌 Trạng thái: <b>🏖️ Nghỉ phép / Vắng</b>\n"
                f"💬 Ghi chú: <i>{reason}</i>\n"
                f"⏰ Thời gian ghi nhận: {now_hm} - {today_display}"
            )
            send_gtalk_message(reply)
            return {"status": "recorded_excuse", "sup": sup_info["full_name"], "reason": reason}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    # ── 2. KIỂM TRA TIN CHECK-IN (Có từ khóa checkin hoặc gửi kèm ảnh) ──
    is_checkin = bool(re.search(r'(?i)#?check\s*in|diem\s*danh|kct|ktc|buu\s*cuc', text) or has_photo)
    if not is_checkin:
        return {"status": "ignored", "reason": "not checkin message"}

    sup_id, sup_info = detect_target_sup(sender_name, text)
    if not sup_id:
        return {"status": "ignored", "reason": "cannot identify sup for checkin"}

    location = detect_location(text, sup_info)

    # Đánh giá đúng giờ hay trễ (Cut-off là 08:00:59)
    cutoff_time = now.replace(hour=8, minute=0, second=59, microsecond=0)
    is_on_time = (now <= cutoff_time)
    status_label = "✅ Đúng giờ" if is_on_time else "⚠️ Gửi bù (Trễ)"
    photo_label = "📷 Có ảnh TimestampCam" if has_photo else "📝 Gửi text"

    try:
        ws, records = get_today_records(today_display)
        row_idx = records[sup_id]["row_idx"]
        if row_idx:
            update_sup_record(
                ws,
                row_idx,
                status=status_label,
                checkin_time=now_hm,
                location=location,
                note="Đúng giờ" if is_on_time else "Gửi bù sau 08:00",
                details=photo_label
            )

        reply = (
            f"✅ <b>XÁC NHẬN CHECK-IN ĐẦU NGÀY</b>\n"
            f"👤 SUP: <b>{sup_info['full_name']}</b>\n"
            f"📍 Địa điểm: <b>{location}</b>\n"
            f"⏰ Thời gian: <b>{now_hm}</b> ({status_label})\n"
            f"📌 Ghi nhận: {photo_label}\n"
            f"🔗 <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Xem bảng theo dõi Check-in</a>"
        )
        send_gtalk_message(reply)
        return {"status": "recorded_checkin", "sup": sup_info["full_name"], "status_label": status_label}
    except Exception as e:
        return {"status": "error", "message": str(e)}

# ─── KỊCH BẢN HẸN GIỜ (07:55, 08:00, 11:00) ──────────────────

last_job_executed = {}

def job_remind_0755():
    """07:55 - Nhắc nhở các SUP chưa check-in (còn 5 phút trước hạn 08:00)."""
    now = get_vn_now()
    today_display = now.strftime("%d/%m/%Y")
    try:
        ws, records = get_today_records(today_display)
        missing_sups = []
        for s_id, rec in records.items():
            status = rec.get("status", "")
            # Nếu chưa check-in và không phải nghỉ phép
            if "Chưa check-in" in status and not any(kw in status for kw in ["Nghỉ", "Vắng", "Đúng giờ", "Gửi bù"]):
                missing_sups.append(rec["name"])

        if not missing_sups:
            print(f"[{now.strftime('%H:%M:%S')}] 07:55 - Tất cả SUP đã check-in hoặc xin nghỉ. Không cần nhắc.")
            return

        lines = [
            "⏰ <b>[NHẮC NHỞ CHECK-IN 07:55]</b>",
            "<i>Còn 5 phút trước giờ chốt điểm danh đầu ca sáng (08:00)!</i>",
            "",
            "📌 <b>Danh sách SUP chưa gửi check-in TimestampCam:</b>"
        ]
        for idx, name in enumerate(missing_sups, 1):
            lines.append(f"{idx}. <b>{name}</b>")

        lines.extend([
            "",
            "👉 <i>Yêu cầu các SUP chụp ảnh TimestampCam kèm vị trí kho/bưu cục gửi vào nhóm ngay nhé!</i>",
            f"🔗 <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Bảng theo dõi Check-in</a>"
        ])
        send_gtalk_message("\n".join(lines))
        print(f"[{now.strftime('%H:%M:%S')}] Đã gửi nhắc nhở 07:55 cho {len(missing_sups)} SUP.")
    except Exception as e:
        print(f"[ERROR job_remind_0755]: {e}")

def job_cutoff_0800():
    """08:00 - Cảnh báo quá hạn và yêu cầu các SUP chưa check-in gửi bù ngay."""
    now = get_vn_now()
    today_display = now.strftime("%d/%m/%Y")
    try:
        ws, records = get_today_records(today_display)
        overdue_sups = []
        for s_id, rec in records.items():
            status = rec.get("status", "")
            if "Chưa check-in" in status:
                overdue_sups.append(rec["name"])
                # Cập nhật thành Quá hạn (Chờ bù)
                if rec["row_idx"]:
                    update_sup_record(ws, rec["row_idx"], status="❌ Quá hạn (Chờ bù)", note="Quá 08:00 chưa gửi")

        if not overdue_sups:
            print(f"[{now.strftime('%H:%M:%S')}] 08:00 - Tất cả SUP đã hoàn thành check-in đúng giờ hoặc xin phép.")
            return

        lines = [
            "⚠️ <b>[CẢNH BÁO QUÁ HẠN CHECK-IN 08:00]</b>",
            "<i>Đã quá 08:00 sáng - Hệ thống đã chốt danh sách check-in đúng giờ!</i>",
            "",
            "❌ <b>Các SUP sau chưa check-in đúng hạn:</b>"
        ]
        for idx, name in enumerate(overdue_sups, 1):
            lines.append(f"{idx}. <b>{name}</b>")

        lines.extend([
            "",
            "👉 <b>Yêu cầu các SUP khẩn trương chụp ảnh TimestampCam gửi bù ngay vào nhóm!</b>",
            f"🔗 <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Bảng theo dõi Check-in</a>"
        ])
        send_gtalk_message("\n".join(lines))
        print(f"[{now.strftime('%H:%M:%S')}] Đã gửi cảnh báo quá hạn 08:00 cho {len(overdue_sups)} SUP.")
    except Exception as e:
        print(f"[ERROR job_cutoff_0800]: {e}")

def job_recap_1100():
    """11:00 - Báo cáo tổng hợp tình hình điểm danh ca sáng của 3 SUP."""
    now = get_vn_now()
    today_display = now.strftime("%d/%m/%Y")
    try:
        ws, records = get_today_records(today_display)
        lines = [
            f"📊 <b>BÁO CÁO CHECK-IN SUP ĐẦU NGÀY - {today_display} (11:00)</b>",
            "─────────────────────────────"
        ]

        for idx, (s_id, rec) in enumerate(records.items(), 1):
            status = rec.get("status", "Chưa check-in")
            time_str = rec.get("time", "")
            loc_str = rec.get("location", "")
            note_str = rec.get("note", "")

            icon = "✅"
            if "Nghỉ" in status or "Vắng" in status:
                icon = "🏖️"
            elif "Trễ" in status or "bù" in status:
                icon = "⚠️"
            elif "Chưa" in status or "Quá hạn" in status:
                icon = "❌"

            detail_parts = []
            if time_str:
                detail_parts.append(time_str)
            if loc_str:
                detail_parts.append(loc_str)
            if note_str:
                detail_parts.append(f"<i>({note_str})</i>")

            detail_text = f" - {' | '.join(detail_parts)}" if detail_parts else ""
            lines.append(f"{idx}. <b>{rec['name']}:</b> {icon} {status}{detail_text}")

        lines.extend([
            "─────────────────────────────",
            f"🔗 <b>Chi tiết:</b> <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Google Sheet Check-in SUP</a>"
        ])
        send_gtalk_message("\n".join(lines))
        print(f"[{now.strftime('%H:%M:%S')}] Đã gửi báo cáo tổng hợp 11:00.")
    except Exception as e:
        print(f"[ERROR job_recap_1100]: {e}")

def check_sup_schedule():
    """Hàm được gọi định kỳ trong vòng lặp scheduler của Webhook Server."""
    now = get_vn_now()
    today_str = now.strftime("%Y-%m-%d")
    hm = now.strftime("%H:%M")

    schedule_jobs = {
        "07:55": ("job_0755", job_remind_0755),
        "08:00": ("job_0800", job_cutoff_0800),
        "11:00": ("job_1100", job_recap_1100)
    }

    if hm in schedule_jobs:
        job_key_name, job_func = schedule_jobs[hm]
        trigger_key = f"{today_str}_{job_key_name}"
        if last_job_executed.get(trigger_key):
            return
        last_job_executed[trigger_key] = True
        print(f"⏰ [SUP SCHEDULER] Kích hoạt {job_key_name} lúc {hm}...")
        try:
            job_func()
        except Exception as e:
            print(f"❌ Lỗi chạy {job_key_name}: {e}")
