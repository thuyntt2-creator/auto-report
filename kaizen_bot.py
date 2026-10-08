# -*- coding: utf-8 -*-
"""
Module: kaizen_bot.py
Hệ thống Bot Điểm danh KAIZEN tự động dành cho Vùng NTB qua G-Talk Webhook & Google Sheets.

Quy định vận hành KAIZEN:
- Tần suất điểm danh: Thứ 2, Thứ 4, Thứ 6 hàng tuần (hỗ trợ bật test mọi ngày).
- Ca đầu ngày (Sáng): Báo cáo trước 10:00 (Nhắc nhở lúc 09:55, Tổng kết lúc 10:05).
- Ca cuối ngày (Tối): Báo cáo trước 22:00 (Nhắc nhở lúc 21:55, Chốt sổ lúc 22:05).
- Yêu cầu ảnh: Chụp app TimestampCam đủ 3 ảnh (Tổng quan BC, Nhà VS, Mặt tiền BC).
  Thiếu 3 ảnh tính là báo cáo thiếu. Gửi sau deadline tính là báo trễ.
- Cú pháp báo cáo nhanh: ID bưu cục (warehouse_id) HOẶC Tên bưu cục kèm ảnh.
"""

import os
import sys
import re
import json
import time
import unicodedata
import threading
import argparse
from datetime import datetime, date, timedelta, timezone
from collections import defaultdict
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import gspread
from google.oauth2.credentials import Credentials as UserCredentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials

# ─── CẤU HÌNH HỆ THỐNG ───────────────────────────────────────
VN_TZ = timezone(timedelta(hours=7))

def get_vn_now() -> datetime:
    return datetime.now(VN_TZ).replace(tzinfo=None)

def get_vn_today() -> date:
    return get_vn_now().date()

# G-Talk Bot Configuration
GTALK_API_URL  = "https://mbff.ghn.vn/api/gtalk/send-message"
GTALK_OA_TOKEN = "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"

# Group ID test (hiện tại test cả 5 tỉnh vào group này)
TEST_GROUP_ID = "2077278419534073856"

# Cấu hình mapping 5 tỉnh tới Group ID chính thức:
PROVINCE_GROUPS = {
    "Ninh Thuận": "2102572223276539904",
    "Bình Thuận": "2102572259390967808",
    "Đắk Nông":   "2102572112975503360",
    "Lâm Đồng":   "2102572168718503936",
    "Khánh Hòa":  "2101117815083507712",
}

# Google Sheets Configuration
SPREADSHEET_ID  = "1nIHWPNn75j8x6-SGzCR9i_e8CPyrWHy65J-FDV0w-14"
COCAU_TAB_NAME  = "Cơ cấu"
LOG_TAB_NAME    = "Điểm danh KAIZEN"

# Danh sách bưu cục/kho tạm thời MIỄN BÁO CÁO KAIZEN
EXCLUDED_WAREHOUSE_IDS = {
    "23098000": "Dùng chung mặt bằng với Kho Trung Chuyển Khánh Hòa (1909)",
}

# Lịch trình & Ca làm việc
# 0 = Thứ 2, 2 = Thứ 4, 4 = Thứ 6 hàng tuần
ACTIVE_WEEKDAYS = [0, 2, 4]
TEST_EVERYDAY   = False  # Chỉ kích hoạt tự động vào đúng Thứ 2, Thứ 4, Thứ 6 hàng tuần

# Giờ mốc quy định
MORNING_CUTOFF_HOUR = 10
MORNING_CUTOFF_MIN  = 0
EVENING_CUTOFF_HOUR = 22
EVENING_CUTOFF_MIN  = 0

# Auth credentials candidates
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
AUTH_CANDIDATES = [
    os.path.join(BASE_DIR, "authorized_user.json"),
    r"C:\Users\lap4all\Documents\Auto report\authorized_user.json",
    r"C:\Users\lap4all\Desktop\auto-report\authorized_user.json",
    "authorized_user.json",
]
SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, "credentials.json"),
    r"C:\Users\lap4all\Documents\Auto report\credentials.json",
    r"C:\Users\lap4all\Desktop\auto-report\credentials.json",
    "credentials.json",
]

# ─── GOOGLE SHEETS CLIENT ────────────────────────────────────
def get_gspread_client():
    env_json = os.environ.get("GOOGLE_AUTH_JSON")
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
    raise PermissionError("Không thể xác thực Google Sheets từ authorized_user.json hoặc credentials.json!")

def get_spreadsheet():
    gc = get_gspread_client()
    return gc.open_by_key(SPREADSHEET_ID)

# ─── TIỆN ÍCH CHUẨN HÓA CHỮ TIẾNG VIỆT ─────────────────────
def remove_accents(input_str: str) -> str:
    if not input_str:
        return ""
    input_str = input_str.replace('đ', 'd').replace('Đ', 'd')
    nfkd = unicodedata.normalize('NFKD', input_str)
    return ''.join([c for c in nfkd if not unicodedata.combining(c)]).lower()

def clean_hub_name(name: str) -> str:
    """Loại bỏ các tiền tố mã tỉnh như (NTH), (DNO), (KHA), Kho, Bưu cục để so khớp linh hoạt."""
    cleaned = re.sub(r'\([A-Z]{2,4}\)', '', name, flags=re.IGNORECASE)
    cleaned = re.sub(r'\b(kho trung chuyen|kho chuyen tiep|kho|buu cuc|bc|kct|ktc)\b', '', cleaned, flags=re.IGNORECASE)
    return cleaned.strip()

# ─── MASTER DATA CƠ CẤU ──────────────────────────────────────
_CACHED_COCAU = None
_CACHED_TIME = 0

def load_cocau_data(force_reload=False):
    """
    Tải danh sách 99 bưu cục & kho từ tab 'Cơ cấu'.
    Cột A: warehouse_id
    Cột B: Bưu cục
    Cột C: Tỉnh
    Cột D: Am
    """
    global _CACHED_COCAU, _CACHED_TIME
    now_ts = time.time()
    if not force_reload and _CACHED_COCAU and (now_ts - _CACHED_TIME < 300):
        return _CACHED_COCAU

    sh = get_spreadsheet()
    ws = sh.worksheet(COCAU_TAB_NAME)
    rows = ws.get_all_values()

    hubs = []
    for idx, r in enumerate(rows[1:], 2):
        if len(r) >= 4 and r[0].strip():
            wid = str(r[0]).strip()
            name = str(r[1]).strip()
            prov = str(r[2]).strip()
            am = str(r[3]).strip()

            # Bỏ qua các kho tạm thời miễn báo cáo (ví dụ: CK Diên Điền dùng chung KTC Khánh Hòa)
            if wid in EXCLUDED_WAREHOUSE_IDS:
                continue
            
            # Chuẩn hóa trường hợp dòng đặc thù Nhân Cơ 1
            if "(DNO)" in name and (prov == "(DNO) Nhân Cơ 1" or not prov):
                prov = "Đắk Nông"

            cleaned = clean_hub_name(name)
            norm_c = remove_accents(cleaned)
            norm_full = remove_accents(name)

            hubs.append({
                "row_in_cocau": idx,
                "warehouse_id": wid,
                "hub_name": name,
                "clean_name": cleaned,
                "norm_name": norm_c,
                "norm_full": norm_full,
                "province": prov,
                "am": am,
                "norm_am": remove_accents(am),
            })

    _CACHED_COCAU = hubs
    _CACHED_TIME = now_ts
    return hubs

# ─── QUẢN LÝ TAB 'ĐIỂM DANH KAIZEN' TRÊN GOOGLE SHEETS ───────
ATTENDANCE_HEADERS = [
    "Ngày", "Thứ", "Mã BC (warehouse_id)", "Tên Bưu Cục / Kho", "Tỉnh", "AM Quản Lý",
    "Trạng Thái Ca Sáng", "Giờ Nộp Sáng", "Số Ảnh Sáng",
    "Trạng Thái Ca Tối", "Giờ Nộp Tối", "Số Ảnh Tối",
    "Ghi Chú", "Cập Nhật Lúc"
]

WEEKDAY_NAMES = {
    0: "Thứ Hai", 1: "Thứ Ba", 2: "Thứ Tư", 3: "Thứ Năm",
    4: "Thứ Sáu", 5: "Thứ Bảy", 6: "Chủ Nhật"
}

def ensure_attendance_worksheet():
    """Đảm bảo tab 'Điểm danh KAIZEN' tồn tại và có header chuẩn."""
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet(LOG_TAB_NAME)
    except Exception:
        ws = sh.add_worksheet(title=LOG_TAB_NAME, rows=500, cols=20)
        ws.append_row(ATTENDANCE_HEADERS)
        # Format tiêu đề bảng màu xanh GHN
        try:
            ws.format("A1:N1", {
                "backgroundColor": {"red": 0.08, "green": 0.45, "blue": 0.3},
                "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
                "horizontalAlignment": "CENTER"
            })
            ws.freeze(rows=1)
        except Exception:
            pass
    return ws

def ensure_today_attendance_rows(date_obj: date = None):
    """
    Đảm bảo có đủ 99 dòng tương ứng 99 bưu cục cho ngày được chỉ định.
    Nếu ngày đó chưa có dòng nào, khởi tạo toàn bộ 99 bưu cục với trạng thái '❌ Chưa báo cáo'.
    """
    if not date_obj:
        date_obj = get_vn_today()

    date_str = date_obj.strftime("%d/%m/%Y")
    day_name = WEEKDAY_NAMES.get(date_obj.weekday(), "")

    ws = ensure_attendance_worksheet()
    all_values = ws.get_all_values()

    # Kiểm tra những bưu cục đã có dòng trong ngày date_str
    existing_wids = set()
    for row in all_values[1:]:
        if len(row) >= 3 and row[0].strip() == date_str:
            existing_wids.add(row[2].strip())

    cocau = load_cocau_data()
    missing_hubs = [h for h in cocau if h["warehouse_id"] not in existing_wids]

    if missing_hubs:
        new_rows = []
        now_ts_str = get_vn_now().strftime("%d/%m/%Y %H:%M:%S")
        for h in missing_hubs:
            new_rows.append([
                date_str,
                day_name,
                h["warehouse_id"],
                h["hub_name"],
                h["province"],
                h["am"],
                "❌ Chưa báo cáo",  # Trạng thái Sáng
                "",                 # Giờ nộp Sáng
                "0/3 ảnh",          # Số ảnh Sáng
                "❌ Chưa báo cáo",  # Trạng thái Tối
                "",                 # Giờ nộp Tối
                "0/3 ảnh",          # Số ảnh Tối
                "",                 # Ghi chú
                now_ts_str          # Cập nhật lúc
            ])
        ws.append_rows(new_rows)
        time.sleep(1)

    return ws, date_str

def get_today_attendance_map(date_obj: date = None):
    """
    Lấy toàn bộ dữ liệu điểm danh của ngày date_obj dạng dictionary:
    {warehouse_id: {row_idx, data...}}
    """
    ws, date_str = ensure_today_attendance_rows(date_obj)
    all_values = ws.get_all_values()

    records = {}
    for idx, r in enumerate(all_values[1:], 2):
        if len(r) >= 6 and r[0].strip() == date_str:
            wid = r[2].strip()
            records[wid] = {
                "row_idx": idx,
                "date": r[0].strip(),
                "day_name": r[1].strip() if len(r) > 1 else "",
                "wid": wid,
                "name": r[3].strip() if len(r) > 3 else "",
                "province": r[4].strip() if len(r) > 4 else "",
                "am": r[5].strip() if len(r) > 5 else "",
                "morning_status": r[6].strip() if len(r) > 6 else "❌ Chưa báo cáo",
                "morning_time": r[7].strip() if len(r) > 7 else "",
                "morning_photos": r[8].strip() if len(r) > 8 else "0/3 ảnh",
                "evening_status": r[9].strip() if len(r) > 9 else "❌ Chưa báo cáo",
                "evening_time": r[10].strip() if len(r) > 10 else "",
                "evening_photos": r[11].strip() if len(r) > 11 else "0/3 ảnh",
                "note": r[12].strip() if len(r) > 12 else "",
                "updated_at": r[13].strip() if len(r) > 13 else "",
            }
    return ws, records

# ─── GỬI TIN NHẮN GTALK ─────────────────────────────────────
def send_gtalk_message(text: str, channel_id: str = None, oa_token: str = GTALK_OA_TOKEN):
    """Gửi tin nhắn định dạng HTML tới nhóm GTalk."""
    target_channel = str(channel_id or TEST_GROUP_ID)
    client_msg_id = str(int(get_vn_now().timestamp() * 1000))
    payload = {
        "channelId": target_channel,
        "clientMsgId": client_msg_id,
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

# ─── PARSER PAYLOAD TIN NHẮN & HÌNH ẢNH ──────────────────────
def parse_gtalk_message_payload(data: dict):
    """
    Trích xuất text/caption, danh sách hình ảnh và channelId từ payload webhook GTalk.
    Hỗ trợ đếm chính xác số lượng ảnh gửi trong 1 tin nhắn hoặc album.
    """
    msg_obj = data.get("message") or {}
    channel_id = str(data.get("channelId") or data.get("channel_id") or msg_obj.get("channelId") or "")

    sender_obj = data.get("sender") or msg_obj.get("sender") or {}
    sender_name = sender_obj.get("displayName") or sender_obj.get("name") or data.get("senderName") or ""

    content_raw = data.get("content") or msg_obj.get("content") or {}
    content_dict = {}
    if isinstance(content_raw, dict):
        content_dict = content_raw
    elif isinstance(content_raw, str):
        try:
            parsed = json.loads(content_raw)
            if isinstance(parsed, dict):
                content_dict = parsed
        except Exception:
            pass

    text = ""
    if content_dict.get("caption"):
        text = str(content_dict["caption"]).strip()
    elif content_dict.get("text"):
        text = str(content_dict["text"]).strip()
    elif isinstance(content_raw, str) and not content_dict:
        text = content_raw.strip()
    elif data.get("text"):
        text = str(data["text"]).strip()

    # Đếm số lượng hình ảnh trong mảng items
    items = content_dict.get("items") or []
    image_count = 0
    photo_file_ids = []

    if isinstance(items, list):
        for item in items:
            if isinstance(item, dict):
                img = item.get("image") or item.get("photo") or item.get("file")
                if img:
                    image_count += 1
                    if isinstance(img, dict) and img.get("fileId"):
                        photo_file_ids.append(str(img["fileId"]))
                    elif isinstance(img, str):
                        photo_file_ids.append(img)

    # Nếu không nằm trong items, kiểm tra contentType
    if image_count == 0:
        content_type = str(data.get("contentType") or msg_obj.get("contentType") or "").lower()
        if "image" in content_type or "photo" in content_type:
            image_count = 1
        elif "image" in str(data) or "fileId" in str(data):
            image_count = 1

    return {
        "channel_id": channel_id,
        "sender_name": sender_name,
        "text": text,
        "image_count": image_count,
        "photo_file_ids": photo_file_ids,
    }

# ─── SO KHỚP BƯU CỤC & CA LÀM VIỆC ──────────────────────────
def detect_hub(text: str, sender_name: str, cocau_list: list):
    """
    Nhận diện bưu cục từ tin nhắn:
    1. Ưu tiên so khớp số warehouse_id (4 đến 9 chữ số) xuất hiện trong text.
    2. Nếu người gửi là AM, ưu tiên so khớp các bưu cục do chính AM đó quản lý.
    3. So khớp tên bưu cục / bưu cục alias không dấu (longest match).
    """
    if not text:
        # Nếu chỉ có ảnh không có text, thử xem người gửi có phải AM quản lý duy nhất 1 BC không
        return None

    norm_text = remove_accents(text)
    norm_sender = remove_accents(sender_name)

    # 1. So khớp mã warehouse_id chính xác
    wid_matches = re.findall(r'\b(\d{4,9})\b', text)
    if wid_matches:
        cocau_by_wid = {h["warehouse_id"]: h for h in cocau_list}
        for candidate_wid in wid_matches:
            if candidate_wid in cocau_by_wid:
                return cocau_by_wid[candidate_wid]

    # 2. Xác định xem người gửi có phải AM nào trong danh sách không
    sender_hubs = []
    for h in cocau_list:
        if h["norm_am"] and h["norm_am"] in norm_sender:
            sender_hubs.append(h)

    # Nếu sender là AM, ưu tiên so khớp tên bưu cục của AM đó trước
    search_pools = [sender_hubs, cocau_list] if sender_hubs else [cocau_list]

    for pool in search_pools:
        # Sắp xếp các bưu cục theo độ dài từ khóa giảm dần (longest keyword match)
        candidates = sorted(pool, key=lambda x: len(x["norm_name"]), reverse=True)
        for h in candidates:
            # So khớp theo clean_name (ví dụ: 'thuan nam', 'duc trong', 'phan rang'...)
            if len(h["norm_name"]) >= 3:
                pattern = r'\b' + re.escape(h["norm_name"]) + r'\b'
                if re.search(pattern, norm_text):
                    return h
            # So khớp theo norm_full nếu có cả tiền tố
            if len(h["norm_full"]) >= 5 and h["norm_full"] in norm_text:
                return h

    return None

def detect_shift(text: str, now_dt: datetime) -> str:
    """
    Xác định ca báo cáo: 'morning' (đầu ngày) hoặc 'evening' (cuối ngày).
    - Có thể ép ca bằng tag: #sang, #ca1, #morning, #dau_ngay / #toi, #ca2, #evening, #cuoi_ngay
    - Mặc định: Trước 14:00 là ca đầu ngày, từ 14:00 trở đi là ca cuối ngày.
    """
    norm_txt = remove_accents(text)
    if any(k in norm_txt for k in ["#sang", "#ca1", "ca sang", "dau ngay", "sang"]):
        return "morning"
    if any(k in norm_txt for k in ["#toi", "#ca2", "ca toi", "cuoi ngay", "toi"]):
        return "evening"

    if now_dt.hour < 14:
        return "morning"
    else:
        return "evening"

# ─── XỬ LÝ LƯỢT CHECK-IN VÀ GHI GOOGLE SHEET ────────────────
def process_checkin_message(parsed_msg: dict, now_dt: datetime = None):
    """
    Xử lý 1 tin nhắn/ảnh check-in KAIZEN:
    - Tìm bưu cục
    - Xác định ca (Sáng / Tối)
    - Kiểm tra số ảnh (yêu cầu >= 3 ảnh TimestampCam)
    - Kiểm tra giờ nộp (<= 10:00 cho ca sáng, <= 22:00 cho ca tối)
    - Ghi vào Google Sheet
    - Gửi tin nhắn phản hồi vào nhóm
    """
    if not now_dt:
        now_dt = get_vn_now()

    text = parsed_msg["text"]
    image_count = parsed_msg["image_count"]
    sender_name = parsed_msg["sender_name"]
    channel_id = parsed_msg["channel_id"]

    # Bỏ qua nếu tin nhắn do chính Bot gửi (tránh echo loop)
    if any(kw in text for kw in [
        "XÁC NHẬN BÁO CÁO KAIZEN",
        "NHẮC NHỞ BÁO CÁO KAIZEN",
        "TỔNG KẾT BÁO CÁO KAIZEN",
        "CHỐT SỔ BÁO CÁO KAIZEN"
    ]):
        return {"status": "skipped", "reason": "bot message"}

    # Luôn tiếp nhận và phản hồi khi có người gửi báo cáo (kể cả gửi làm mẫu hôm nay)
    cocau = load_cocau_data()
    hub = detect_hub(text, sender_name, cocau)

    if not hub:
        wid_matches = re.findall(r'\b(\d{3,10})\b', text)
        norm_t = remove_accents(text)

        # Kiểm tra nếu gửi báo cáo cho kho đang tạm thời được miễn (ví dụ: CK Diên Điền)
        for ex_wid, ex_reason in EXCLUDED_WAREHOUSE_IDS.items():
            if (ex_wid in wid_matches) or (ex_wid in text) or ("dien dien" in norm_t and "ck" in norm_t):
                reply_ex = (
                    f"ℹ️ <b>THÔNG BÁO MIỄN BÁO CÁO KAIZEN:</b>\n"
                    f"Bưu cục/Kho <b>(KHO) CK Diên Điền (23098000)</b> hiện dùng chung mặt bằng với <b>Kho Trung Chuyển Khánh Hòa (1909)</b> nên tạm thời được miễn báo cáo riêng."
                )
                send_gtalk_message(reply_ex, channel_id)
                return {"status": "excluded_hub", "wid": ex_wid, "reason": ex_reason}

        # Nếu có gửi ảnh kèm số nhưng không khớp warehouse_id
        if wid_matches and image_count > 0:
            reply_err = (
                f"⚠️ <b>LƯU Ý ĐIỂM DANH KAIZEN:</b>\n"
                f"Mã kho/bưu cục <b>{wid_matches[0]}</b> không tồn tại trong danh sách Cơ cấu Vùng NTB.\n"
                f"Vui lòng kiểm tra lại mã warehouse_id hoặc cú pháp gửi."
            )
            send_gtalk_message(reply_err, channel_id)
            return {"status": "error", "reason": "invalid_wid", "wid": wid_matches[0]}
        return {"status": "ignored", "reason": "no_hub_detected"}

    shift = detect_shift(text, now_dt)
    now_hm = now_dt.strftime("%H:%M:%S")
    now_ts_str = now_dt.strftime("%d/%m/%Y %H:%M:%S")

    # Xác định đúng giờ hay trễ
    if shift == "morning":
        cutoff_time = now_dt.replace(hour=MORNING_CUTOFF_HOUR, minute=MORNING_CUTOFF_MIN, second=59, microsecond=0)
        is_on_time = (now_dt <= cutoff_time)
        shift_label = "🌅 Ca đầu ngày (trước 10:00)"
    else:
        cutoff_time = now_dt.replace(hour=EVENING_CUTOFF_HOUR, minute=EVENING_CUTOFF_MIN, second=59, microsecond=0)
        is_on_time = (now_dt <= cutoff_time)
        shift_label = "🌙 Ca cuối ngày (trước 22:00)"

    # Lấy thông tin ngày hiện tại trên Sheet
    ws, records = get_today_attendance_map(now_dt.date())
    wid = hub["warehouse_id"]
    rec = records.get(wid)

    if not rec or not rec.get("row_idx"):
        return {"status": "error", "message": f"Cannot find row for warehouse {wid}"}

    row_idx = rec["row_idx"]

    # Xử lý cộng dồn ảnh nếu trước đó đã gửi một phần ảnh trong cùng ca
    prev_photos_str = rec["morning_photos"] if shift == "morning" else rec["evening_photos"]
    prev_count = 0
    m_p = re.search(r'(\d+)', prev_photos_str)
    if m_p:
        prev_count = int(m_p.group(1))

    total_photos = image_count
    # Nếu tin nhắn mới có ảnh, cộng dồn với số ảnh đã nộp trước đó trong ca (nếu tin trước chưa đủ 3 ảnh)
    if prev_count > 0 and prev_count < 3 and image_count > 0:
        total_photos = min(3, prev_count + image_count)

    # Đánh giá trạng thái hoàn thành & ảnh
    if total_photos >= 3:
        status_label = "✅ Đúng giờ" if is_on_time else "⚠️ Báo trễ"
        photo_label = f"✅ Đủ {total_photos}/3 ảnh TimestampCam"
        photo_status = "Đủ ảnh"
    else:
        status_label = f"📸 Thiếu ảnh ({total_photos}/3 ảnh)" + ("" if is_on_time else " [Trễ]")
        photo_label = f"⚠️ Mới có {total_photos}/3 ảnh (Cần bổ sung đủ 3 ảnh: Tổng quan BC, Nhà VS, Mặt tiền)"
        photo_status = "Thiếu ảnh"

    photos_cell = f"{total_photos}/3 ảnh"
    note_cell = f"Người gửi: {sender_name}" if sender_name else ""

    # Cập nhật vào Google Sheet
    # Dòng: A=Ngày, B=Thứ, C=Mã BC, D=Tên BC, E=Tỉnh, F=AM,
    # G=Trạng thái Sáng, H=Giờ Sáng, I=Ảnh Sáng,
    # J=Trạng thái Tối, K=Giờ Tối, L=Ảnh Tối,
    # M=Ghi chú, N=Cập nhật lúc
    if shift == "morning":
        ws.update(
            values=[[status_label, now_hm, photos_cell]],
            range_name=f"G{row_idx}:I{row_idx}"
        )
        ws.update(
            values=[[note_cell, now_ts_str]],
            range_name=f"M{row_idx}:N{row_idx}"
        )
    else:
        ws.update(
            values=[[status_label, now_hm, photos_cell]],
            range_name=f"J{row_idx}:L{row_idx}"
        )
        ws.update(
            values=[[note_cell, now_ts_str]],
            range_name=f"M{row_idx}:N{row_idx}"
        )

    # Gửi tin nhắn phản hồi xác nhận vào nhóm
    time_status_tag = "✅ Đúng hạn" if is_on_time else "⚠️ Báo trễ"
    extra_note = ""
    if now_dt.weekday() not in ACTIVE_WEEKDAYS:
        extra_note = "\n<i>(🧪 Ghi nhận gửi mẫu thử nghiệm — Lịch điểm danh chính thức: Thứ 2, Thứ 4, Thứ 6)</i>"

    reply_msg = (
        f"✅ <b>XÁC NHẬN BÁO CÁO KAIZEN</b>\n"
        f"🏢 <b>Bưu cục:</b> {hub['hub_name']} (Mã: <code>{wid}</code>)\n"
        f"📍 <b>Tỉnh:</b> {hub['province']} | 👤 <b>AM:</b> {hub['am']}\n"
        f"📅 <b>Ca:</b> {shift_label}\n"
        f"⏰ <b>Thời gian:</b> {now_hm} ({time_status_tag})\n"
        f"📸 <b>Hình ảnh:</b> {photo_label}\n"
        f"📊 <b>Kết quả:</b> <b>{status_label}</b>\n"
        f"🔗 <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Xem bảng theo dõi KAIZEN</a>"
        f"{extra_note}"
    )
    target_ch = channel_id or PROVINCE_GROUPS.get(hub['province']) or TEST_GROUP_ID
    send_gtalk_message(reply_msg, target_ch)

    return {
        "status": "success",
        "hub": hub["hub_name"],
        "wid": wid,
        "shift": shift,
        "photos": total_photos,
        "status_label": status_label,
        "is_on_time": is_on_time
    }

# ─── KỊCH BẢN TỰ ĐỘNG: NHẮC NHỞ & TỔNG KẾT ───────────────────
def build_province_group_target_map():
    """
    Nhóm danh sách các tỉnh theo từng Channel ID.
    Ví dụ trong lúc test: {'2077278419534073856': ['Khánh Hòa', 'Lâm Đồng', 'Đắk Nông', 'Bình Thuận', 'Ninh Thuận']}
    Khi có 5 group riêng: mỗi channel_id sẽ ứng với 1 tỉnh tương ứng.
    """
    group_map = defaultdict(list)
    for prov, g_id in PROVINCE_GROUPS.items():
        group_map[g_id].append(prov)
    return group_map

def run_reminder(shift: str = "morning", force: bool = False):
    """
    Gửi tin nhắn nhắc nhở trước giờ chốt 5 phút (09:55 sáng hoặc 21:55 tối).
    Liệt kê chi tiết từng AM và các bưu cục chưa nộp / nộp thiếu ảnh.
    """
    now = get_vn_now()
    if not force and not TEST_EVERYDAY and now.weekday() not in ACTIVE_WEEKDAYS:
        return

    today_date = now.date()
    ws, records = get_today_attendance_map(today_date)
    group_map = build_province_group_target_map()

    shift_title = "ĐẦU NGÀY (TRƯỚC 10:00)" if shift == "morning" else "CUỐI NGÀY (TRƯỚC 22:00)"
    time_str = "09:55" if shift == "morning" else "21:55"
    deadline_str = "10:00" if shift == "morning" else "22:00"

    for channel_id, prov_list in group_map.items():
        # Lấy danh sách bưu cục thuộc các tỉnh của group này
        target_recs = [r for r in records.values() if r["province"] in prov_list]
        
        # Lọc ra các bưu cục chưa báo cáo hoặc còn thiếu ảnh
        missing_by_am = defaultdict(list)
        for r in target_recs:
            status = r["morning_status"] if shift == "morning" else r["evening_status"]
            photos = r["morning_photos"] if shift == "morning" else r["evening_photos"]
            
            # Nếu chưa báo cáo hoặc thiếu ảnh
            if "Chưa báo cáo" in status or "Thiếu ảnh" in status:
                am_name = r["am"] or "Chưa phân công"
                extra = " <i>(Thiếu ảnh)</i>" if "Thiếu ảnh" in status else ""
                missing_by_am[am_name].append(f"{r['name']} (<code>{r['wid']}</code>){extra}")

        if not missing_by_am:
            print(f"[{now.strftime('%H:%M:%S')}] Group {channel_id}: 100% bưu cục đã hoàn thành báo cáo KAIZEN {shift}.")
            continue

        total_missing = sum(len(bcs) for bcs in missing_by_am.values())
        prov_header = ", ".join(prov_list) if len(prov_list) < 5 else "5 Tỉnh NTB"

        lines = [
            f"⏰ <b>[NHẮC NHỞ BÁO CÁO KAIZEN {shift_title}] - {time_str}</b>",
            f"<i>Khu vực: <b>{prov_header}</b> — Còn 5 phút trước giờ chốt ({deadline_str})!</i>",
            f"📌 <b>Còn {total_missing} bưu cục/kho chưa hoàn thành đủ 3 ảnh:</b>",
            ""
        ]

        for idx, (am_name, bc_list) in enumerate(sorted(missing_by_am.items(), key=lambda x: -len(x[1])), 1):
            lines.append(f"{idx}. 👤 <b>AM {am_name}</b> ({len(bc_list)} BC):")
            for bc_item in bc_list:
                lines.append(f"   ▫️ {bc_item}")

        lines.extend([
            "",
            f"👉 <b>Yêu cầu các AM/NVXL khẩn trương gửi mã bưu cục kèm 3 ảnh TimestampCam trước {deadline_str}!</b>",
            f"<i>(3 ảnh gồm: Tổng quan BC, Nhà VS, Không gian trước mặt tiền BC)</i>",
            f"🔗 <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Bảng theo dõi KAIZEN Google Sheet</a>"
        ])

        msg_body = "\n".join(lines)
        ok, res = send_gtalk_message(msg_body, channel_id)
        print(f"[{now.strftime('%H:%M:%S')}] Gửi nhắc nhở {shift} tới group {channel_id}: {ok} - {res}")

def run_recap(shift: str = "morning", force: bool = False):
    """
    Gửi tin nhắn tổng kết sau giờ chốt 5 phút (10:05 sáng hoặc 22:05 tối).
    Báo cáo tỷ lệ hoàn thành, danh sách bưu cục trễ / thiếu ảnh / chưa nộp.
    """
    now = get_vn_now()
    if not force and not TEST_EVERYDAY and now.weekday() not in ACTIVE_WEEKDAYS:
        return

    today_date = now.date()
    today_str = today_date.strftime("%d/%m/%Y")
    ws, records = get_today_attendance_map(today_date)
    group_map = build_province_group_target_map()

    shift_title = "ĐẦU NGÀY (CA SÁNG)" if shift == "morning" else "CUỐI NGÀY (CHỐT SỔ)"
    time_str = "10:05" if shift == "morning" else "22:05"

    for channel_id, prov_list in group_map.items():
        target_recs = [r for r in records.values() if r["province"] in prov_list]
        total_hubs = len(target_recs)
        if total_hubs == 0:
            continue

        done_on_time = []
        done_late    = []
        missing_pics = []
        not_reported = []

        for r in target_recs:
            status = r["morning_status"] if shift == "morning" else r["evening_status"]
            if "Đúng giờ" in status:
                done_on_time.append(r)
            elif "Báo trễ" in status:
                done_late.append(r)
            elif "Thiếu ảnh" in status:
                missing_pics.append(r)
            else:
                not_reported.append(r)

        rate = ((len(done_on_time) + len(done_late)) / total_hubs) * 100 if total_hubs > 0 else 0
        prov_header = ", ".join(prov_list) if len(prov_list) < 5 else "5 Tỉnh NTB"

        lines = [
            f"📊 <b>[TỔNG KẾT ĐIỂM DANH KAIZEN {shift_title}] - {time_str}</b>",
            f"📅 Ngày: <b>{today_str}</b> | Khu vực: <b>{prov_header}</b>",
            "─────────────────────────────",
            f"🏢 <b>Tổng số bưu cục/kho:</b> {total_hubs}",
            f"✅ <b>Đúng giờ:</b> {len(done_on_time)} BC ({len(done_on_time)*100/total_hubs:.1f}%)",
            f"⚠️ <b>Báo trễ:</b> {len(done_late)} BC ({len(done_late)*100/total_hubs:.1f}%)",
            f"📸 <b>Thiếu ảnh (&lt;3 ảnh):</b> {len(missing_pics)} BC ({len(missing_pics)*100/total_hubs:.1f}%)",
            f"❌ <b>Chưa báo cáo:</b> {len(not_reported)} BC ({len(not_reported)*100/total_hubs:.1f}%)",
            f"🎯 <b>Tỷ lệ hoàn thành:</b> <b>{rate:.1f}%</b>",
            "─────────────────────────────"
        ]

        # Liệt kê các bưu cục chưa báo cáo
        if not_reported:
            lines.append("❌ <b>DANH SÁCH CHƯA BÁO CÁO:</b>")
            unreported_by_am = defaultdict(list)
            for r in not_reported:
                unreported_by_am[r["am"] or "Khác"].append(r["name"])
            for am_name, bcs in sorted(unreported_by_am.items(), key=lambda x: -len(x[1])):
                lines.append(f"• <b>AM {am_name}:</b> {', '.join(bcs)}")
            lines.append("")

        # Liệt kê bưu cục thiếu ảnh
        if missing_pics:
            lines.append("📸 <b>DANH SÁCH THIẾU ẢNH (&lt;3 ẢNH):</b>")
            for r in missing_pics:
                p_count = r["morning_photos"] if shift == "morning" else r["evening_photos"]
                lines.append(f"• <b>{r['name']}</b> (AM {r['am']}): mới gửi {p_count}")
            lines.append("")

        lines.extend([
            f"🔗 <b>Chi tiết:</b> <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Google Sheet Điểm danh KAIZEN</a>"
        ])

        msg_body = "\n".join(lines)
        ok, res = send_gtalk_message(msg_body, channel_id)
        print(f"[{now.strftime('%H:%M:%S')}] Gửi tổng kết {shift} tới group {channel_id}: {ok} - {res}")

# ─── SCHEDULER CHẠY NGẦM ĐỊNH KỲ ───────────────────────────
last_scheduled_execution = {}

def check_kaizen_schedule():
    """Hàm kiểm tra lịch định kỳ: 09:55, 10:05, 21:55, 22:05."""
    now = get_vn_now()
    today_str = now.strftime("%Y-%m-%d")
    hm = now.strftime("%H:%M")

    # Kiểm tra ngày trong tuần nếu không ở chế độ TEST_EVERYDAY
    if not TEST_EVERYDAY and now.weekday() not in ACTIVE_WEEKDAYS:
        return

    schedule_jobs = {
        "09:55": ("job_remind_morning", lambda: run_reminder(shift="morning")),
        "10:05": ("job_recap_morning",  lambda: run_recap(shift="morning")),
        "21:55": ("job_remind_evening", lambda: run_reminder(shift="evening")),
        "22:05": ("job_recap_evening",  lambda: run_recap(shift="evening")),
    }

    if hm in schedule_jobs:
        job_key_name, job_func = schedule_jobs[hm]
        trigger_key = f"{today_str}_{job_key_name}"
        if last_scheduled_execution.get(trigger_key):
            return
        last_scheduled_execution[trigger_key] = True
        print(f"⏰ [KAIZEN SCHEDULER] Kích hoạt {job_key_name} lúc {hm}...")
        try:
            job_func()
        except Exception as e:
            print(f"❌ Lỗi chạy {job_key_name}: {e}")

# ─── ENTRYPOINT CHO WEBHOOK SERVER ──────────────────────────
def handle_kaizen_webhook(data: dict):
    """
    Hàm được gọi từ Flask Webhook Server khi có request POST từ GTalk.
    """
    try:
        parsed = parse_gtalk_message_payload(data)
        # Chỉ xử lý nếu có text hoặc có ảnh
        if not parsed["text"] and parsed["image_count"] == 0:
            return {"status": "skipped", "reason": "empty_payload"}

        res = process_checkin_message(parsed)
        return res
    except Exception as e:
        print(f"[KAIZEN WEBHOOK ERROR]: {e}")
        return {"status": "error", "message": str(e)}

# ─── STANDALONE FLASK SERVER ─────────────────────────────────
def start_standalone_server(port: int = 5056):
    """Chạy Flask Webhook Server độc lập cho Bot KAIZEN."""
    from flask import Flask, request, jsonify

    app = Flask(__name__)

    @app.route("/", methods=["GET"])
    def home():
        return jsonify({
            "service": "GHN NTB KAIZEN Attendance Bot",
            "status": "online",
            "server_time": get_vn_now().strftime("%d/%m/%Y %H:%M:%S")
        })

    @app.route("/gtalk/webhook", methods=["POST"])
    @app.route("/webhook", methods=["POST"])
    def gtalk_hook():
        data = request.get_json(force=True, silent=True) or {}
        res = handle_kaizen_webhook(data)
        return jsonify(res)

    @app.route("/cron/<action>/<shift>", methods=["GET", "POST"])
    def manual_trigger(action, shift):
        if action == "remind":
            run_reminder(shift=shift, force=True)
            return jsonify({"status": "triggered_reminder", "shift": shift})
        elif action == "recap":
            run_recap(shift=shift, force=True)
            return jsonify({"status": "triggered_recap", "shift": shift})
        return jsonify({"status": "unknown_action"}), 400

    def scheduler_loop():
        print("⏰ KAIZEN Scheduler Loop đã chạy (Giám sát: 09:55, 10:05, 21:55, 22:05)...")
        while True:
            try:
                check_kaizen_schedule()
            except Exception as e:
                print(f"Scheduler loop error: {e}")
            time.sleep(20)

    threading.Thread(target=scheduler_loop, daemon=True).start()
    print(f"🚀 Bot KAIZEN Server đang chạy trên port {port}...")
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)

# ─── CLI COMMANDS ───────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Bot Điểm danh KAIZEN GHN Vùng NTB")
    parser.add_argument("--server", action="store_true", help="Chạy standalone Flask Webhook Server")
    parser.add_argument("--port", type=int, default=5056, help="Port cho server (default: 5056)")
    parser.add_argument("--init-sheet", action="store_true", help="Khởi tạo bảng tính 'Điểm danh KAIZEN'")
    parser.add_argument("--remind", choices=["morning", "evening"], help="Kích hoạt tin nhắc nhở ngay")
    parser.add_argument("--recap", choices=["morning", "evening"], help="Kích hoạt tin tổng kết ngay")
    parser.add_argument("--test-msg", type=str, help="Gửi tin nhắn thử vào group test")
    parser.add_argument("--test-checkin", type=str, help="Giả lập báo cáo (ví dụ: '1909' hoặc 'Thuận Nam')")
    parser.add_argument("--photos", type=int, default=3, help="Số lượng ảnh giả lập (default: 3)")
    parser.add_argument("--sender", type=str, default="Nguyễn Tiến Lực", help="Tên người gửi giả lập")

    args = parser.parse_args()

    if args.init_sheet:
        print("🔄 Đang khởi tạo bảng tính 'Điểm danh KAIZEN'...")
        ws, d_str = ensure_today_attendance_rows()
        print(f"✅ Đã khởi tạo thành công tab '{LOG_TAB_NAME}' cho ngày {d_str}!")

    elif args.remind:
        print(f"⏰ Kích hoạt nhắc nhở ca {args.remind}...")
        run_reminder(shift=args.remind, force=True)

    elif args.recap:
        print(f"📊 Kích hoạt tổng kết ca {args.recap}...")
        run_recap(shift=args.recap, force=True)

    elif args.test_msg:
        print(f"📨 Đang gửi tin test tới group {TEST_GROUP_ID}...")
        ok, res = send_gtalk_message(args.test_msg)
        print("Kết quả:", ok, res)

    elif args.test_checkin:
        print(f"🧪 Giả lập gửi báo cáo: '{args.test_checkin}' ({args.photos} ảnh, người gửi: {args.sender})...")
        payload = {
            "channelId": TEST_GROUP_ID,
            "sender": {"displayName": args.sender},
            "content": json.dumps({
                "caption": args.test_checkin,
                "items": [{"image": {"fileId": f"fake_{i}"}} for i in range(args.photos)]
            })
        }
        res = handle_kaizen_webhook(payload)
        print("Kết quả xử lý:", json.dumps(res, ensure_ascii=False, indent=2))

    elif args.server:
        start_standalone_server(port=args.port)

    else:
        print("Bot Điểm danh KAIZEN NTB.")
        print("Dùng các tham số: --init-sheet, --remind [morning|evening], --recap [morning|evening], --test-checkin, --server")
        ws, d_str = ensure_today_attendance_rows()
        print(f"Dữ liệu sẵn sàng cho ngày {d_str}.")
