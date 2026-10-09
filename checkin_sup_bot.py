# -*- coding: utf-8 -*-
"""
Module: checkin_sup_bot.py
Xử lý Điểm danh & Check-in đầu ngày dành cho SUP/AM và HR Vùng NTB qua G-Talk Webhook.
- Nhận tin nhắn / ảnh check-in TimestampCam từ Group G-Talk: 2099483038556782592
- Tự động phân tích nhân sự (3 SUP/AM + 6 HR), kho/bưu cục (tra cứu từ tab Cơ cấu), thời gian gửi và ghi vào Google Sheet
- Nhắc nhở lúc 07:55 (trước 8h 5 phút)
- Cảnh báo trễ & yêu cầu gửi bù lúc 08:00
- Báo cáo tổng hợp lúc 11:00 (chia rõ nhóm SUP/AM và nhóm HR)
- Hỗ trợ xin miễn check-in / nghỉ phép theo cú pháp "AM/SUP/HR [Tên] xin miễn check in vì lý do..." hoặc tick tay trực tiếp trên Google Sheet.
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
GTALK_CHANNEL_ID = "2099483038556782592"
GTALK_API_URL    = "https://mbff.ghn.vn/api/gtalk/send-message"

SPREADSHEET_ID   = "1sR4bqfatBj7bI2KWzwAPDWeifsfg2FICOZYauuzqBeE"
SHEET_TAB_NAME   = "Sheet1"
SHEET_COCAU_TAB  = "Cơ cấu"

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

# ─── DANH SÁCH NHÂN SỰ CẦN CHECK-IN (3 SUP/AM + 6 HR) ──────
MEMBERS = {
    # ── 1. NHÓM SUP / AM (3 người) ──
    "am_luc_nt": {
        "id": "am_luc_nt",
        "emp_code": "AM_LUC",
        "full_name": "Nguyễn Tiến Lực",
        "display_name": "Nguyễn Tiến Lực",
        "role": "SUP/AM",
        "aliases": ["nguyễn tiến lực", "nguyen tien luc", "tiến lực", "tien luc", "am lực", "am luc", "sup lực", "sup luc", "lực", "luc"],
        "avoid_confusions": []
    },
    "am_hoang_nm": {
        "id": "am_hoang_nm",
        "emp_code": "AM_HOANG",
        "full_name": "Nguyễn Minh Hoàng",
        "display_name": "Nguyễn Minh Hoàng",
        "role": "SUP/AM",
        "aliases": ["nguyễn minh hoàng", "nguyen minh hoang", "minh hoàng", "minh hoang", "am hoàng", "am hoang", "sup hoàng", "sup hoang", "hoàng", "hoang"],
        "avoid_confusions": []
    },
    "am_khanh_nn": {
        "id": "am_khanh_nn",
        "emp_code": "AM_KHANH",
        "full_name": "Nguyễn Ngọc Khánh",
        "display_name": "Nguyễn Ngọc Khánh",
        "role": "SUP/AM",
        "aliases": ["nguyễn ngọc khánh", "nguyen ngoc khanh", "ngọc khánh", "ngoc khanh", "am khánh", "am khanh", "sup khánh", "sup khanh", "khánh", "khanh"],
        "avoid_confusions": ["khanh hoa", "khanh vinh", "khanh son"]
    },

    # ── 2. NHÓM HR / HRBP (6 người) ──
    "hr_hoa_ttk": {
        "id": "hr_hoa_ttk",
        "emp_code": "3164867",
        "full_name": "Trương Thị Kim Hoa",
        "display_name": "Trương Thị Kim Hoa (3164867)",
        "role": "HR",
        "aliases": ["3164867", "trương thị kim hoa", "truong thi kim hoa", "kim hoa", "hr hoa", "hrbp hoa", "hoa"],
        "avoid_confusions": ["khanh hoa", "ninh hoa", "dong ninh hoa", "hoa ninh", "hoa dinh", "xuan hoa"]
    },
    "hr_tuananh_l": {
        "id": "hr_tuananh_l",
        "emp_code": "3169241",
        "full_name": "Lê Tuấn Anh",
        "display_name": "Lê Tuấn Anh (3169241)",
        "role": "HR",
        "aliases": ["3169241", "lê tuấn anh", "le tuan anh", "tuấn anh", "tuan anh", "hr tuấn anh", "hr tuan anh", "hrbp tuấn anh", "hrbp tuan anh"],
        "avoid_confusions": []
    },
    "hr_phuong_nht": {
        "id": "hr_phuong_nht",
        "emp_code": "3172797",
        "full_name": "Nguyễn Hoàng Trúc Phương",
        "display_name": "Nguyễn Hoàng Trúc Phương (3172797)",
        "role": "HR",
        "aliases": ["3172797", "nguyễn hoàng trúc phương", "nguyen hoang truc phuong", "trúc phương", "truc phuong", "hr trúc phương", "hr truc phuong", "hrbp trúc phương", "hrbp truc phuong", "hr phương", "hr phuong", "hrbp phương", "hrbp phuong", "phương", "phuong"],
        "avoid_confusions": ["phuong dinh", "tuyen quang", "lien huong", "phuoc dinh"]
    },
    "hr_dat_vdq": {
        "id": "hr_dat_vdq",
        "emp_code": "3176531",
        "full_name": "Vũ Đình Quốc Đạt",
        "display_name": "Vũ Đình Quốc Đạt (3176531)",
        "role": "HR",
        "aliases": ["3176531", "vũ đình quốc đạt", "vu dinh quoc dat", "quốc đạt", "quoc dat", "hr đạt", "hr dat", "hrbp đạt", "hrbp dat", "đạt", "dat"],
        "avoid_confusions": []
    },
    "hr_thinh_nt": {
        "id": "hr_thinh_nt",
        "emp_code": "3181498",
        "full_name": "Nguyễn Trường Thịnh",
        "display_name": "Nguyễn Trường Thịnh (3181498)",
        "role": "HR",
        "aliases": ["3181498", "nguyễn trường thịnh", "nguyen truong thinh", "trường thịnh", "truong thinh", "hr thịnh", "hr thinh", "hrbp thịnh", "hrbp thinh", "thịnh", "thinh"],
        "avoid_confusions": []
    },
    "hr_tu_ttt": {
        "id": "hr_tu_ttt",
        "emp_code": "3098843",
        "full_name": "Thái Thị Thanh Tú",
        "display_name": "Thái Thị Thanh Tú (3098843)",
        "role": "HR",
        "aliases": ["3098843", "thái thị thanh tú", "thai thi thanh tu", "thanh tú", "thanh tu", "hr tú", "hr tu", "hrbp tú", "hrbp tu", "tú", "tu"],
        "avoid_confusions": ["tu bong", "cu jut", "tuy duc", "tuy phong", "tuyen quang"]
    }
}

# Alias cũ để tương thích
SUPS = MEMBERS

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

def get_worksheet_cocau():
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        return sh.worksheet(SHEET_COCAU_TAB)
    except Exception:
        return None

# ─── CACHE BƯU CỤC TỪ TAB CƠ CẤU ───────────────────────────
CACHED_LOCATIONS = []
CACHED_LOCATIONS_TIME = 0

def get_location_registry():
    global CACHED_LOCATIONS, CACHED_LOCATIONS_TIME
    now_ts = time.time()
    # Cache 1 tiếng
    if CACHED_LOCATIONS and (now_ts - CACHED_LOCATIONS_TIME < 3600):
        return CACHED_LOCATIONS
    try:
        ws_cc = get_worksheet_cocau()
        if not ws_cc:
            return []
        rows = ws_cc.get_all_values()
        locs = []
        for r in rows[1:]:
            if len(r) >= 2 and r[1].strip():
                raw_name = r[1].strip()
                clean_name = re.sub(r'^\([A-Z0-9]+\)\s*', '', raw_name).strip()
                locs.append({
                    "raw": raw_name,
                    "clean": clean_name,
                    "norm": remove_accents(clean_name),
                    "norm_raw": remove_accents(raw_name)
                })
        # Ưu tiên các bưu cục có tên dài và chi tiết trước
        locs.sort(key=lambda x: len(x["clean"]), reverse=True)
        CACHED_LOCATIONS = locs
        CACHED_LOCATIONS_TIME = now_ts
        return locs
    except Exception as e:
        print(f"[WARN] Không thể đọc tab Cơ cấu: {e}")
        return CACHED_LOCATIONS or []

def ensure_today_rows(ws, target_date_str=None):
    """Đảm bảo có đủ 9 dòng cho 3 SUP và 6 HR trong ngày target_date_str (DD/MM/YYYY)."""
    if not target_date_str:
        target_date_str = get_vn_today().strftime("%d/%m/%Y")

    rows = ws.get_all_values()
    if not rows:
        headers = ['Ngày', 'Nhân sự (SUP/HR)', 'Trạng thái', 'Giờ check-in', 'Kho / Bưu cục check-in', 'Ghi chú', 'Chi tiết / Ảnh', 'Thời gian cập nhật']
        ws.append_row(headers)
        rows = [headers]

    # Kiểm tra xem nhân sự nào đã có dòng hôm nay
    existing_members = set()
    for idx, r in enumerate(rows[1:], 2):
        if len(r) >= 2 and r[0].strip() == target_date_str:
            row_name = r[1].strip()
            for m_id, m_info in MEMBERS.items():
                if m_info["full_name"].lower() in row_name.lower():
                    existing_members.add(m_id)

    # Thêm dòng cho nhân sự còn thiếu
    appended = False
    for m_id, m_info in MEMBERS.items():
        if m_id not in existing_members:
            ws.append_row([
                target_date_str,
                m_info["display_name"],
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
    """Lấy dữ liệu check-in của 9 nhân sự trong ngày từ Google Sheet."""
    ws = get_worksheet()
    target_date_str = ensure_today_rows(ws, target_date_str)
    rows = ws.get_all_values()

    records = {}
    for m_id, m_info in MEMBERS.items():
        records[m_id] = {
            "row_idx": None,
            "member_id": m_id,
            "name": m_info["full_name"],
            "display_name": m_info["display_name"],
            "role": m_info["role"],
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
            row_name = r[1].strip()
            for m_id, m_info in MEMBERS.items():
                if m_info["full_name"].lower() in row_name.lower():
                    records[m_id]["row_idx"] = idx
                    records[m_id]["status"] = r[2].strip() if len(r) > 2 else "Chưa check-in"
                    records[m_id]["time"] = r[3].strip() if len(r) > 3 else ""
                    records[m_id]["location"] = r[4].strip() if len(r) > 4 else ""
                    records[m_id]["note"] = r[5].strip() if len(r) > 5 else ""
                    records[m_id]["details"] = r[6].strip() if len(r) > 6 else ""
                    records[m_id]["updated_at"] = r[7].strip() if len(r) > 7 else ""

    return ws, records

def update_sup_record(ws, row_idx, status, checkin_time="", location="", note="", details=""):
    """Cập nhật dòng dữ liệu nhân sự trên Google Sheet."""
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

# ─── PARSE PAYLOAD TIN NHẮN & ẢNH GTALK ─────────────────────

def parse_gtalk_message_payload(data: dict):
    msg_obj = data.get("message") or {}
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

    has_photo = False
    photo_file_id = ""
    items = content_dict.get("items") or []
    if isinstance(items, list):
        for item in items:
            if isinstance(item, dict):
                img = item.get("image") or item.get("photo") or item.get("file")
                if img:
                    has_photo = True
                    if isinstance(img, dict) and img.get("fileId"):
                        photo_file_id = str(img["fileId"])
                    elif isinstance(img, str):
                        photo_file_id = img

    if not has_photo:
        content_type = str(data.get("contentType") or msg_obj.get("contentType") or "").lower()
        if "image" in content_type or "photo" in content_type or "file" in content_type:
            has_photo = True
        elif "image" in str(data) or "fileId" in str(data):
            has_photo = True

    return text, has_photo, photo_file_id

# ─── NHẬN DIỆN NHÂN SỰ & NỘI DUNG CHECK-IN ──────────────────

def remove_accents(input_str: str) -> str:
    """Loại bỏ dấu tiếng Việt chuẩn xác (bao gồm cả chữ đ/Đ)."""
    if not input_str:
        return ""
    input_str = input_str.replace('đ', 'd').replace('Đ', 'd')
    nfkd = unicodedata.normalize('NFKD', input_str)
    return ''.join([c for c in nfkd if not unicodedata.combining(c)]).lower()

def detect_location(text: str) -> str:
    """
    Trích xuất kho hoặc bưu cục từ tin nhắn check-in.
    Tra cứu đối chiếu với 99 bưu cục trong tab 'Cơ cấu' và các kho chuyển tiếp lớn.
    """
    norm_txt = remove_accents(text)

    # 1. Các Kho Chuyển Tiếp / Kho Trung Chuyển lớn của NTB
    if re.search(r'\b(kct|ktc|kho)\s*(dt|d\s*trong|duc\s*trong|ductrong)\b|\bkctdt\b|\bduc\s*trong\b', norm_txt) and 'duc trong 1' not in norm_txt and 'duc trong 2' not in norm_txt:
        return "Kho Chuyển Tiếp Đức Trọng"
    if re.search(r'\b(kct|ktc|kho)\s*(bt|b\s*thuan|bthuan|binh\s*thuan)\b|\bkctbt\b|\bktcbt\b|\bbinh\s*thuan\b', norm_txt):
        return "Kho Chuyển Tiếp Bình Thuận"
    if re.search(r'\b(ktc|kct|kho)\s*(kh|k\s*hoa|khoa|khanh\s*hoa)\b|\bktckh\b|\bkctkh\b|\bkhanh\s*hoa\b', norm_txt) and not re.search(r'\b(khanh\s*vinh|khanh\s*son|bac\s*nha\s*trang)\b', norm_txt):
        return "Kho Trung Chuyển Khánh Hòa"
    if re.search(r'\b(kct|ktc|kho)\s*(dn|dno|dak\s*nong|daknong)\b|\bkctdn\b|\bkctdno\b|\bdak\s*nong\b|\bdno\b', norm_txt):
        return "Kho Chuyển Tiếp Đắk Nông"
    if re.search(r'\b(kct|ktc|kho)\s*(bl|b\s*loc|bao\s*loc|baoloc)\b|\bkctbl\b|\bbao\s*loc\b', norm_txt) and 'bao loc 1' not in norm_txt and 'bao loc 3' not in norm_txt:
        return "Kho Chuyển Tiếp Bảo Lộc"

    # 2. Tra cứu động đối chiếu với danh sách bưu cục từ tab Cơ cấu
    locations = get_location_registry()
    for loc in locations:
        # Khớp theo tên chuẩn bưu cục (đã bỏ tiền tố mã tỉnh)
        if re.search(r'\b' + re.escape(loc["norm"]) + r'\b', norm_txt):
            return loc["raw"]

    # 3. Nhận diện dạng 'BC <Tên>' hoặc 'Bưu cục <Tên>'
    bc_match = re.search(r'(?i)\b(?:bc|bưu\s*cục)\s+([^\n\r,\:\;]+)', text)
    if bc_match:
        name_clean = bc_match.group(1).strip()
        return f"BC {name_clean.title()}"

    # 4. Nếu có text sau từ checkin mà chưa khớp các từ trên
    clean = re.sub(r'(?i)#?checkin|#?check\s*in|diem\s*danh|gui\s*checkin', '', text).strip()
    clean = re.sub(r'^[\s\-\:\,\.]+', '', clean).strip()
    if len(clean) > 2 and not any(kw in clean.lower() for kw in ["sup", "am", "hr", "hrbp", "chup", "anh", "hinh", "camera", "gui"]):
        return clean

    return "Bưu cục NTB"

def detect_target_member(sender_name: str, text: str):
    """
    Nhận diện nhân sự (3 SUP/AM + 6 HR) chặt chẽ, CHỐNG NHẢY NHẦM TUYỆT ĐỐI:
    1. So khớp người gửi trước (Display Name trên GTalk).
    2. So khớp từ khóa / Họ tên / Mã NV / Chức danh trong nội dung tin nhắn.
    3. TUYỆT ĐỐI KHÔNG FALLBACK THEO KHO (tránh tình trạng người này đến kho người khác bị gán nhầm).
    """
    norm_sender = remove_accents(sender_name)
    norm_text = remove_accents(text)

    # 1. So khớp người gửi trên GTalk (nếu có Display Name)
    if norm_sender:
        for m_id, m_info in MEMBERS.items():
            for alias in m_info["aliases"]:
                n_alias = remove_accents(alias)
                if len(n_alias) <= 3:
                    if re.search(r'\b' + re.escape(n_alias) + r'\b', norm_sender):
                        return m_id, m_info
                else:
                    if n_alias in norm_sender:
                        return m_id, m_info

    # 2. So khớp từ khóa trong text theo độ dài giảm dần (ưu tiên cụm từ dài/chính xác nhất)
    candidates = []
    for m_id, m_info in MEMBERS.items():
        for alias in m_info["aliases"]:
            candidates.append((len(alias), alias, m_id, m_info))

    candidates.sort(key=lambda x: x[0], reverse=True)

    for length, alias, m_id, m_info in candidates:
        n_alias = remove_accents(alias)
        target_text = norm_text

        # Loại bỏ các từ gây nhầm lẫn nếu có
        for avoid in m_info.get("avoid_confusions", []):
            target_text = re.sub(r'\b' + re.escape(remove_accents(avoid)) + r'\b', ' ', target_text)

        pattern = r'\b' + re.escape(n_alias) + r'\b'
        if re.search(pattern, target_text):
            return m_id, m_info

    # KHÔNG FALLBACK THEO KHO
    return None, None

# Hàm bọc để giữ tính tương thích
def detect_target_sup(sender_name: str, text: str):
    return detect_target_member(sender_name, text)

# ─── XỬ LÝ WEBHOOK TỪ GTALK ──────────────────────────────────

def handle_sup_webhook(data: dict):
    """
    Xử lý payload khi có tin nhắn/ảnh gửi vào Group Check-in (2099483038556782592).
    """
    now = get_vn_now()
    now_hm = now.strftime("%H:%M:%S")
    today_display = now.strftime("%d/%m/%Y")

    sender_obj = data.get("sender") or {}
    sender_name = sender_obj.get("displayName") or sender_obj.get("name") or data.get("senderName") or ""

    # Trích xuất caption text và ảnh
    text, has_photo, photo_file_id = parse_gtalk_message_payload(data)
    norm_txt = remove_accents(text)

    # Bỏ qua nếu là tin do chính bot gửi
    if any(kw in text for kw in ["XÁC NHẬN CHECK-IN", "NHẮC NHỞ", "CẢNH BÁO 08:00", "BÁO CÁO CHECK-IN"]):
        return {"status": "skipped", "reason": "bot message"}

    # ── 1. KIỂM TRA TIN BÁO NGHỈ / CÔNG TÁC / MIỄN CHECK-IN ──
    is_excuse = bool(re.search(r'(?i)#?off\b|xin\s*nghi|nghi\s*phep|bao\s*vang|cong\s*tac|di\s*hop|xin\s*mien|mien\s*check\s*in', norm_txt))
    if is_excuse:
        m_id, m_info = detect_target_member(sender_name, text)
        if not m_id:
            return {"status": "ignored", "reason": "cannot identify member for excuse"}

        # Trích xuất lý do
        reason = ""
        m_reason = re.search(r'(?i)(?:vì\s*lý\s*do|vi\s*ly\s*do|lý\s*do|ly\s*do)[:\s]+(.+)', text)
        if m_reason:
            reason = m_reason.group(1).strip()
        else:
            cleaned = text
            for alias in m_info["aliases"]:
                cleaned = re.sub(r'(?i)\b' + re.escape(alias) + r'\b', '', cleaned)
            cleaned = re.sub(r'(?i)\b(am|sup|hr|hrbp)\b', '', cleaned)
            cleaned = re.sub(r'(?i)#?off[:\s]*|xin\s*mi[eễ]n\s*check\s*in[:\s]*|mi[eễ]n\s*check\s*in[:\s]*|xin\s*ngh[iỉ]\s*ph[eé]p[:\s]*|xin\s*ngh[iỉ][:\s]*|ngh[iỉ]\s*ph[eé]p[:\s]*|b[aá]o\s*v[aắ]ng[:\s]*', '', cleaned)
            cleaned = re.sub(r'^[\s\-\:\,\.]+', '', cleaned).strip()
            reason = cleaned

        if not reason or len(reason) < 2:
            reason = "Xin miễn check-in / Nghỉ phép"

        try:
            ws, records = get_today_records(today_display)
            row_idx = records[m_id]["row_idx"]
            if row_idx:
                update_sup_record(ws, row_idx, status="🏖️ Miễn check-in", note=reason, details=f"Báo lúc {now_hm}")

            reply = (
                f"📝 <b>XÁC NHẬN MIỄN CHECK-IN / NGHỈ PHÉP</b>\n"
                f"👤 {m_info['role']}: <b>{m_info['full_name']}</b>\n"
                f"📌 Trạng thái: <b>🏖️ Miễn check-in / Xin phép</b>\n"
                f"💬 Lý do: <i>{reason}</i>\n"
                f"⏰ Thời gian ghi nhận: {now_hm} - {today_display}"
            )
            send_gtalk_message(reply)
            return {"status": "recorded_excuse", "member": m_info["full_name"], "reason": reason}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    # ── 2. KIỂM TRA TIN CHECK-IN (Có từ khóa checkin hoặc gửi kèm ảnh) ──
    is_checkin = bool(re.search(r'(?i)#?check\s*in|diem\s*danh|kct|ktc|buu\s*cuc|kho|bt|kh|dn|dt|bl|pt|la\s*gi', text) or has_photo)
    if not is_checkin:
        return {"status": "ignored", "reason": "not checkin message"}

    m_id, m_info = detect_target_member(sender_name, text)
    if not m_id:
        return {"status": "ignored", "reason": "cannot identify member for checkin"}

    location = detect_location(text)

    # Đánh giá đúng giờ hay trễ (Cut-off là 08:00:59)
    cutoff_time = now.replace(hour=8, minute=0, second=59, microsecond=0)
    is_on_time = (now <= cutoff_time)
    status_label = "✅ Đúng giờ" if is_on_time else "⚠️ Gửi bù (Trễ)"

    try:
        ws, records = get_today_records(today_display)
        row_idx = records[m_id]["row_idx"]
        if row_idx:
            update_sup_record(
                ws,
                row_idx,
                status=status_label,
                checkin_time=now_hm,
                location=location,
                note="Đúng giờ" if is_on_time else "Gửi bù sau 08:00",
                details=""
            )

        reply = (
            f"✅ <b>XÁC NHẬN CHECK-IN ĐẦU NGÀY</b>\n"
            f"👤 {m_info['role']}: <b>{m_info['full_name']}</b>\n"
            f"📍 Địa điểm: <b>{location}</b>\n"
            f"⏰ Thời gian: <b>{now_hm}</b> ({status_label})\n"
            f"🔗 <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Xem bảng theo dõi Check-in</a>"
        )
        send_gtalk_message(reply)
        return {"status": "recorded_checkin", "member": m_info["full_name"], "status_label": status_label}
    except Exception as e:
        return {"status": "error", "message": str(e)}

# ─── KỊCH BẢN HẸN GIỜ (07:55, 08:00, 11:00) ──────────────────

last_job_executed = {}

def job_remind_0755():
    """07:55 - Nhắc nhở các nhân sự chưa check-in (còn 5 phút trước hạn 08:00)."""
    now = get_vn_now()
    today_display = now.strftime("%d/%m/%Y")
    try:
        ws, records = get_today_records(today_display)
        missing_members = []
        for m_id, rec in records.items():
            status = rec.get("status", "")
            if "Chưa check-in" in status and not any(kw in status for kw in ["Nghỉ", "Vắng", "Đúng giờ", "Gửi bù", "Miễn"]):
                missing_members.append(f"{rec['name']} ({rec['role']})")

        if not missing_members:
            print(f"[{now.strftime('%H:%M:%S')}] 07:55 - Tất cả nhân sự đã check-in hoặc xin nghỉ. Không cần nhắc.")
            return

        lines = [
            "⏰ <b>[NHẮC NHỞ CHECK-IN 07:55]</b>",
            "<i>Còn 5 phút trước giờ chốt điểm danh đầu ca sáng (08:00)!</i>",
            "",
            "📌 <b>Danh sách nhân sự chưa gửi check-in TimestampCam:</b>"
        ]
        for idx, item in enumerate(missing_members, 1):
            lines.append(f"{idx}. <b>{item}</b>")

        lines.extend([
            "",
            "👉 <i>Yêu cầu các AM/SUP và HR khẩn trương gửi ảnh TimestampCam kèm vị trí bưu cục/kho vào nhóm nhé!</i>",
            f"🔗 <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Bảng theo dõi Check-in</a>"
        ])
        send_gtalk_message("\n".join(lines))
        print(f"[{now.strftime('%H:%M:%S')}] Đã gửi nhắc nhở 07:55 cho {len(missing_members)} người.")
    except Exception as e:
        print(f"[ERROR job_remind_0755]: {e}")

def job_cutoff_0800():
    """08:00 - Cảnh báo quá hạn và yêu cầu nhân sự chưa check-in gửi bù ngay."""
    now = get_vn_now()
    today_display = now.strftime("%d/%m/%Y")
    try:
        ws, records = get_today_records(today_display)
        overdue_members = []
        for m_id, rec in records.items():
            status = rec.get("status", "")
            if "Chưa check-in" in status:
                overdue_members.append(f"{rec['name']} ({rec['role']})")
                if rec["row_idx"]:
                    update_sup_record(ws, rec["row_idx"], status="❌ Quá hạn (Chờ bù)", note="Quá 08:00 chưa gửi")

        if not overdue_members:
            print(f"[{now.strftime('%H:%M:%S')}] 08:00 - Tất cả nhân sự đã hoàn thành check-in đúng giờ hoặc xin phép.")
            return

        lines = [
            "⚠️ <b>[CẢNH BÁO QUÁ HẠN CHECK-IN 08:00]</b>",
            "<i>Đã quá 08:00 sáng - Hệ thống đã chốt danh sách check-in đúng giờ!</i>",
            "",
            "❌ <b>Các nhân sự sau chưa check-in đúng hạn:</b>"
        ]
        for idx, item in enumerate(overdue_members, 1):
            lines.append(f"{idx}. <b>{item}</b>")

        lines.extend([
            "",
            "👉 <b>Yêu cầu các bạn khẩn trương chụp ảnh TimestampCam gửi bù ngay vào nhóm!</b>",
            f"🔗 <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Bảng theo dõi Check-in</a>"
        ])
        send_gtalk_message("\n".join(lines))
        print(f"[{now.strftime('%H:%M:%S')}] Đã gửi cảnh báo quá hạn 08:00 cho {len(overdue_members)} người.")
    except Exception as e:
        print(f"[ERROR job_cutoff_0800]: {e}")

def job_recap_1100():
    """11:00 - Báo cáo tổng hợp tình hình điểm danh ca sáng của SUP và HR."""
    now = get_vn_now()
    today_display = now.strftime("%d/%m/%Y")
    try:
        ws, records = get_today_records(today_display)
        lines = [
            f"📊 <b>BÁO CÁO CHECK-IN ĐẦU NGÀY - {today_display} (11:00)</b>",
            "─────────────────────────────"
        ]

        sup_items = []
        hr_items = []

        for m_id, rec in records.items():
            status = rec.get("status", "Chưa check-in")
            time_str = rec.get("time", "")
            loc_str = rec.get("location", "")
            note_str = rec.get("note", "")

            icon = "✅"
            if "Nghỉ" in status or "Vắng" in status or "Miễn" in status:
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
            line_str = f"• <b>{rec['name']}:</b> {icon} {status}{detail_text}"

            if rec.get("role") == "SUP/AM":
                sup_items.append(line_str)
            else:
                hr_items.append(line_str)

        lines.append("👔 <b>KHỐI VẬN HÀNH (AM/SUP):</b>")
        lines.extend(sup_items if sup_items else ["(Không có dữ liệu)"])
        lines.append("")
        lines.append("💼 <b>KHỐI NHÂN SỰ (HR/HRBP):</b>")
        lines.extend(hr_items if hr_items else ["(Không có dữ liệu)"])

        lines.extend([
            "─────────────────────────────",
            f"🔗 <b>Chi tiết:</b> <a href='https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}'>Google Sheet Điểm Danh</a>"
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
        print(f"⏰ [SUP/HR SCHEDULER] Kích hoạt {job_key_name} lúc {hm}...")
        try:
            job_func()
        except Exception as e:
            print(f"❌ Lỗi chạy {job_key_name}: {e}")
