# -*- coding: utf-8 -*-
"""
Hệ thống Bot Điểm danh & Tính phạt Cut-off Báo Cáo AM Vùng NTB
Tự động hóa ghi nhận báo cáo qua GTalk Webhook, nhắc nhở cut-off và chốt sổ Daily N+1.
"""

import os
import sys
import re
import time
import json
import sqlite3
import argparse
import threading
import unicodedata
from datetime import datetime, date, timedelta, timezone

VN_TZ = timezone(timedelta(hours=7))

def get_vn_now() -> datetime:
    """Trả về thời gian hiện tại chuẩn theo múi giờ Việt Nam (GMT+7)."""
    return datetime.now(VN_TZ).replace(tzinfo=None)

def get_vn_today() -> date:
    """Trả về ngày hiện tại chuẩn theo múi giờ Việt Nam (GMT+7)."""
    return get_vn_now().date()

import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from flask import Flask, request, jsonify

# ─── THIẾT LẬP ENCODING UTF-8 ─────────────────────────────────
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "am_config.json")
DB_PATH = os.path.join(BASE_DIR, "diem_danh.db")

app = Flask(__name__)

# ─── NẠP CẤU HÌNH ─────────────────────────────────────────────
def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

# ─── DATABASE SQLITE ──────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        cursor = conn.cursor()
        # Bảng 1: Lưu nguyên văn mọi tin nhắn đến (Audit Trail)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS raw_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            message_id TEXT,
            channel_id TEXT,
            sender_id TEXT,
            sender_name TEXT,
            raw_text TEXT,
            received_at DATETIME,
            detected_am_id TEXT,
            detected_am_name TEXT,
            detected_milestone INTEGER,
            is_valid INTEGER,
            error_reason TEXT,
            evaluation_status TEXT,
            late_minutes INTEGER,
            penalty_amount INTEGER
        )
        """)
        # Bảng 2: Bản ghi điểm danh theo ngày, từng AM, từng mốc
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS attendance_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            am_id TEXT,
            am_name TEXT,
            milestone_id INTEGER,
            submitted_at DATETIME,
            status TEXT,
            late_minutes INTEGER,
            penalty_amount INTEGER,
            raw_message_id INTEGER,
            updated_at DATETIME,
            UNIQUE(date, am_id, milestone_id)
        )
        """)

        # Bảng 3: Lưu các trường hợp xin phép trễ / off phép (Ad-hoc excuses)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS excuses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            am_id TEXT,
            am_name TEXT,
            milestone_id INTEGER,
            reason TEXT,
            raw_text TEXT,
            created_at DATETIME,
            excuse_type TEXT DEFAULT 'LATE_PERMIT',
            UNIQUE(date, am_id, milestone_id)
        )
        """)
        try:
            cursor.execute("ALTER TABLE excuses ADD COLUMN excuse_type TEXT DEFAULT 'LATE_PERMIT'")
        except Exception:
            pass
        conn.commit()

# ─── BỘ BÓC TÁCH (PARSER) ─────────────────────────────────────
def normalize_text(text: str) -> str:
    if not text:
        return ""
    s = unicodedata.normalize("NFC", str(text)).strip().lower()
    return re.sub(r'[\r\t]', ' ', s)

def remove_accents(input_str: str) -> str:
    if not input_str:
        return ""
    s = str(input_str).replace('đ', 'd').replace('Đ', 'd')
    nfkd_form = unicodedata.normalize('NFKD', s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

class AttendanceParser:
    def __init__(self, config):
        self.config = config
        self.ams = [am for am in config.get("ams", []) if am.get("is_active", True)]
        self.milestones = config.get("milestones", {})

    def detect_am(self, text: str, sender_name: str = ""):
        norm_txt = normalize_text(text)
        no_acc_txt = remove_accents(norm_txt)
        norm_sender = normalize_text(sender_name)
        no_acc_sender = remove_accents(norm_sender)

        # ── LOẠI TRỪ ADMIN / QUẢN LÝ / TRỢ LÝ (Không phải AM báo cáo) ──
        EXCLUDED_SENDERS = [
            "nguyễn thị thanh thủy", "nguyen thi thanh thuy", "thanh thủy admin", "thuy admin"
        ]
        is_admin_sender = any(exc in norm_sender or exc in no_acc_sender for exc in EXCLUDED_SENDERS)

        # ── TIER 1: Header / Khai báo trực tiếp (Ưu tiên số 1) ──
        # Bắt mẫu: 'Khu vực AM DuyPĐ', 'AM LongNT', 'KV AM TienTH', 'AM: Nguyễn Thanh Long', 'Báo cáo ... AM NgaHB', 'AM Duy xin miễn...'
        header_m = re.search(r'(?:khu\s+vực\s+|kv\s+)?am\s*[:\-\s]\s*([^\n\r,\:\;]+)', norm_txt)
        if header_m:
            header_segment = header_m.group(1).strip()
            no_acc_header = remove_accents(header_segment)
            best_len = 0
            best_am = None
            for am in self.ams:
                for alias in am.get("aliases", []):
                    na = normalize_text(alias)
                    noa = remove_accents(na)
                    # Khớp ở đầu header segment hoặc dưới dạng từ độc lập
                    p = r'^(?:\b|_)' + re.escape(na) + r'(?:\b|_)'
                    p_noa = r'^(?:\b|_)' + re.escape(noa) + r'(?:\b|_)'
                    if re.search(p, header_segment) or re.search(p_noa, no_acc_header):
                        if len(na) > best_len:
                            best_len = len(na)
                            best_am = am
            if best_am:
                return best_am

        # Nếu người gửi là Admin mà không khai báo rõ ràng tên AM ở Tier 1 -> Bỏ qua, không nhận diện
        if is_admin_sender:
            return None

        # ── TIER 2: Tên người gửi GTalk (sender_name) ──
        if sender_name:
            best_len = 0
            best_am = None
            for am in self.ams:
                for alias in am.get("aliases", []):
                    na = normalize_text(alias)
                    noa = remove_accents(na)
                    p = r'(?:\b|_)' + re.escape(na) + r'(?:\b|_)'
                    p_noa = r'(?:\b|_)' + re.escape(noa) + r'(?:\b|_)'
                    if re.search(p, norm_sender) or re.search(p_noa, no_acc_sender):
                        if len(na) > best_len:
                            best_len = len(na)
                            best_am = am
            if best_am:
                return best_am

        # ── TIER 3 & 4: Tìm kiếm trong nội dung tin nhắn có lọc địa danh (Geo-Guards) ──
        geo_ambiguous = {
            'khánh', 'khanh', 'long', 'lâm', 'lam', 'bình', 'binh',
            'hải', 'hai', 'sơn', 'son', 'đông', 'dong', 'nam', 'bắc', 'bac',
            'thủy', 'thuy', 'an', 'hòa', 'hoa', 'linh', 'thơ', 'tho', 'thu', 'thư'
        }
        geo_prev = r'(?:diên|dien|kho|\(kho\)|bưu cục|buu cuc|bc|tỉnh|tinh|tp|thành phố|thanh pho|đại|dai|phú|phu|cam|đắk|dak|hạ|ha|phước|phuoc|di|quảng|quang|phúc|phuc|yên|yen|đức|duc)\s+$'
        geo_next = r'^\s+(?:hòa|hoa|vĩnh|vinh|sơn|son|lâm|lam|điền|dien|nam|bắc|bac|đông|dong|tây|tay|thuận|thuan|định|dinh|trang|nghĩa|nghia|thọ|tho)'

        best_match = None
        longest_alias_len = 0

        for am in self.ams:
            for alias in am.get("aliases", []):
                na = normalize_text(alias)
                noa = remove_accents(na)
                words = na.split()

                p = r'(?:\b|_)' + re.escape(na) + r'(?:\b|_)'
                p_noa = r'(?:\b|_)' + re.escape(noa) + r'(?:\b|_)'

                # Từ đơn trong nội dung tin nhắn bắt buộc phải đúng dấu tiếng Việt để tránh đụng từ (ví dụ: 'thọ' nhầm 'tho')
                if len(words) == 1 and len(na) <= 5:
                    m_obj = re.search(p, norm_txt)
                else:
                    m_obj = re.search(p, norm_txt) or re.search(p_noa, no_acc_txt)

                if m_obj:
                    # Kiểm tra xem có phải tên bưu cục/địa danh tỉnh huyện trùng với tên AM hay không
                    if na in geo_ambiguous or noa in geo_ambiguous:
                        start, end = m_obj.start(), m_obj.end()
                        before = norm_txt[:start]
                        after = norm_txt[end:]
                        if re.search(geo_prev, before) or re.search(geo_next, after):
                            continue
                    if len(na) > longest_alias_len:
                        longest_alias_len = len(na)
                        best_match = am

        return best_match

    def detect_report_type(self, text: str, channel_id: str = "", submit_time: datetime = None):
        if submit_time is None:
            submit_time = get_vn_now()

        norm_txt = normalize_text(text)
        no_accent_txt = remove_accents(norm_txt)

        # Nếu gửi vào Group B (Group Báo cáo Điểm nóng) -> Mặc định là Mốc 5
        gtalk_cfg = self.config.get("gtalk", {})
        group_b_id = str(self.config.get("channel_id_group_b") or gtalk_cfg.get("channel_id_group_b") or "2097270568973508608")
        if channel_id and str(channel_id) == group_b_id:
            return 5, "BC Điểm nóng (GTC <50%)"

        m1_keywords = ["tổng hợp đầu ngày", "dau ngay", "gtc ngày n-1", "tỷ lệ gtc", "nvpttt"]
        m4_keywords = [
            "ltc tts", "đơn ltc", "luân chuyển tts", "luan chuyen tts",
            "lc trước 23h", "lc truoc 23h", "lc được trước 23h", "lc duoc truoc 23h",
            "không lc được", "khong lc duoc", "không lc trước", "khong lc truoc",
            "không lc", "khong lc", "đơn tts không lc", "tts không lc", "tts khong lc"
        ]
        m2_keywords = [
            "gán giaotts", "gan giaotts", "gán tts ca 1", "gan tts ca 1", "ca 1",
            "trước 9h", "truoc 9h", "trước 10h", "truoc 10h", "trước 11h", "truoc 11h",
            "trước 9h00", "truoc 9h00", "trước 10h00", "truoc 10h00", "trước 11h00", "truoc 11h00"
        ]
        m3_keywords = [
            "gán tts ca 2", "gan tts ca 2", "ca 2",
            "trước: 16h", "trước 16h", "truoc 16h", "trước 16h00", "truoc 16h00",
            "trước: 15h", "trước 15h", "truoc 15h", "trước 15h00", "truoc 15h00"
        ]
        m5_keywords = [
            "điểm nóng", "diem nong", "gtc <50%", "gtc < 50%", "gtc dưới 50", "gtc duoi 50", "dưới 50%", "duoi 50%",
            "xuất hàng xong", "xuat hang xong", "thời gian xuất hàng", "time xuất hàng",
            "nhân viên đi làm", "nhan vien di lam", "nv đi làm", "nv di lam",
            "bưu cục :", "bưu cục:", "bưu cục (", "buu cuc (", "bưu cục -", "buu cuc -"
        ]

        if any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m1_keywords):
            return 1, "Tổng hợp đầu ngày"
        if any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m4_keywords):
            return 4, "LTC TTS"
        if re.search(r'\blc\b.*?(?:23h|trước|truoc)', norm_txt) or re.search(r'\btts\b.*?\blc\b', norm_txt):
            return 4, "LTC TTS"
        has_m2 = any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m2_keywords) or bool(re.search(r'\b(?:trước\s*)?(?:9h|10h|11h)\b', norm_txt))
        has_m3 = any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m3_keywords) or bool(re.search(r'\b(?:trước\s*)?(?:15h|16h)\b', norm_txt))

        # Nếu có từ khóa của cả Ca 1 và Ca 2 (ví dụ AM copy nhầm mẫu có dòng 'trước 16h' vào báo cáo ca 1 'trước 11h')
        if has_m2 and has_m3:
            if submit_time.hour < 13:
                return 2, "Gán TTS ca 1"
            else:
                return 3, "Gán TTS ca 2"

        # Nếu gửi buổi sáng (< 13:00): Ưu tiên Ca 1 (Mốc 2, cut-off 11:00)
        if submit_time.hour < 13:
            if has_m2:
                return 2, "Gán TTS ca 1"
            if has_m3:
                return 3, "Gán TTS ca 2"
        else:
            # Nếu gửi buổi chiều (>= 13:00): Ưu tiên Ca 2 (Mốc 3, cut-off 16:00)
            if has_m3:
                return 3, "Gán TTS ca 2"
            if has_m2:
                return 2, "Gán TTS ca 1"

        if re.search(r'tồn\s*/\s*tổng|ton\s*/\s*tong|tồn\s*:\s*\d+|ton\s*:\s*\d+', norm_txt) or any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m5_keywords):
            return 5, "BC Điểm nóng (GTC <50%)"

        if "gán tts" in norm_txt or "gan tts" in no_accent_txt:
            # Nếu gửi buổi sáng (< 13h) -> Mặc định là Ca 1 trừ khi chỉ định rõ ràng ca 2 / 16h
            if submit_time.hour < 13:
                if any(k in norm_txt for k in ["ca 2", "15h", "16h", "chiều", "chieu"]):
                    return 3, "Gán TTS ca 2"
                return 2, "Gán TTS ca 1"
            else:
                # Gửi buổi chiều (>= 13h) -> Mặc định là Ca 2
                if any(k in norm_txt for k in ["ca 1", "11h", "10h", "9h"]):
                    return 2, "Gán TTS ca 1"
                return 3, "Gán TTS ca 2"

        return None, "Không xác định"

    def validate_content(self, report_type: int, text: str):
        norm_txt = normalize_text(text)
        if report_type == 1:
            has_gtc = "gtc" in norm_txt
            has_ns = "nhân sự" in norm_txt or "nvpttt" in norm_txt or "nhan su" in norm_txt
            if not (has_gtc or has_ns):
                return False, "Thiếu thông tin GTC hoặc Nhân sự"
        elif report_type in (2, 3):
            if "tts" not in norm_txt and "đơn" not in norm_txt and "don" not in norm_txt:
                return False, "Thiếu số liệu đơn TTS"
        elif report_type == 4:
            if "ltc" not in norm_txt and "lc" not in norm_txt:
                return False, "Thiếu số liệu luân chuyển LTC"
        elif report_type == 5:
            has_ton = any(k in norm_txt for k in ["tồn", "ton"])
            has_xuat = any(k in norm_txt for k in ["xuất hàng", "xuat hang", "time", "thời gian"])
            has_nv = any(k in norm_txt for k in ["nhân viên", "nhan vien", "nv", "đi làm", "di lam", "hiện tại"])
            if not (has_ton or has_xuat or has_nv or "bưu cục" in norm_txt):
                return False, "Thiếu thông tin bưu cục hoặc tồn/nhân sự/xuất hàng"

        return True, "Hợp lệ"


def generate_hub_variants(raw_hub: str):
    """
    Sinh các biến thể tên bưu cục để bắt được linh hoạt khi AM viết tắt/rút gọn.
    Ví dụ: '(LDO) Tân Hà Lâm Hà' -> ['(LDO) Tân Hà Lâm Hà', 'Tân Hà Lâm Hà', 'Tân Hà Lâm', 'Tân Hà']
    """
    clean = re.sub(r'^\([A-Za-z0-9]+\)\s*', '', raw_hub).strip()
    variants = [raw_hub, clean]
    sub_district = re.sub(r'\s*[\-\–]\s*.*$', '', clean)
    sub_lamha = re.sub(r'\s+lâm\s+hà$', '', clean, flags=re.I)
    sub_trunc = re.sub(r'\s+hà$', '', clean, flags=re.I)
    for v in [sub_district, sub_lamha, sub_trunc]:
        if len(v.strip()) >= 4 and v.strip() not in variants:
            variants.append(v.strip())
    return variants

def detect_hubs_in_text(text: str):
    """
    Quét tìm xem trong nội dung tin nhắn có chứa tên bưu cục nào
    thuộc tab 'BC GTC dưới 50' hay không.
    Trả về danh sách các bưu cục tìm thấy: [{'raw_hub': ..., 'clean_hub': ..., 'am': am_dict}]
    """
    try:
        from sync_attendance_sheets import get_sheet_client, load_config
        gc = get_sheet_client()
        sh = gc.open_by_key('147nvGXc2D7UJNJGsWaFjaIZ6FkJDD7Zmevn3lBs-Bl0')
        ws = sh.worksheet('BC GTC dưới 50')
        rows = ws.get_all_values()
        if len(rows) <= 1:
            return []

        config = load_config()
        parser = AttendanceParser(config)

        norm_msg = normalize_text(text)
        no_accent_msg = remove_accents(norm_msg)

        matched = []
        for r in rows[1:]:
            if len(r) >= 2 and r[0] and r[1]:
                # Bỏ qua bưu cục đang bàn giao / miễn trừ
                full_row_text = " ".join(r).lower()
                no_acc_row = remove_accents(full_row_text)
                if any(kw in full_row_text or kw in no_acc_row for kw in [
                    "ban giao", "bàn giao", "chuyen giao", "chuyển giao", "mien", "miễn", "loai tru", "loại trừ"
                ]):
                    continue

                raw_hub = r[0].strip()
                am_str = r[1].strip()
                variants = generate_hub_variants(raw_hub)
                for v in variants:
                    nv = normalize_text(v)
                    no_acc_v = remove_accents(nv)
                    if nv in norm_msg or no_acc_v in no_accent_msg:
                        matched_am = parser.detect_am(am_str, am_str)
                        matched.append({
                            'raw_hub': raw_hub,
                            'clean_hub': re.sub(r'^\([A-Za-z0-9]+\)\s*', '', raw_hub).strip(),
                            'am_name_sheet': am_str,
                            'am': matched_am
                        })
                        break
        return matched
    except Exception as e:
        print(f"⚠️ Lỗi detect_hubs_in_text: {e}")
        return []

# ─── TÍNH PHẠT & GHI NHẬN ─────────────────────────────────────
def detect_excuse_request(raw_text: str, sender_name: str = "", dt: datetime = None, channel_id: str = None):
    """
    Tự động phát hiện khi AM nhắn tin xin phép báo cáo trễ / xin off phép / xin miễn báo cáo.
    Bảo đảm không nhận diện nhầm khi AM gửi báo cáo nghiệp vụ (chứa ghi chú nhân sự nghỉ).
    """
    if dt is None:
        dt = get_vn_now()

    norm_txt = normalize_text(raw_text)
    no_accent_txt = remove_accents(norm_txt)

    # 1. BẢO VỆ TUYỆT ĐỐI: Bỏ qua ngay lập tức nếu tin nhắn là báo cáo vận hành
    report_keywords = [
        "tổng hợp đầu ngày", "dau ngay", "đầu ngày", "tỷ lệ gtc", "ty le gtc", "gtc ngày n-1", "gtc",
        "nvpttt", "leadtime", "đơn giao tồn", "gán giaotts", "gan giaotts", "ltc tts",
        "xuất hàng xong", "xuat hang xong", "thời gian xuất hàng", "tồn / tổng", "ton / tong",
        "bưu cục :", "bưu cục:", "doanh thu sme", "trả hàng:", "tra hang:"
    ]
    report_hits = sum(1 for k in report_keywords if k in norm_txt or remove_accents(k) in no_accent_txt)
    if report_hits >= 2 or any(k in norm_txt for k in ["tổng hợp đầu ngày", "gán tts ca", "ltc tts", "xuất hàng xong"]):
        return None

    # 2. Kiểm tra từ khóa xin phép của AM (xin trễ, xin miễn báo cáo, xin off)
    excuse_patterns = [
        r'(?:am|kv|em|mình|tôi)?\s*xin\s+(?:phép\s+)?(?:báo\s+cáo\s+)?(?:nộp\s+)?trễ',
        r'(?:am|kv|em|mình|tôi)?\s*xin\s+trễ',
        r'báo\s+cáo\s+trễ',
        r'xin\s+nộp\s+trễ',
        r'(?:am|kv|em|mình|tôi)?\s*xin\s+(?:phép\s+)?(?:miễn|khong|không|ko|k|loại\s+trừ|loai\s+tru)\s+(?:báo\s+cáo|bc|nộp)',
        r'\b(?:miễn|loại\s+trừ|loai\s+tru)\s+(?:báo\s+cáo|bc|nộp)\b',
        r'\b(?:không|khong|ko|k)\s+(?:báo\s+cáo|bc|nộp|làm\s+được|lam\s+duoc)\b',
        r'xin\s+(?:miễn|loại\s+trừ|loai\s+tru)',
        r'xin\s+không\s+nộp',
        r'xin\s+không\s+báo\s+cáo',
        r'xin\s+khong\s+bao\s+cao',
        r'xin\s+k\s+báo\s+cáo',
        r'xin\s+k\s+bc',
        r'không\s+báo\s+cáo\s+được',
        r'khong\s+bao\s+cao\s+duoc',
        r'chưa\s+báo\s+cáo\s+được',
        r'chua\s+bao\s+cao\s+duoc',
        r'xin\s+(?:phép\s+)?(?:đi\s+tuyến|di\s+tuyen)',
        r'(?:am|kv|em|mình|tôi)\s+xin\s+(?:phép\s+)?off',
        r'(?:am|kv|em|mình|tôi)\s+xin\s+(?:phép\s+)?nghỉ',
        r'\bnghỉ\s+phép\b',
        r'\boff\s+phép\b',
        r'(?:khoa|khoá)\s+(?:id|acc|tai\s*khoan|tài\s*khoản|user).*(?:xin|khong|không|ko|k|chua|chưa|miễn|loại)',
        r'(?:ban\s*giao|bàn\s*giao).*(?:xin|khong|không|ko|k|chua|chưa|miễn|loại)'
    ]

    is_excuse = any(re.search(p, norm_txt) or re.search(remove_accents(p), no_accent_txt) for p in excuse_patterns)
    if not is_excuse:
        return None

    config = load_config()
    parser = AttendanceParser(config)
    detected_am = parser.detect_am(raw_text, sender_name)
    if not detected_am and sender_name:
        norm_sender = normalize_text(sender_name)
        no_acc_sender = remove_accents(norm_sender)
        if not any(exc in norm_sender or exc in no_acc_sender for exc in ["nguyễn thị thanh thủy", "nguyen thi thanh thuy", "thanh thủy admin", "thuy admin"]):
            detected_am = parser.detect_am(sender_name, "")
    if not detected_am:
        return None

    gtalk_cfg = config.get("gtalk", {})
    group_b_id = str(config.get("channel_id_group_b") or gtalk_cfg.get("channel_id_group_b") or "2097270568973508608")
    is_group_b = channel_id and str(channel_id) == group_b_id

    # ─── 3. KIỂM TRA TRƯỜNG HỢP ĐẶC BIỆT (ƯU TIÊN CAO NHẤT) ───
    is_id_locked = any(k in norm_txt or remove_accents(k) in no_accent_txt for k in [
        "khoa id", "khoá id", "bi khoa id", "bị khoá id", "khoa acc", "khoá acc",
        "khoa tai khoan", "khoá tài khoản", "khoa user", "khoá user", "chua mo id", "chưa mở id"
    ])
    is_handover = any(k in norm_txt or remove_accents(k) in no_accent_txt for k in [
        "ban giao", "bàn giao", "chua co du lieu", "chưa có dữ liệu",
        "chua xem duoc", "chưa xem được", "chua coi duoc", "chưa coi được",
        "chua cap quyen", "chưa cấp quyền"
    ])

    if is_id_locked:
        milestones = [1, 5, 2, 3, 4]
        scope_label = "Cả ngày (5 Mốc - Khóa ID hệ thống)"
        next_reminder = "Chúc AM sớm mở lại tài khoản để tiếp tục công việc nhé!"
        reason = "Bị khóa ID / Tài khoản hệ thống"
        is_exemption = True
    elif is_handover:
        milestones = [5]
        scope_label = "Mốc 5 (BC Điểm nóng 10h00 - Đang bàn giao bưu cục)"
        next_reminder = "Các mốc vận hành chung khác (1, 2, 3, 4) vẫn báo cáo đúng timeline quy định nhé!"
        reason = "Đang bàn giao bưu cục / Chưa có dữ liệu"
        is_exemption = True
    elif any(k in norm_txt for k in ["cả ngày", "ca ngay", "nguyên ngày", "hôm nay off", "hom nay off"]):
        milestones = [1, 5, 2, 3, 4]
        scope_label = "Cả ngày (Tất cả các mốc)"
        next_reminder = "Chúc AM nghỉ ngơi / công tác tốt nhé!"
        reason = "Nghỉ phép cả ngày"
        is_exemption = True
    elif re.search(r'\b(?:mốc\s*2|moc\s*2|ca\s*1)\b', norm_txt):
        milestones = [2]
        scope_label = "Mốc 2 (Gán TTS Ca 1 11h00)"
        next_reminder = "Các mốc khác vẫn báo cáo đúng timeline quy định nhé!"
        reason = None
        is_exemption = False
    elif re.search(r'\b(?:mốc\s*3|moc\s*3|ca\s*2)\b', norm_txt):
        milestones = [3]
        scope_label = "Mốc 3 (Gán TTS Ca 2 16h00)"
        next_reminder = "Mốc tối (20h00 - LTC TTS) vẫn báo cáo đúng timeline quy định nhé!"
        reason = None
        is_exemption = False
    elif re.search(r'\b(?:mốc\s*5|moc\s*5|điểm\s*nóng|diem\s*nong)\b', norm_txt) or is_group_b:
        milestones = [5]
        scope_label = "Mốc 5 (BC Điểm nóng 10h00)"
        next_reminder = "Các mốc khác vẫn báo cáo đúng timeline quy định nhé!"
        reason = None
        is_exemption = False
    elif re.search(r'\b(?:mốc\s*1|moc\s*1|đầu\s*ngày|dau\s*ngay)\b', norm_txt):
        milestones = [1]
        scope_label = "Mốc 1 (Tổng hợp đầu ngày 08h00)"
        next_reminder = "Các mốc tiếp theo vẫn báo cáo đúng timeline quy định nhé!"
        reason = None
        is_exemption = False
    elif re.search(r'\b(?:mốc\s*4|moc\s*4|ltc|luân\s*chuyển)\b', norm_txt):
        milestones = [4]
        scope_label = "Mốc 4 (LTC TTS 20h00)"
        next_reminder = "Nhờ AM lưu ý các ca tiếp theo nhé!"
        reason = None
        is_exemption = False
    elif re.search(r'\b(?:chiều|chieu|16h(?:00)?)\b', norm_txt):
        milestones = [3]
        scope_label = "Ca chiều (Mốc 3 - Gán TTS Ca 2 16h00)"
        next_reminder = "Mốc tối (20h00 - LTC TTS) vẫn báo cáo đúng timeline quy định nhé!"
        reason = None
        is_exemption = False
    elif re.search(r'\b(?:ca\s*tối|ca\s*toi|20h(?:00)?)\b', norm_txt):
        milestones = [4]
        scope_label = "Ca tối (Mốc 4 - LTC TTS 20h00)"
        next_reminder = "Nhờ AM nộp bù trước khi kết thúc ca làm việc nhé!"
        reason = None
        is_exemption = False
    elif re.search(r'\b(?:sáng|sang|11h(?:00)?|9h(?:00)?)\b', norm_txt):
        milestones = [1, 5, 2]
        scope_label = "Ca sáng (Mốc 1 - Đầu ngày, Mốc 5 - Điểm nóng, Mốc 2 - Gán TTS)"
        next_reminder = "Ca chiều (16h00) và Ca tối (20h00) vẫn báo cáo đúng timeline quy định nhé!"
        reason = None
        is_exemption = False
    else:
        cur_hour = dt.hour
        if cur_hour < 12:
            milestones = [1, 5, 2]
            scope_label = "Ca sáng (Mốc 1 - Đầu ngày, Mốc 5 - Điểm nóng, Mốc 2 - Gán TTS)"
            next_reminder = "Ca chiều (16h00) và Ca tối (20h00) vẫn báo cáo đúng timeline quy định nhé!"
        elif 12 <= cur_hour < 18:
            milestones = [3]
            scope_label = "Ca chiều (Mốc 3 - Gán TTS Ca 2 16h00)"
            next_reminder = "Mốc tối (20h00 - LTC TTS) vẫn báo cáo đúng timeline quy định nhé!"
        else:
            milestones = [4]
            scope_label = "Ca tối (Mốc 4 - LTC TTS 20h00)"
            next_reminder = "Nhờ AM nộp bù trước khi hết ca nhé!"
        reason = None
        is_exemption = False

    if reason is None:
        reason = "Bận việc đột xuất / Đi tuyến"
        for cue in ["lý do", "ly do", "chưa vô được", "chưa vào được", "vì", "vi", "hẹn", "hen", "gặp", "gap", "bận", "ban", "đi tuyến", "di tuyen", "do"]:
            if f" {cue} " in f" {norm_txt} ":
                idx = norm_txt.find(cue)
                reason_part = raw_text[idx + len(cue):].strip().lstrip(":").strip()
                if reason_part:
                    reason = reason_part[:50].split("\n")[0].strip()
                break

    if not is_exemption:
        is_exemption = any(k in norm_txt for k in [
            "miễn", "mien", "không báo cáo", "khong bao cao", "ko báo cáo", "ko bao cao",
            "không nộp", "khong nop", "k nộp", "k báo cáo", "k bc", "cả ngày", "ca ngay",
            "nguyên ngày", "off", "nghỉ", "nghi", "loại trừ", "loai tru", "không làm được",
            "khong lam duoc", "k làm được", "k lam duoc", "đi tuyến", "di tuyen"
        ])
    excuse_type = "EXEMPTION" if is_exemption else "LATE_PERMIT"

    today_str = dt.strftime("%Y-%m-%d")
    with get_db() as conn:
        cur = conn.cursor()
        for m_id in milestones:
            cur.execute("""
            INSERT OR REPLACE INTO excuses (date, am_id, am_name, milestone_id, reason, raw_text, created_at, excuse_type)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                today_str, detected_am["id"], detected_am["full_name"], m_id,
                reason, raw_text, dt.strftime("%Y-%m-%d %H:%M:%S"), excuse_type
            ))
            if is_exemption:
                cur.execute("""
                INSERT INTO attendance_records (date, am_id, am_name, milestone_id, submitted_at, status, late_minutes, penalty_amount, updated_at)
                VALUES (?, ?, ?, ?, ?, 'EXEMPT', 0, 0, CURRENT_TIMESTAMP)
                ON CONFLICT(date, am_id, milestone_id) DO UPDATE SET
                    status = 'EXEMPT',
                    late_minutes = 0,
                    penalty_amount = 0,
                    updated_at = CURRENT_TIMESTAMP
                """, (today_str, detected_am["id"], detected_am["full_name"], m_id, dt.strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()

    if is_exemption:
        action_title = "XÁC NHẬN GHI NHẬN MIỄN BÁO CÁO"
        action_note = "Đã ghi nhận miễn báo cáo mốc này (0đ phạt, không cần nộp bù)."
    else:
        action_title = "XÁC NHẬN GHI NHẬN XIN PHÉP BÁO CÁO TRỄ"
        action_note = "Nhờ AM nộp bù trong vòng 30 phút. Nếu nhắm trễ trên 30 phút, AM vui lòng nhắn xin miễn báo cáo mốc này để tránh sai lệch dữ liệu gán ca nhé!"

    reply_msg = (
        f"📝 <b>{action_title}</b>\n"
        f"👤 <b>AM:</b> {detected_am['full_name']}\n"
        f"🕒 <b>Phạm vi áp dụng:</b> {scope_label}\n"
        f"📌 <b>Lý do ghi nhận:</b> {reason}\n"
        f"👉 <i>Lưu ý: {action_note} {next_reminder}</i>"
    )

    return {
        "am": detected_am,
        "milestones": milestones,
        "scope_label": scope_label,
        "reason": reason,
        "reply_msg": reply_msg
    }

def evaluate_submission(dt: datetime, milestone_id: int, is_valid: bool, config: dict, am_id: str = None):
    if not is_valid:
        return "INVALID", 0, config["fines"]["invalid"], "Báo cáo sai cấu trúc/thiếu tiêu chí"

    ms_cfg = config["milestones"].get(str(milestone_id))
    if not ms_cfg:
        return "ON_TIME", 0, 0, "Hợp lệ"

    cutoff_str = ms_cfg["cutoff"]
    cutoff_hour, cutoff_min = map(int, cutoff_str.split(":"))
    cutoff_dt = dt.replace(hour=cutoff_hour, minute=cutoff_min, second=59, microsecond=999999)

    if dt <= cutoff_dt:
        return "ON_TIME", 0, 0, "Đúng hạn"
    else:
        diff_sec = (dt - cutoff_dt).total_seconds()
        late_min = max(1, int(diff_sec // 60) + 1)

        # Kiểm tra xem AM có xin phép trễ mốc này trong ngày chưa
        has_excuse = False
        excuse_reason = ""
        if am_id:
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                    SELECT reason FROM excuses WHERE date = ? AND am_id = ? AND milestone_id = ?
                    """, (dt.strftime("%Y-%m-%d"), am_id, milestone_id))
                    row = cur.fetchone()
                    if row:
                        has_excuse = True
                        excuse_reason = row['reason']
            except Exception:
                pass

        # Hạn chót nộp bù (Hard deadline theo quy định):
        # - Ca 1 (Mốc 2): Cut-off 11:00, nộp bù đến 12:00. Sau 12:00 tính Không nộp (Phạt 100k)
        # - Ca 2 (Mốc 3): Cut-off 16:00, nộp bù đến 17:00. Sau 17:00 tính Không nộp (Phạt 100k)
        # LƯU Ý: Nếu AM ĐÃ XIN PHÉP TRƯỚC ĐÓ -> Miễn phạt, vẫn ghi nhận hợp lệ (0đ)
        hard_cutoff_map = {
            2: (12, 0),
            3: (17, 0)
        }
        if milestone_id in hard_cutoff_map:
            limit_h, limit_m = hard_cutoff_map[milestone_id]
            limit_dt = dt.replace(hour=limit_h, minute=limit_m, second=0, microsecond=0)
            if dt > limit_dt:
                if has_excuse:
                    return "ON_TIME", late_min, 0, f"Đã xin phép ({excuse_reason}) — Nộp bù sau {limit_h:02d}:{limit_m:02d} (Miễn phạt)"
                return "NOT_SUBMITTED", late_min, config["fines"]["not_submitted"], f"Quá hạn nộp bù (sau {limit_h:02d}:{limit_m:02d}) — Tính Không nộp"

        if has_excuse:
            return "ON_TIME", late_min, 0, f"Đã xin phép ({excuse_reason}) — Miễn phạt 50k"

        return "LATE", late_min, config["fines"]["late"], f"Trễ {late_min} phút"

def record_submission(sender_name, sender_id, raw_text, channel_id, msg_id, submit_time=None):
    if submit_time is None:
        submit_time = get_vn_now()

    # BẢO VỆ TUYỆT ĐỐI: Nếu là tin nhắn xin phép / miễn nộp / off phép -> Bỏ qua để detect_excuse_request xử lý!
    if detect_excuse_request(raw_text, sender_name, submit_time, channel_id):
        return None, "Tin nhắn xin phép / miễn nộp"

    config = load_config()
    parser = AttendanceParser(config)

    detected_am = parser.detect_am(raw_text, sender_name)
    m_id, m_name = parser.detect_report_type(raw_text, channel_id, submit_time)

    # Nếu là Mốc 5 hoặc gửi vào Group B hoặc tìm thấy bưu cục điểm nóng:
    matched_hubs = detect_hubs_in_text(raw_text)
    gtalk_cfg = config.get("gtalk", {})
    group_b_id = str(config.get("channel_id_group_b") or gtalk_cfg.get("channel_id_group_b") or "2097270568973508608")

    if not m_id and matched_hubs:
        m_id = 5
        m_name = "BC Điểm nóng (GTC <50%)"

    if m_id == 5 or (channel_id and str(channel_id) == group_b_id):
        # BẢO VỆ CHỐNG BẮT NHẦM CHAT THƯỜNG TRONG GROUP B:
        # Chỉ xử lý là báo cáo Mốc 5 nếu:
        # 1. Có bưu cục điểm nóng khớp trong tin nhắn (matched_hubs)
        # 2. HOẶC tin nhắn có từ khóa / cấu trúc số liệu báo cáo rõ ràng
        report_cues = ["tồn", "ton", "nvpttt", "xuất hàng", "xuat hang", "điểm nóng", "diem nong", "gtc", "báo cáo", "bao cao", "bưu cục", "buu cuc", "bc "]
        has_report_struct = any(k in raw_text.lower() for k in report_cues)
        if not matched_hubs and not has_report_struct:
            return None, "Chat thông thường trong Group B (không phải báo cáo Mốc 5)"

        m_id = 5
        m_name = "BC Điểm nóng (GTC <50%)"
        # Báo cáo Mốc 5: Bưu cục thuộc quyền AM nào trong tab 'BC GTC dưới 50' thì ưu tiên AM đó
        if matched_hubs and matched_hubs[0].get("am"):
            detected_am = matched_hubs[0]["am"]
        elif matched_hubs and not detected_am:
            detected_am = matched_hubs[0].get("am")

    if not m_id:
        return None, "Không phải mẫu báo cáo 1-5"

    if not detected_am:
        return None, "Không xác định được AM"

    # Kiểm tra nội dung bắt buộc
    is_valid, reason = parser.validate_content(m_id, raw_text)
    status, late_min, penalty, note = evaluate_submission(submit_time, m_id, is_valid, config, detected_am["id"])

    # Lưu DB
    am_id = detected_am["id"]
    am_name = detected_am["full_name"]
    today_str = submit_time.strftime("%Y-%m-%d")

    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
        INSERT INTO raw_messages (
            message_id, channel_id, sender_id, sender_name, raw_text,
            received_at, detected_am_id, detected_am_name, detected_milestone,
            is_valid, error_reason, evaluation_status, late_minutes, penalty_amount
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            msg_id, channel_id, sender_id, sender_name, raw_text,
            submit_time.strftime("%Y-%m-%d %H:%M:%S"), am_id, am_name, m_id,
            1 if is_valid else 0, reason, status, late_min, penalty
        ))
        raw_msg_id = cur.lastrowid

        # Cập nhật attendance_records (nếu đã nhận đúng AM)
        if am_id:
            # Kiểm tra xem tin nhắn này trước đó đã từng ghi nhận vào mốc nào khác chưa (AM sửa tin nhắn cũ đổi mốc)
            cur.execute("""
            SELECT detected_milestone FROM raw_messages
            WHERE message_id = ? AND id != ? AND detected_am_id = ?
            ORDER BY id DESC LIMIT 1
            """, (msg_id, raw_msg_id, am_id))
            prev_row = cur.fetchone()
            if prev_row and prev_row["detected_milestone"]:
                old_m_id = prev_row["detected_milestone"]
                if old_m_id != m_id:
                    print(f"🔄 AM {am_name} đã sửa tin nhắn từ Mốc {old_m_id} sang Mốc {m_id}. Tự động dọn bản ghi Mốc {old_m_id} cũ.")
                    cur.execute("""
                    DELETE FROM attendance_records
                    WHERE date = ? AND am_id = ? AND milestone_id = ?
                    """, (today_str, am_id, old_m_id))
                    try:
                        from sync_attendance_sheets import get_sheet_client, SPREADSHEET_ID, SHEET_TITLE
                        gc = get_sheet_client()
                        sh = gc.open_by_key(SPREADSHEET_ID)
                        ws = sh.worksheet(SHEET_TITLE)
                        cells = ws.findall(am_name)
                        date_display = submit_time.strftime("%d/%m/%Y")
                        for cell in cells:
                            row_vals = ws.row_values(cell.row)
                            if len(row_vals) > 0 and date_display in row_vals[0]:
                                col_idx = 4 + (old_m_id - 1)
                                ws.update_cell(cell.row, col_idx, "❌ Chưa nộp (100k)")
                                break
                    except Exception as e_sheet:
                        print(f"⚠️ Sheet reset cell error: {e_sheet}")

            # Kiểm tra xem AM này đã có bản ghi nào trước đó chưa
            cur.execute("""
            SELECT id, status, penalty_amount FROM attendance_records
            WHERE date = ? AND am_id = ? AND milestone_id = ?
            """, (today_str, am_id, m_id))
            existing = cur.fetchone()

            if existing:
                # Nếu trước đó đã nộp Đúng hạn thì không bị ghi đè thành Trễ
                if existing["status"] == "ON_TIME" and status != "ON_TIME":
                    status = "ON_TIME"
                    penalty = 0
                    note = "Cập nhật bổ sung (Đã ghi nhận đúng hạn trước đó)"
                else:
                    cur.execute("""
                    UPDATE attendance_records
                    SET submitted_at = ?, status = ?, late_minutes = ?, penalty_amount = ?,
                        raw_message_id = ?, updated_at = ?
                    WHERE id = ?
                    """, (
                        submit_time.strftime("%Y-%m-%d %H:%M:%S"), status, late_min, penalty,
                        raw_msg_id, get_vn_now().strftime("%Y-%m-%d %H:%M:%S"), existing["id"]
                    ))
            else:
                cur.execute("""
                INSERT INTO attendance_records (
                    date, am_id, am_name, milestone_id, submitted_at, status,
                    late_minutes, penalty_amount, raw_message_id, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    today_str, am_id, am_name, m_id,
                    submit_time.strftime("%Y-%m-%d %H:%M:%S"), status, late_min, penalty,
                    raw_msg_id, get_vn_now().strftime("%Y-%m-%d %H:%M:%S")
                ))

        conn.commit()

    # Kiểm tra tiến độ nộp các bưu cục điểm nóng (Mốc 5)
    m5_progress_note = ""
    if m_id == 5 and am_id:
        try:
            m5_req_map = get_m5_required_ams()
            req_hubs = m5_req_map.get(am_id, {}).get("hubs", [])
            if req_hubs:
                with get_db() as conn_m5:
                    cur_m = conn_m5.cursor()
                    cur_m.execute("""
                    SELECT raw_text FROM raw_messages
                    WHERE detected_am_id = ? AND detected_milestone = 5 AND DATE(received_at) = ?
                    """, (am_id, today_str))
                    all_m5_texts = [row["raw_text"] for row in cur_m.fetchall()]

                reported_hubs = set()
                combined_texts = all_m5_texts + [raw_text]
                for t in combined_texts:
                    found_h = detect_hubs_in_text(t)
                    for fh in found_h:
                        reported_hubs.add(fh["raw_hub"])

                covered = [h for h in req_hubs if h in reported_hubs]
                missing = [h for h in req_hubs if h not in reported_hubs]

                if len(req_hubs) > 1:
                    if not missing:
                        m5_progress_note = f"\n🎉 <b>Tiến độ điểm nóng:</b> Đã báo cáo đủ <b>{len(req_hubs)}/{len(req_hubs)} bưu cục</b>!"
                    else:
                        miss_str = ", ".join(missing)
                        m5_progress_note = f"\n📊 <b>Tiến độ điểm nóng:</b> Đã báo cáo <b>{len(covered)}/{len(req_hubs)} bưu cục</b>.\n⚠️ <i>Lưu ý: Còn thiếu bưu cục <b>{miss_str}</b>. Nhờ AM khẩn trương nộp bổ sung BC này trước 10:00!</i>"
                else:
                    if not missing:
                        m5_progress_note = f"\n🎉 <b>Tiến độ điểm nóng:</b> Đã hoàn thành báo cáo bưu cục ({covered[0]})!"
                    else:
                        m5_progress_note = f"\n⚠️ <i>Lưu ý: Báo cáo chưa đề cập đúng bưu cục điểm nóng ({req_hubs[0]}). Nhờ AM kiểm tra lại!</i>"
        except Exception as e_m5:
            print(f"⚠️ Lỗi check tiến độ bưu cục M5: {e_m5}")

    return {
        "am_name": am_name,
        "milestone_id": m_id,
        "milestone_name": m_name,
        "status": status,
        "late_minutes": late_min,
        "penalty": penalty,
        "note": note,
        "hubs": [h["raw_hub"] for h in matched_hubs] if matched_hubs else [],
        "m5_progress_note": m5_progress_note,
        "submit_time": submit_time.strftime("%H:%M:%S")
    }, "OK"

# ─── GỬI TIN GTALK ───────────────────────────────────────────
def send_gtalk_message(text: str, channel_id: str = None):
    config = load_config()
    token = config["gtalk"]["token"]
    ch_id = channel_id or config["gtalk"]["channel_id"]
    api_url = config["gtalk"]["api_url"]

    payload = {
        "channelId": str(ch_id),
        "clientMsgId": str(int(time.time() * 1000)),
        "content": {"parseMode": "HTML", "text": text},
        "oaToken": token,
    }
    try:
        r = requests.post(api_url, json=payload, timeout=15, verify=False)
        res = r.json() if r.status_code == 200 else {}
        if r.status_code == 200 and res.get("errorCode") == "success":
            return True, "OK"
        return False, f"HTTP {r.status_code} - {r.text[:200]}"
    except Exception as e:
        return False, str(e)

def get_m5_required_ams():
    """
    Đọc tab 'BC GTC dưới 50' từ Google Sheet để xác định chính xác những AM nào
    có bưu cục GTC < 50% bắt buộc phải nộp báo cáo Mốc 5 (BC Điểm nóng).
    Trả về: dict {am_id: {'am': am_dict, 'hubs': [list_of_hubs]}}
    """
    try:
        from sync_attendance_sheets import get_sheet_client, load_config
        gc = get_sheet_client()
        sh = gc.open_by_key('147nvGXc2D7UJNJGsWaFjaIZ6FkJDD7Zmevn3lBs-Bl0')
        ws = sh.worksheet('BC GTC dưới 50')
        rows = ws.get_all_values()
        if len(rows) <= 1:
            return {}

        config = load_config()
        parser = AttendanceParser(config)
        am_hubs_map = {}
        for r in rows[1:]:
            if len(r) >= 2 and r[0] and r[1]:
                # TỰ ĐỘNG BỎ QUA BƯU CỤC ĐANG BÀN GIAO / MIỄN BÁO CÁO (PHƯƠNG ÁN B - CÁCH 1)
                full_row_text = " ".join(r).lower()
                no_acc_row = remove_accents(full_row_text)
                if any(kw in full_row_text or kw in no_acc_row for kw in [
                    "ban giao", "bàn giao", "chuyen giao", "chuyển giao", "mien", "miễn", "loai tru", "loại trừ"
                ]):
                    continue

                hub = r[0].strip()
                am_str = r[1].strip()
                matched_am = parser.detect_am(am_str, am_str)
                if matched_am:
                    aid = matched_am['id']
                    if aid not in am_hubs_map:
                        am_hubs_map[aid] = {'am': matched_am, 'hubs': []}
                    am_hubs_map[aid]['hubs'].append(hub)
        return am_hubs_map
    except Exception as e:
        print(f"⚠️ Lỗi đọc tab BC GTC dưới 50: {e}")
        return {}


# ─── XÂY DỰNG BẢNG ĐIỂM DANH & TÍNH PHẠT ───────────────────────
def generate_milestone_recap(milestone_id: int, target_date: date = None):
    if target_date is None:
        target_date = get_vn_today()

    # Tự động khôi phục DB từ Google Sheet nếu có (đảm bảo không mất dữ liệu khi restart/redeploy)
    try:
        from sync_attendance_sheets import restore_db_from_sheet
        restore_db_from_sheet(target_date)
    except Exception as e:
        print(f"⚠️ restore_db_from_sheet error: {e}")

    config = load_config()
    ms_cfg = config["milestones"].get(str(milestone_id))
    if not ms_cfg:
        return ""

    ms_name = ms_cfg["name"]
    cutoff = ms_cfg["cutoff"]
    date_str = target_date.strftime("%Y-%m-%d")
    date_display = target_date.strftime("%d/%m/%Y")

    active_ams = [am for am in config["ams"] if am.get("is_active", True)]
    m5_map = {}
    if milestone_id == 5:
        m5_map = get_m5_required_ams()
        if m5_map:
            active_ams = [info['am'] for info in m5_map.values()]
        else:
            active_ams = []
    total_ams = len(active_ams)

    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT am_id, status, late_minutes, penalty_amount, strftime('%H:%M', submitted_at) as submit_time
        FROM attendance_records
        WHERE date = ? AND milestone_id = ?
        """, (date_str, milestone_id))
        records = {row["am_id"]: dict(row) for row in cur.fetchall()}

        cur.execute("""
        SELECT am_id, reason, excuse_type
        FROM excuses
        WHERE date = ? AND milestone_id = ?
        """, (date_str, milestone_id))
        excuse_dict = {row["am_id"]: {"reason": row["reason"], "type": dict(row).get("excuse_type", "LATE_PERMIT")} for row in cur.fetchall()}

    on_time_list = []
    late_list = []
    excused_list = []
    late_permit_list = []
    missing_list = []
    invalid_list = []

    for am in active_ams:
        am_id = am["id"]
        am_label = am["display_name"]
        rec = records.get(am_id)

        hub_info_str = ""
        if milestone_id == 5 and rec:
            req_hubs = m5_map.get(am_id, {}).get("hubs", [])
            if len(req_hubs) > 1:
                with get_db() as conn_sub:
                    cur_s = conn_sub.cursor()
                    cur_s.execute("""
                    SELECT raw_text FROM raw_messages
                    WHERE detected_am_id = ? AND detected_milestone = 5 AND DATE(received_at) = ?
                    """, (am_id, date_str))
                    all_m5_texts = [row["raw_text"] for row in cur_s.fetchall()]
                reported_hubs = set()
                for t in all_m5_texts:
                    for fh in detect_hubs_in_text(t):
                        reported_hubs.add(fh["raw_hub"])
                covered = [h for h in req_hubs if h in reported_hubs]
                missing = [h for h in req_hubs if h not in reported_hubs]
                if missing:
                    clean_miss = [re.sub(r'^\([A-Za-z0-9]+\)\s*', '', h).strip() for h in missing]
                    hub_info_str = f" - ⚠️ thiếu BC {', '.join(clean_miss)}"
                else:
                    hub_info_str = f" - đủ {len(req_hubs)}/{len(req_hubs)} BC"

        if not rec:
            if am_id in excuse_dict:
                if excuse_dict[am_id]["type"] == "EXEMPTION":
                    excused_list.append(f"{am_label} ({excuse_dict[am_id]['reason']})")
                else:
                    late_permit_list.append(f"{am_label} (Xin trễ: {excuse_dict[am_id]['reason']})")
            else:
                missing_list.append(am_label)
        elif rec["status"] == "ON_TIME":
            on_time_list.append(f"{am_label} ({rec['submit_time']}{hub_info_str})")
        elif rec["status"] == "LATE":
            late_list.append(f"{am_label} ({rec['submit_time']} - trễ {rec['late_minutes']}p{hub_info_str})")
        elif rec["status"] == "INVALID":
            invalid_list.append(f"{am_label} ({rec['submit_time']})")

    lines = [
        f"⏰ <b>ĐIỂM DANH MỐC {milestone_id}: {ms_name.upper()}</b>",
        f"📅 <i>Ngày: {date_display} | Mốc cut-off: {cutoff}</i>",
        "──────────────────────────────",
        f"📊 <b>Tiến độ:</b> {len(on_time_list)}/{total_ams} AM nộp đúng hạn",
        ""
    ]

    if on_time_list:
        lines.append("✅ <b>Đã nộp đúng hạn:</b>")
        lines.append(", ".join(on_time_list))
        lines.append("")

    if late_list:
        lines.append("⚠️ <b>Báo cáo trễ (Phạt 50k):</b>")
        for idx, item in enumerate(late_list, 1):
            lines.append(f"  {idx}. {item}")
        lines.append("")

    if excused_list:
        lines.append("📝 <b>Miễn báo cáo (0đ):</b>")
        for idx, item in enumerate(excused_list, 1):
            lines.append(f"  {idx}. {item}")
        lines.append("")

    if late_permit_list:
        lines.append("⏳ <b>Đã xin phép nộp trễ (Cần nộp bù):</b>")
        for idx, item in enumerate(late_permit_list, 1):
            lines.append(f"  {idx}. {item}")
        lines.append("")

    if missing_list:
        lines.append(f"❌ <b>Chưa nộp / Chưa xin phép tính tới {cutoff} ({len(missing_list)} AM):</b>")
        for idx, item in enumerate(missing_list, 1):
            lines.append(f"  {idx}. {item}")
        lines.append("")
        lines.append(f"👉 <i>Lưu ý: Báo cáo nộp sau {cutoff} sẽ tính Báo cáo trễ (Phạt 50.000đ). Nhờ các AM khẩn trương nộp bù!</i>")
    else:
        lines.append("🎉 <b>100% AM đã hoàn thành báo cáo mốc này!</b>")

    return "\n".join(lines)


def generate_milestone_reminder(milestone_id: int, target_date: date = None):
    """
    Tạo tin nhắn nhắc nhở nộp bù sau giờ cut-off (sau 30 phút).
    Chỉ gửi khi vẫn còn AM chưa nộp báo cáo. Nếu 100% đã nộp thì trả về None (không gửi).
    """
    if target_date is None:
        target_date = get_vn_today()

    try:
        from sync_attendance_sheets import restore_db_from_sheet
        restore_db_from_sheet(target_date)
    except Exception:
        pass

    config = load_config()
    ms_cfg = config["milestones"].get(str(milestone_id))
    if not ms_cfg:
        return None

    ms_name = ms_cfg["name"]
    cutoff = ms_cfg["cutoff"]
    date_str = target_date.strftime("%Y-%m-%d")

    active_ams = [am for am in config["ams"] if am.get("is_active", True)]
    if milestone_id == 5:
        m5_map = get_m5_required_ams()
        if m5_map:
            active_ams = [info['am'] for info in m5_map.values()]
        else:
            active_ams = []

    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT am_id, status
        FROM attendance_records
        WHERE date = ? AND milestone_id = ?
        """, (date_str, milestone_id))
        records = {row["am_id"]: dict(row) for row in cur.fetchall()}

    missing_list = []
    missing_hubs_list = []

    for am in active_ams:
        am_id = am["id"]
        rec = records.get(am_id)
        if not rec:
            missing_list.append(am["display_name"])
        elif milestone_id == 5:
            # Kiểm tra xem AM này có còn thiếu bưu cục nào chưa gửi không
            req_hubs = m5_map.get(am_id, {}).get("hubs", [])
            if req_hubs:
                with get_db() as conn_sub:
                    cur_s = conn_sub.cursor()
                    cur_s.execute("""
                    SELECT raw_text FROM raw_messages
                    WHERE detected_am_id = ? AND detected_milestone = 5 AND DATE(received_at) = ?
                    """, (am_id, date_str))
                    all_m5_texts = [row["raw_text"] for row in cur_s.fetchall()]
                reported_hubs = set()
                for t in all_m5_texts:
                    for fh in detect_hubs_in_text(t):
                        reported_hubs.add(fh["raw_hub"])
                missing = [h for h in req_hubs if h not in reported_hubs]
                if missing:
                    clean_miss = [re.sub(r'^\([A-Za-z0-9]+\)\s*', '', h).strip() for h in missing]
                    missing_hubs_list.append(f"<b>{am['display_name']}</b> (còn thiếu: <i>{', '.join(clean_miss)}</i>)")

    if not missing_list and not missing_hubs_list:
        return None  # 100% đã hoàn thành và đủ bưu cục, không cần nhắc nhở

    lines = [
        f"🔔 <b>NHẮC NHỞ NỘP BÙ: MỐC {milestone_id} - {ms_name.upper()}</b>",
        f"⏰ <i>Mốc cut-off: {cutoff}</i>",
        "──────────────────────────────"
    ]

    if missing_list:
        lines.append(f"❌ <b>AM chưa nộp ({len(missing_list)}/{len(active_ams)}):</b>")
        for idx, item in enumerate(missing_list, 1):
            lines.append(f"  {idx}. {item}")
        lines.append("")
        lines.append("⚠️ <i>Lưu ý: Báo cáo nộp bây giờ đã tính Báo cáo trễ (Phạt 50.000đ). Nếu không nộp trong ngày sẽ tính Không nộp (Phạt 100.000đ).</i>")

    if missing_hubs_list:
        lines.append("")
        lines.append("📌 <b>Nhắc gửi bổ sung bưu cục điểm nóng (Vẫn tính đúng hạn - 0đ):</b>")
        for item in missing_hubs_list:
            lines.append(f"  • {item}")
        lines.append("👉 <i>Nhờ các AM gửi bổ sung số liệu cho các bưu cục trên vào nhóm ạ!</i>")

    return "\n".join(lines)


def generate_daily_recap(target_date: date = None):
    if target_date is None:
        target_date = get_vn_today()

    # Tự động khôi phục DB từ Google Sheet nếu có
    try:
        from sync_attendance_sheets import restore_db_from_sheet
        restore_db_from_sheet(target_date)
    except Exception as e:
        print(f"⚠️ restore_db_from_sheet error: {e}")

    config = load_config()
    date_str = target_date.strftime("%Y-%m-%d")
    date_display = target_date.strftime("%d/%m/%Y")

    active_ams = [am for am in config["ams"] if am.get("is_active", True)]
    milestones = config["milestones"]

    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT am_id, milestone_id, status, late_minutes, penalty_amount
        FROM attendance_records
        WHERE date = ?
        """, (date_str,))
        records = {}
        for row in cur.fetchall():
            records[(row["am_id"], row["milestone_id"])] = dict(row)

        cur.execute("""
        SELECT am_id, am_name, milestone_id, reason, excuse_type
        FROM excuses
        WHERE date = ?
        """, (date_str,))
        excuses_map = {}
        for row in cur.fetchall():
            aid = row["am_id"]
            if aid not in excuses_map:
                excuses_map[aid] = {
                    "am_name": row["am_name"],
                    "milestones": set(),
                    "exemptions": set(),
                    "reasons": []
                }
            excuses_map[aid]["milestones"].add(row["milestone_id"])
            if dict(row).get("excuse_type") == "EXEMPTION":
                excuses_map[aid]["exemptions"].add(row["milestone_id"])
            if row["reason"] and row["reason"] not in excuses_map[aid]["reasons"]:
                excuses_map[aid]["reasons"].append(row["reason"])

    # Đọc danh sách AM được Admin tick Miễn Phạt / Nghỉ Phép trên Google Sheet
    excused_emp_ids = set()
    try:
        from sync_attendance_sheets import get_admin_excused_ams, sync_daily_to_sheet
        excused_emp_ids = get_admin_excused_ams(target_date)
    except Exception as e_sheet:
        print(f"Không thể đọc Google Sheet: {e_sheet}")

    total_region_fine = 0
    am_fines = []
    excused_ams = []
    m5_required = get_m5_required_ams()

    for am in active_ams:
        am_id = am["id"]
        emp_id = str(am.get("employee_id", ""))
        am_fine = 0
        violations = []

        if emp_id and emp_id in excused_emp_ids:
            # AM này được Admin tick Nghỉ phép / Miễn phạt
            excused_ams.append(am["display_name"])
            continue

        # Kiểm tra từng mốc từ 1 đến 4
        for m_id in (1, 2, 3, 4):
            rec = records.get((am_id, m_id))
            has_excuse = am_id in excuses_map and m_id in excuses_map[am_id]["milestones"]
            is_exempt = am_id in excuses_map and m_id in excuses_map[am_id]["exemptions"]

            if is_exempt:
                # AM xin miễn báo cáo mốc này -> Miễn phạt 0đ
                pass
            elif not rec or rec["status"] == "NOT_SUBMITTED":
                # Không nộp: Kể cả có xin trễ mà cuối cùng KHÔNG nộp -> Phạt 100k!
                am_fine += config["fines"]["not_submitted"]
                if has_excuse:
                    violations.append(f"Xin trễ nhưng không nộp M{m_id} (100k)")
                else:
                    violations.append(f"Không nộp M{m_id} (100k)")
            elif rec["status"] == "LATE":
                if has_excuse or rec.get("penalty_amount", 0) == 0:
                    pass  # Đã xin trễ và có nộp bù -> Miễn phạt 50k!
                else:
                    am_fine += config["fines"]["late"]
                    violations.append(f"Trễ M{m_id} (50k)")
            elif rec["status"] == "INVALID":
                am_fine += config["fines"]["invalid"]
                violations.append(f"Sai ĐK M{m_id} (200k)")

        # Mốc 5 (BC Điểm nóng: Chỉ tính phạt nếu AM có bưu cục < 50% trong tab 'BC GTC dưới 50')
        if am_id in m5_required:
            rec_m5 = records.get((am_id, 5))
            has_excuse_m5 = am_id in excuses_map and 5 in excuses_map[am_id]["milestones"]
            is_exempt_m5 = am_id in excuses_map and 5 in excuses_map[am_id]["exemptions"]

            if is_exempt_m5:
                # AM xin miễn báo cáo Mốc 5 -> Miễn phạt 0đ
                pass
            elif not rec_m5 or rec_m5["status"] == "NOT_SUBMITTED":
                am_fine += config["fines"]["not_submitted"]
                if has_excuse_m5:
                    violations.append("Xin trễ nhưng không nộp M5 (100k)")
                else:
                    violations.append("Không nộp M5 (100k)")
            elif rec_m5["status"] == "LATE":
                if has_excuse_m5 or rec_m5.get("penalty_amount", 0) == 0:
                    pass  # Đã xin trễ và có nộp bù -> Miễn phạt 50k!
                else:
                    am_fine += config["fines"]["late"]
                    violations.append("Trễ M5 (50k)")
            elif rec_m5["status"] == "INVALID":
                am_fine += config["fines"]["invalid"]
                violations.append("Sai ĐK M5 (200k)")
            elif rec_m5["status"] == "LATE":
                am_fine += config["fines"]["late"]
                violations.append("Trễ M5 (50k)")
            elif rec_m5["status"] == "INVALID":
                am_fine += config["fines"]["invalid"]
                violations.append("Sai ĐK M5 (200k)")

        total_region_fine += am_fine
        if am_fine > 0:
            am_fines.append({
                "am": am["display_name"],
                "fine": am_fine,
                "violations": ", ".join(violations)
            })

    # Sắp xếp người bị phạt nhiều nhất lên đầu
    am_fines.sort(key=lambda x: x["fine"], reverse=True)

    sheet_url = "https://docs.google.com/spreadsheets/d/147nvGXc2D7UJNJGsWaFjaIZ6FkJDD7Zmevn3lBs-Bl0/edit"

    lines = [
        f"📊 <b>CHỐT PHẠT BÁO CÁO NGÀY {date_display} - VÙNG NTB</b>",
        f"💰 <b>Tổng tiền phạt:</b> <b>{total_region_fine:,.0f} VNĐ</b>",
        ""
    ]

    violators = [item for item in am_fines if item["fine"] > 0]
    if violators:
        lines.append("🚫 <b>DANH SÁCH VI PHẠM:</b>")
        for idx, v in enumerate(violators, 1):
            lines.append(f"<b>{idx}. {v['am']}:</b> <code>{v['fine']:,.0f}đ</code> ({v['violations']})")
        lines.append("")
    else:
        lines.append("🎉 <i>Hôm nay cả vùng 100% đúng hạn, không có phát sinh phạt!</i>")
        lines.append("")

    # Danh sách xin phép trễ
    if excuses_map:
        lines.append("📝 <b>DANH SÁCH XIN PHÉP TRỄ:</b>")
        for aid, ex_info in excuses_map.items():
            ms_set = ex_info["milestones"]
            valid_ms = [m for m in ms_set if m is not None]
            if None in ms_set or ms_set == {1, 2, 3, 4, 5} or ms_set == {1, 2, 3, 4}:
                scope_str = "Cả ngày"
            elif ms_set == {1, 2, 5}:
                scope_str = "Ca sáng (M1, M2, M5)"
            elif ms_set == {3}:
                scope_str = "Ca chiều (M3: Gán TTS)"
            elif ms_set == {4}:
                scope_str = "Ca tối (M4: LTC TTS)"
            elif valid_ms:
                ms_sorted = sorted(valid_ms)
                scope_str = "Mốc " + ", ".join(str(m) for m in ms_sorted)
            else:
                scope_str = "Cả ngày"

            reason_str = " - ".join(ex_info["reasons"]) if ex_info["reasons"] else "Có báo trước"
            lines.append(f"• <b>{ex_info['am_name']}:</b> {scope_str} (Lý do: <i>{reason_str}</i>)")
        lines.append("")

    if excused_ams:
        lines.append(f"🏖️ <b>Nghỉ phép / Miễn phạt:</b> {', '.join(excused_ams)}")
        lines.append("")

    lines.append(f"🔗 <b>Bảng chi tiết & thu tiền:</b> <a href=\"{sheet_url}\">Mở Google Sheet</a>")
    lines.append("👉 <i>Mọi dữ liệu giờ nộp từng giây đã được lưu vết minh bạch trên Google Sheet.</i>")

    # Tự động đồng bộ số liệu mới nhất lên Google Sheet
    try:
        from sync_attendance_sheets import sync_daily_to_sheet
        threading.Thread(target=sync_daily_to_sheet, args=(target_date,), daemon=True).start()
    except Exception:
        pass

    return "\n".join(lines)

# ─── BỘ HẸN GIỜ (SCHEDULER TỰ ĐỘNG) ───────────────────────────
def start_scheduler():
    def scheduler_loop():
        last_triggered = {}
        print("⏰ Scheduler đã khởi động: Giám sát các mốc 08:00, 10:00, 11:00, 16:00, 20:00 & 20:15...")

        while True:
            now = get_vn_now()
            today_str = now.strftime("%Y-%m-%d")
            hm = now.strftime("%H:%M")

            config = load_config()
            cfg_gtalk = config.get("gtalk", {})
            group_a = cfg_gtalk.get("channel_id_group_a", "2097277790030348288")
            group_b = cfg_gtalk.get("channel_id_group_b", "2097270568973508608")

            # Danh sách mốc cần bắn recap và nhắc nộp bù theo từng Group
            schedule_map = {
                "08:00": (1, "Mốc 1 (Đầu ngày)", group_a, "recap"),
                "08:30": (1, "Mốc 1 (Nhắc nộp bù)", group_a, "reminder"),
                "10:00": (5, "Mốc 5 (Điểm nóng)", group_b, "recap"),
                "10:30": (5, "Mốc 5 (Nhắc nộp bù)", group_b, "reminder"),
                "11:00": (2, "Mốc 2 (Gán TTS Ca 1)", group_a, "recap"),
                "11:30": (2, "Mốc 2 (Nhắc nộp bù)", group_a, "reminder"),
                "16:00": (3, "Mốc 3 (Gán TTS Ca 2)", group_a, "recap"),
                "16:30": (3, "Mốc 3 (Nhắc nộp bù)", group_a, "reminder"),
                "20:00": (4, "Mốc 4 (LTC TTS)", group_a, "recap"),
            }

            if hm in schedule_map:
                m_id, label, target_group, action_type = schedule_map[hm]
                trigger_key = f"{today_str}_{m_id}_{action_type}"
                if trigger_key not in last_triggered:
                    last_triggered[trigger_key] = True
                    print(f"\n[{now.strftime('%H:%M:%S')}] 🔔 Kích hoạt {label} tại Group {target_group}!")
                    if action_type == "recap":
                        msg = generate_milestone_recap(m_id)
                    else:
                        msg = generate_milestone_reminder(m_id)

                    if msg:
                        ok, err = send_gtalk_message(msg, channel_id=target_group)
                        if ok:
                            print(f"✅ Đã gửi {label} lên GTalk Group {target_group} thành công!")
                        else:
                            print(f"❌ Lỗi bắn GTalk: {err}")
                    else:
                        print(f"ℹ️ Không có AM nào thiếu báo cáo cho {label}, bỏ qua gửi tin.")

            # Chốt sổ cả ngày lúc 20:15 (Ưu tiên gửi ảnh Dashboard cực nét, fallback text nếu lỗi)
            if hm == "20:15":
                daily_key = f"{today_str}_daily"
                if daily_key not in last_triggered:
                    last_triggered[daily_key] = True
                    print(f"\n[{now.strftime('%H:%M:%S')}] 🏆 Kích hoạt bảng tổng hợp phạt Daily N+1!")
                    sent_image = False
                    try:
                        from render_daily_recap_card import send_daily_attendance_recap_image
                        ok_img, err_img = send_daily_attendance_recap_image()
                        if ok_img:
                            sent_image = True
                            print("✅ Đã bắn ẢNH bảng chốt sổ Daily N+1 lên GTalk thành công!")
                        else:
                            print(f"⚠️ Không gửi được ảnh: {err_img}, chuyển sang gửi text.")
                    except Exception as e_img:
                        print(f"⚠️ Lỗi render ảnh ({e_img}), chuyển sang gửi text.")

                    if not sent_image:
                        msg = generate_daily_recap()
                        ok, err = send_gtalk_message(msg)
                        if ok:
                            print("✅ Đã bắn bảng chốt sổ Daily N+1 (Text) lên GTalk thành công!")
                        else:
                            print(f"❌ Lỗi bắn GTalk: {err}")

            time.sleep(20)

    t = threading.Thread(target=scheduler_loop, daemon=True)
    t.start()

# ─── FLASK WEBHOOK ROUTES ────────────────────────────────────
def extract_text(data: dict) -> str:
    content = data.get("content")
    if isinstance(content, dict) and content.get("text"):
        return content["text"]
    if isinstance(content, str) and content.strip():
        return content

    msg = data.get("message")
    if isinstance(msg, dict):
        mc = msg.get("content")
        if isinstance(mc, dict) and mc.get("text"):
            return mc["text"]
        if isinstance(mc, str):
            return mc
        if msg.get("text"):
            return msg["text"]

    if data.get("text"):
        return data["text"]
    return ""

@app.route("/", methods=["GET"])
def index():
    now_str = get_vn_now().strftime("%Y-%m-%d %H:%M:%S")
    return jsonify({
        "status": "online",
        "service": "GTalk Attendance & Cut-off Bot - Vùng NTB",
        "time": now_str,
        "database": DB_PATH
    })

@app.route("/webhook", methods=["POST"])
def webhook():
    ts = get_vn_now().strftime("%H:%M:%S")
    try:
        data = request.get_json(force=True, silent=True) or {}
    except Exception:
        data = {}

    msg_text = extract_text(data)
    if not msg_text:
        return jsonify({"status": "skipped", "reason": "empty message"})

    # Bỏ qua tin nhắn do bot tự gửi để tránh loop
    if "ĐIỂM DANH MỐC" in msg_text or "BẢNG TỔNG HỢP ĐIỂM DANH" in msg_text:
        return jsonify({"status": "skipped", "reason": "bot message"})

    msg_obj = data.get("message") if isinstance(data.get("message"), dict) else {}
    channel_id = str(data.get("channelId") or data.get("channel_id") or msg_obj.get("channelId") or "")
    sender_obj = data.get("sender") or msg_obj.get("sender") or {}
    sender_name = sender_obj.get("displayName") or sender_obj.get("name") or data.get("senderName") or ""
    sender_id = sender_obj.get("id") or str(data.get("senderId") or "")
    msg_id = str(data.get("id") or msg_obj.get("id") or time.time())

    print(f"\n[{ts}] 📩 Nhận tin nhắn mới: {msg_text[:80]}...")

    res, status_text = record_submission(
        sender_name=sender_name,
        sender_id=sender_id,
        raw_text=msg_text,
        channel_id=channel_id,
        msg_id=msg_id
    )

    if res:
        print(f"[{ts}] ✅ Đã ghi nhận: {res['am_name']} - Mốc {res['milestone_id']} ({res['status']} - Phạt {res['penalty']:,}đ)")
        return jsonify({"status": "recorded", "data": res})
    else:
        print(f"[{ts}] ℹ️ Bỏ qua: {status_text}")
        return jsonify({"status": "ignored", "reason": status_text})

# ─── MAIN & CLI ──────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Bot Điểm danh & Tính phạt Cut-off AM")
    parser.add_argument("--recap", type=int, help="Chạy điểm danh mốc cut-off (1-5) và gửi GTalk ngay")
    parser.add_argument("--daily", action="store_true", help="Chạy bảng tổng hợp Daily N+1 và gửi GTalk ngay")
    parser.add_argument("--status", action="store_true", help="Xem bảng điểm danh hiện tại trong ngày")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ in ra màn hình, không bắn GTalk thật")
    parser.add_argument("--port", type=int, default=5055, help="Port Webhook Server (mặc định 5055)")
    parser.add_argument("--no-ngrok", action="store_true", help="Không khởi động ngrok tunnel (dùng khi deploy Cloud)")
    args = parser.parse_args()

    init_db()

    if args.recap:
        m_id = args.recap
        print(f"Generating Milestone {m_id} Recap...")
        msg = generate_milestone_recap(m_id)
        print("\n" + "=" * 60)
        print(msg)
        print("=" * 60)
        if not args.dry_run:
            ok, err = send_gtalk_message(msg)
            print(f"Kết quả gửi GTalk: {ok} ({err})")
        return

    if args.daily:
        print("Generating Daily N+1 Recap...")
        msg = generate_daily_recap()
        print("\n" + "=" * 60)
        print(msg)
        print("=" * 60)
        if not args.dry_run:
            ok, err = send_gtalk_message(msg)
            print(f"Kết quả gửi GTalk: {ok} ({err})")
        return

    if args.status:
        config = load_config()
        today_str = date.today().strftime("%Y-%m-%d")
        print(f"\n📊 BẢNG TÌNH HÌNH ĐIỂM DANH HÔM NAY ({today_str}):")
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
            SELECT am_name, milestone_id, status, late_minutes, penalty_amount, submitted_at
            FROM attendance_records WHERE date = ?
            ORDER BY am_name, milestone_id
            """, (today_str,))
            rows = cur.fetchall()
            if not rows:
                print("Chưa có bản ghi điểm danh nào trong ngày hôm nay.")
            else:
                for r in rows:
                    print(f"- {r['am_name']} | Mốc {r['milestone_id']} | {r['status']} ({r['late_minutes']}p trễ) | Phạt: {r['penalty_amount']:,}đ | Lúc: {r['submitted_at']}")
        return

    print("=" * 65)
    print("🚀 BOT ĐIỂM DANH & TÍNH PHẠT CUT-OFF AM (VÙNG NTB)")
    print(f"📍 Database: {DB_PATH}")
    print(f"📍 Webhook Port: {args.port}")
    print("=" * 65)

    # Khởi động ngrok tunnel nếu không có cờ --no-ngrok
    if not args.no_ngrok:
        ngrok_token_file = os.path.join(BASE_DIR, "ngrok_token.txt")
        static_domain = "spring-provoke-valley.ngrok-free.dev"
        try:
            from pyngrok import ngrok, conf
            if os.path.exists(ngrok_token_file):
                with open(ngrok_token_file, "r") as f:
                    tok = f.read().strip()
                if tok:
                    conf.get_default().auth_token = tok
            try:
                ngrok.kill()
                time.sleep(1)
            except Exception:
                pass
            try:
                tunnel = ngrok.connect(args.port, "http", domain=static_domain)
                print(f"✅ Đã kết nối ngrok Domain: {static_domain}")
                print(f"👉 Webhook URL: {tunnel.public_url}/webhook")
            except Exception as e:
                print(f"⚠️ Kết nối static domain {static_domain} không thành công: {e}")
                tunnel = ngrok.connect(args.port, "http")
                print(f"👉 Webhook URL: {tunnel.public_url}/webhook")
        except Exception as ex:
            print(f"ℹ️ Không bật ngrok: {ex}")

    start_scheduler()
    app.run(host="0.0.0.0", port=args.port, debug=False, use_reloader=False)

if __name__ == "__main__":
    main()
