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
from datetime import datetime, date, timedelta

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
            UNIQUE(date, am_id, milestone_id)
        )
        """)
        conn.commit()

# ─── BỘ BÓC TÁCH (PARSER) ─────────────────────────────────────
def normalize_text(text: str) -> str:
    if not text:
        return ""
    s = unicodedata.normalize("NFC", str(text)).strip().lower()
    return re.sub(r'[\r\t]', ' ', s)

def remove_accents(input_str: str) -> str:
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

class AttendanceParser:
    def __init__(self, config):
        self.config = config
        self.ams = [am for am in config.get("ams", []) if am.get("is_active", True)]
        self.milestones = config.get("milestones", {})

    def detect_am(self, text: str, sender_name: str = ""):
        combined = f"{sender_name}\n{text}"
        norm_combined = normalize_text(combined)
        no_accent_combined = remove_accents(norm_combined)

        best_match = None
        longest_alias_len = 0

        for am in self.ams:
            for alias in am.get("aliases", []):
                norm_alias = normalize_text(alias)
                no_accent_alias = remove_accents(norm_alias)

                pattern = r'(?:\b|_)' + re.escape(norm_alias) + r'(?:\b|_)'
                pattern_no_accent = r'(?:\b|_)' + re.escape(no_accent_alias) + r'(?:\b|_)'

                if re.search(pattern, norm_combined) or re.search(pattern_no_accent, no_accent_combined):
                    if len(norm_alias) > longest_alias_len:
                        longest_alias_len = len(norm_alias)
                        best_match = am

        return best_match

    def detect_report_type(self, text: str, channel_id: str = ""):
        norm_txt = normalize_text(text)
        no_accent_txt = remove_accents(norm_txt)

        # Nếu gửi vào Group B (Group Báo cáo Điểm nóng) -> Mặc định là Mốc 5
        group_b_id = str(self.config.get("channel_id_group_b", "2095921878551764992"))
        if channel_id and str(channel_id) == group_b_id:
            return 5, "BC Điểm nóng (GTC <50%)"

        m1_keywords = ["tổng hợp đầu ngày", "dau ngay", "gtc ngày n-1", "tỷ lệ gtc", "nvpttt"]
        m4_keywords = ["ltc tts", "đơn ltc", "luân chuyển tts", "luan chuyen tts", "lc trước 23h"]
        m2_keywords = ["gán giaotts", "gan giaotts", "trước 9h", "truoc 9h", "trước 11h", "truoc 11h", "gán tts ca 1"]
        m3_keywords = ["trước: 16h", "trước 16h", "truoc 16h", "gán tts ca 2", "ca 2"]
        m5_keywords = [
            "điểm nóng", "diem nong", "gtc <50%", "gtc < 50%", "gtc dưới 50",
            "xuất hàng xong", "xuat hang xong", "thời gian xuất hàng", "time xuất hàng",
            "tồn / tổng", "tồn/tổng", "ton / tong", "ton/tong",
            "nhân viên đi làm", "nhan vien di lam", "nv đi làm", "nv di lam",
            "bưu cục :", "bưu cục:"
        ]

        if any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m1_keywords):
            return 1, "Tổng hợp đầu ngày"
        if any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m4_keywords):
            return 4, "LTC TTS"
        if any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m3_keywords):
            return 3, "Gán TTS ca 2"
        if any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m2_keywords):
            return 2, "Gán TTS ca 1"
        if any(k in norm_txt or remove_accents(k) in no_accent_txt for k in m5_keywords):
            return 5, "BC Điểm nóng (GTC <50%)"

        if "gán tts" in norm_txt or "gan tts" in no_accent_txt:
            if "16h" in norm_txt:
                return 3, "Gán TTS ca 2"
            return 2, "Gán TTS ca 1"

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
                raw_hub = r[0].strip()
                am_str = r[1].strip()
                clean_hub = re.sub(r'^\([A-Za-z0-9]+\)\s*', '', raw_hub).strip()
                norm_raw = normalize_text(raw_hub)
                norm_clean = normalize_text(clean_hub)
                no_acc_clean = remove_accents(norm_clean)

                if norm_raw in norm_msg or norm_clean in norm_msg or no_acc_clean in no_accent_msg:
                    matched_am = parser.detect_am(am_str, am_str)
                    matched.append({
                        'raw_hub': raw_hub,
                        'clean_hub': clean_hub,
                        'am_name_sheet': am_str,
                        'am': matched_am
                    })
        return matched
    except Exception as e:
        print(f"⚠️ Lỗi detect_hubs_in_text: {e}")
        return []

# ─── TÍNH PHẠT & GHI NHẬN ─────────────────────────────────────
def detect_excuse_request(raw_text: str, sender_name: str = "", dt: datetime = None):
    """
    Tự động phát hiện khi AM nhắn tin xin phép báo cáo trễ / xin off phép.
    """
    if dt is None:
        dt = datetime.now()

    norm_txt = normalize_text(raw_text)
    no_accent_txt = remove_accents(norm_txt)

    excuse_patterns = [
        r'xin\s+(?:phép\s+)?(?:báo\s+cáo\s+)?(?:nộp\s+)?trễ',
        r'xin\s+trễ',
        r'báo\s+cáo\s+trễ',
        r'xin\s+nộp\s+trễ',
        r'xin\s+(?:phép\s+)?off',
        r'xin\s+(?:phép\s+)?nghỉ',
        r'nghỉ\s+phép',
        r'off\s+phép'
    ]

    is_excuse = any(re.search(p, norm_txt) or re.search(remove_accents(p), no_accent_txt) for p in excuse_patterns)
    if not is_excuse:
        return None

    config = load_config()
    parser = AttendanceParser(config)
    detected_am = parser.detect_am(raw_text, sender_name)
    if not detected_am:
        detected_am = parser.detect_am(sender_name, "")
    if not detected_am:
        return None

    # Xác định phạm vi mốc
    if any(k in norm_txt for k in ["cả ngày", "ca ngay", "nguyên ngày", "hôm nay off", "hom nay off"]):
        milestones = [1, 5, 2, 3, 4]
        scope_label = "Cả ngày (Tất cả các mốc)"
        next_reminder = "Chúc AM nghỉ ngơi / công tác tốt nhé!"
    elif any(k in norm_txt for k in ["chiều", "chieu", "ca 2", "16h", "mốc 3", "moc 3"]):
        milestones = [3]
        scope_label = "Ca chiều (Mốc 3 - Gán TTS Ca 2 16h00)"
        next_reminder = "Mốc tối (20h00 - LTC TTS) vẫn báo cáo đúng timeline quy định nhé!"
    elif any(k in norm_txt for k in ["tối", "toi", "ltc", "20h", "mốc 4", "moc 4"]):
        milestones = [4]
        scope_label = "Ca tối (Mốc 4 - LTC TTS 20h00)"
        next_reminder = "Nhờ AM nộp bù trước khi kết thúc ca làm việc nhé!"
    elif any(k in norm_txt for k in ["sáng", "sang", "ca 1", "đầu ngày", "dau ngay", "11h", "mốc 1", "mốc 2", "mốc 5"]):
        milestones = [1, 5, 2]
        scope_label = "Ca sáng (Mốc 1 - Đầu ngày, Mốc 5 - Điểm nóng, Mốc 2 - Gán TTS)"
        next_reminder = "Ca chiều (16h00) và Ca tối (20h00) vẫn báo cáo đúng timeline quy định nhé!"
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

    reason = "Bận việc đột xuất / Đi tuyến"
    for cue in ["do", "vì", "vi", "hẹn", "hen", "gặp", "gap", "bận", "ban", "đi", "di"]:
        if f" {cue} " in f" {norm_txt} ":
            idx = norm_txt.find(cue)
            reason_part = raw_text[idx:].strip()
            reason = reason_part[:50].split("\n")[0].strip()
            break

    today_str = dt.strftime("%Y-%m-%d")
    with get_db() as conn:
        cur = conn.cursor()
        for m_id in milestones:
            cur.execute("""
            INSERT OR REPLACE INTO excuses (date, am_id, am_name, milestone_id, reason, raw_text, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                today_str, detected_am["id"], detected_am["full_name"], m_id,
                reason, raw_text, dt.strftime("%Y-%m-%d %H:%M:%S")
            ))
        conn.commit()

    reply_msg = (
        f"📝 <b>XÁC NHẬN GHI NHẬN XIN PHÉP BÁO CÁO TRỄ</b>\n"
        f"👤 <b>AM:</b> {detected_am['full_name']}\n"
        f"🕒 <b>Phạm vi áp dụng:</b> {scope_label}\n"
        f"📌 <b>Lý do ghi nhận:</b> {reason}\n"
        f"👉 <i>Lưu ý: Nhờ AM nộp bù trong ngày để được tính Hợp lệ (0đ). {next_reminder}</i>"
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
    cutoff_dt = dt.replace(hour=cutoff_hour, minute=cutoff_min, second=0, microsecond=0)

    if dt <= cutoff_dt:
        return "ON_TIME", 0, 0, "Đúng hạn"
    else:
        diff_sec = (dt - cutoff_dt).total_seconds()
        late_min = max(1, int(diff_sec // 60))

        # Kiểm tra xem AM có xin phép trễ mốc này trong ngày chưa
        if am_id:
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                    SELECT reason FROM excuses WHERE date = ? AND am_id = ? AND milestone_id = ?
                    """, (dt.strftime("%Y-%m-%d"), am_id, milestone_id))
                    row = cur.fetchone()
                    if row:
                        return "ON_TIME", late_min, 0, f"Đã xin phép ({row['reason']}) — Miễn phạt 50k"
            except Exception:
                pass

        return "LATE", late_min, config["fines"]["late"], f"Trễ {late_min} phút"

def record_submission(sender_name, sender_id, raw_text, channel_id, msg_id, submit_time=None):
    if submit_time is None:
        submit_time = datetime.now()

    config = load_config()
    parser = AttendanceParser(config)

    detected_am = parser.detect_am(raw_text, sender_name)
    m_id, m_name = parser.detect_report_type(raw_text, channel_id)

    # Nếu là Mốc 5 hoặc gửi vào Group B: Quét tìm bưu cục trong tin nhắn
    matched_hubs = []
    group_b_id = str(config.get("channel_id_group_b", "2095921878551764992"))
    if m_id == 5 or (channel_id and str(channel_id) == group_b_id):
        m_id = 5
        m_name = "BC Điểm nóng (GTC <50%)"
        matched_hubs = detect_hubs_in_text(raw_text)
        if matched_hubs and not detected_am:
            detected_am = matched_hubs[0]["am"]

    if not m_id:
        return None, "Không phải mẫu báo cáo 1-5"

    is_valid, err_reason = parser.validate_content(m_id, raw_text)

    am_id = detected_am["id"] if detected_am else None
    am_name = detected_am["full_name"] if detected_am else (sender_name or "Chưa rõ AM")

    status, late_min, penalty, note = evaluate_submission(submit_time, m_id, is_valid, config, am_id)

    today_str = submit_time.strftime("%Y-%m-%d")

    with get_db() as conn:
        cur = conn.cursor()
        # Lưu raw_messages
        cur.execute("""
        INSERT INTO raw_messages (
            message_id, channel_id, sender_id, sender_name, raw_text,
            received_at, detected_am_id, detected_am_name, detected_milestone,
            is_valid, error_reason, evaluation_status, late_minutes, penalty_amount
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            msg_id, channel_id, sender_id, sender_name, raw_text,
            submit_time.strftime("%Y-%m-%d %H:%M:%S"), am_id, am_name, m_id,
            1 if is_valid else 0, err_reason if not is_valid else "",
            status, late_min, penalty
        ))
        raw_msg_id = cur.lastrowid

        # Cập nhật attendance_records (nếu đã nhận đúng AM)
        if am_id:
            # Kiểm tra xem AM này đã có bản ghi nào trước đó chưa
            cur.execute("""
            SELECT id, status, penalty_amount FROM attendance_records
            WHERE date = ? AND am_id = ? AND milestone_id = ?
            """, (today_str, am_id, m_id))
            existing = cur.fetchone()

            if existing:
                # Nếu trước đó đã nộp Đúng hạn thì không bị ghi đè thành Trễ
                if existing["status"] == "ON_TIME" and status != "ON_TIME":
                    pass
                else:
                    cur.execute("""
                    UPDATE attendance_records
                    SET submitted_at = ?, status = ?, late_minutes = ?, penalty_amount = ?,
                        raw_message_id = ?, updated_at = ?
                    WHERE id = ?
                    """, (
                        submit_time.strftime("%Y-%m-%d %H:%M:%S"), status, late_min, penalty,
                        raw_msg_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), existing["id"]
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
                    raw_msg_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ))

        conn.commit()

    return {
        "am_name": am_name,
        "milestone_id": m_id,
        "milestone_name": m_name,
        "status": status,
        "late_minutes": late_min,
        "penalty": penalty,
        "note": note,
        "hubs": [h["raw_hub"] for h in matched_hubs] if matched_hubs else [],
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
        target_date = date.today()

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

    on_time_list = []
    late_list = []
    missing_list = []
    invalid_list = []

    for am in active_ams:
        am_id = am["id"]
        am_label = am["display_name"]
        rec = records.get(am_id)

        if not rec:
            missing_list.append(am_label)
        elif rec["status"] == "ON_TIME":
            on_time_list.append(f"{am_label} ({rec['submit_time']})")
        elif rec["status"] == "LATE":
            late_list.append(f"{am_label} ({rec['submit_time']} - trễ {rec['late_minutes']}p)")
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

    if invalid_list:
        lines.append("🚫 <b>Báo cáo sai/thiếu tiêu chí (Phạt 200k):</b>")
        for idx, item in enumerate(invalid_list, 1):
            lines.append(f"  {idx}. {item}")
        lines.append("")

    if missing_list:
        lines.append(f"⏳ <b>Chưa nộp tính tới {cutoff} ({len(missing_list)} AM):</b>")
        for idx, item in enumerate(missing_list, 1):
            lines.append(f"  {idx}. {item}")
        lines.append("")
        lines.append(f"👉 <i>Lưu ý: Báo cáo nộp sau {cutoff} sẽ tính Báo cáo trễ (Phạt 50.000đ). Nhờ các AM khẩn trương nộp bù!</i>")
    else:
        lines.append("🎉 <b>100% AM đã hoàn thành báo cáo mốc này!</b>")

    return "\n".join(lines)


def generate_daily_recap(target_date: date = None):
    if target_date is None:
        # Mặc định là ngày hôm nay nếu chạy tối, hoặc ngày N-1 nếu chạy sáng hôm sau
        target_date = date.today()

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
        SELECT am_id, am_name, milestone_id, reason
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
                    "reasons": []
                }
            excuses_map[aid]["milestones"].add(row["milestone_id"])
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
            if not rec:
                # Không nộp -> Phạt 100k
                am_fine += config["fines"]["not_submitted"]
                violations.append(f"Chưa nộp M{m_id} (100k)")
            elif rec["status"] == "LATE":
                if has_excuse or rec.get("penalty_amount", 0) == 0:
                    pass  # Đã xin phép -> Miễn phạt 50k
                else:
                    am_fine += config["fines"]["late"]
                    violations.append(f"Trễ M{m_id} (50k)")
            elif rec["status"] == "INVALID":
                am_fine += config["fines"]["invalid"]
                violations.append(f"Sai ĐK M{m_id} (200k)")

        # Mốc 5 (BC Điểm nóng: Chỉ tính phạt nếu AM có bưu cục < 50% trong tab 'BC GTC dưới 50')
        m5_required = get_m5_required_ams()
        if am_id in m5_required:
            rec_m5 = records.get((am_id, 5))
            has_excuse_m5 = am_id in excuses_map and 5 in excuses_map[am_id]["milestones"]
            if not rec_m5:
                am_fine += config["fines"]["not_submitted"]
                violations.append("Chưa nộp M5 (100k)")
            elif rec_m5["status"] == "LATE":
                if not (has_excuse_m5 or rec_m5.get("penalty_amount", 0) == 0):
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
            if ms_set == {1, 2, 3, 4, 5} or ms_set == {1, 2, 3, 4}:
                scope_str = "Cả ngày"
            elif ms_set == {1, 2, 5}:
                scope_str = "Ca sáng (M1, M2, M5)"
            elif ms_set == {3}:
                scope_str = "Ca chiều (M3: Gán TTS)"
            elif ms_set == {4}:
                scope_str = "Ca tối (M4: LTC TTS)"
            else:
                ms_sorted = sorted(list(ms_set))
                scope_str = "Mốc " + ", ".join(str(m) for m in ms_sorted)

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
            now = datetime.now()
            today_str = now.strftime("%Y-%m-%d")
            hm = now.strftime("%H:%M")

            config = load_config()
            cfg_gtalk = config.get("gtalk", {})
            group_a = cfg_gtalk.get("channel_id_group_a", "2077278419534073856")
            group_b = cfg_gtalk.get("channel_id_group_b", "2095921878551764992")

            # Danh sách mốc cần bắn recap theo từng Group
            schedule_map = {
                "08:00": (1, "Mốc 1 (Đầu ngày)", group_a),
                "10:00": (5, "Mốc 5 (Điểm nóng)", group_b),
                "11:00": (2, "Mốc 2 (Gán TTS Ca 1)", group_a),
                "16:00": (3, "Mốc 3 (Gán TTS Ca 2)", group_a),
                "20:00": (4, "Mốc 4 (LTC TTS)", group_a),
            }

            if hm in schedule_map:
                m_id, label, target_group = schedule_map[hm]
                trigger_key = f"{today_str}_{m_id}"
                if trigger_key not in last_triggered:
                    last_triggered[trigger_key] = True
                    print(f"\n[{now.strftime('%H:%M:%S')}] 🔔 Kích hoạt điểm danh cut-off {label} tại Group {target_group}!")
                    msg = generate_milestone_recap(m_id)
                    ok, err = send_gtalk_message(msg, channel_id=target_group)
                    if ok:
                        print(f"✅ Đã bắn điểm danh mốc {m_id} lên GTalk Group {target_group} thành công!")
                    else:
                        print(f"❌ Lỗi bắn GTalk: {err}")

            # Chốt sổ cả ngày lúc 20:15
            if hm == "20:15":
                daily_key = f"{today_str}_daily"
                if daily_key not in last_triggered:
                    last_triggered[daily_key] = True
                    print(f"\n[{now.strftime('%H:%M:%S')}] 🏆 Kích hoạt bảng tổng hợp phạt Daily N+1!")
                    msg = generate_daily_recap()
                    ok, err = send_gtalk_message(msg)
                    if ok:
                        print("✅ Đã bắn bảng chốt sổ Daily N+1 lên GTalk thành công!")
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
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return jsonify({
        "status": "online",
        "service": "GTalk Attendance & Cut-off Bot - Vùng NTB",
        "time": now_str,
        "database": DB_PATH
    })

@app.route("/webhook", methods=["POST"])
def webhook():
    ts = datetime.now().strftime("%H:%M:%S")
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
