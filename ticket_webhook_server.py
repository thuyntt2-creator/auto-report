# -*- coding: utf-8 -*-
"""
Script: ticket_webhook_server.py
Flask Webhook Server — Nhận tin nhắn gửi vào GTalk Group trigger (2077278419534073856),
tự động phân tích danh sách AM có ticket tồn và "bắn" cảnh báo riêng tới từng Group AM theo link Google Sheet.

Cách dùng:
  - Nhấp đúp file CHAY_TICKET_WEBHOOK.bat
  - Hoặc gõ lệnh: python ticket_webhook_server.py

Quy trình:
  1. Khi bạn dán / gửi bản tin tồn phiếu vào Group Trợ lý / Quản lý (ID: 2077278419534073856)
  2. Webhook bắt được tin nhắn này
  3. Đọc danh sách Group ID của từng AM từ cột C & D trong Google Sheet
  4. Tự động gửi cảnh báo riêng cho từng AM vào Group của AM đó!
"""

import os
import sys
import re
import time
import json
import unicodedata
import threading
from datetime import datetime, timezone, timedelta

VN_TZ = timezone(timedelta(hours=7))

def get_vn_now() -> datetime:
    """Trả về thời gian hiện tại chuẩn theo múi giờ Việt Nam (GMT+7)."""
    return datetime.now(VN_TZ).replace(tzinfo=None)

import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from flask import Flask, request, jsonify
import gspread
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials

# ─── Encoding UTF-8 ─────────────────────────────────────────
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

# ─── CẤU HÌNH ───────────────────────────────────────────────
OA_TOKEN        = "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
TRIGGER_GROUPS   = ["2097277790030348288", "2097270568973508608", "2077278419534073856", "2095921878551764992"]   # Group nhận tin nhắn → trigger bắn sang các AM
STATIC_DOMAIN    = "spring-provoke-valley.ngrok-free.dev"
SPREADSHEET_ID  = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
SHEET_TAB       = "ticket"
GTALK_API_URL   = "https://mbff.ghn.vn/api/gtalk/send-message"
WEBHOOK_PORT    = int(os.environ.get("PORT", 5055))
NGROK_TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ngrok_token.txt")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
AUTH_CANDIDATES = [
    os.path.join(BASE_DIR, "authorized_user.json"),
    r"c:\Users\lap4all\Desktop\auto-report\authorized_user.json",
    r"C:\Users\lap4all\Documents\Auto report\authorized_user.json",
    r"C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json",
    "authorized_user.json",
]
SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, "credentials.json"),
    r"c:\Users\lap4all\Desktop\auto-report\credentials.json",
    r"C:\Users\lap4all\Documents\Auto report\credentials.json",
    r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
    "credentials.json",
]
# ────────────────────────────────────────────────────────────

app = Flask(__name__)


# ─── GOOGLE SHEETS ──────────────────────────────────────────

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
                creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
                return gspread.authorize(creds)
            except Exception:
                pass
    raise PermissionError("Không thể xác thực Google Sheets!")


def load_am_mapping():
    """Đọc mapping AM → Group ID từ cột C/D của sheet ticket."""
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(SHEET_TAB)
    rows = ws.get_all_values()

    am_map = {}
    token = OA_TOKEN

    for row in rows:
        c_val = str(row[2]).strip() if len(row) > 2 else ""
        d_val = str(row[3]).strip() if len(row) > 3 else ""
        if not c_val or not d_val:
            continue
        c_low = c_val.lower()
        if c_low in ("am", "token", "token oa"):
            token = d_val
            continue
        if c_low in ("email_am", "tên am", "ten am", "name"):
            continue
        am_map[norm(c_val)] = {"display_name": c_val, "group_id": d_val}

    print(f"[SHEET] Loaded {len(am_map)} AMs | Token: {token[:15]}...")
    return am_map, token


def norm(s):
    if not s:
        return ""
    s_norm = unicodedata.normalize("NFC", str(s)).strip().lower()
    s_norm = re.sub(r'^[👤📦▸•\s\*\-]+', '', s_norm).strip()
    return s_norm


# ─── PARSE TIN NHẮN ─────────────────────────────────────────

def safe_get_dict(val):
    """Chuyển string JSON thành dict nếu cần."""
    if isinstance(val, dict):
        return val
    if isinstance(val, str):
        try:
            parsed = json.loads(val)
            if isinstance(parsed, dict):
                return parsed
        except Exception:
            pass
    return {}


def extract_text(data: dict) -> str:
    """Trích xuất text từ nhiều dạng payload GTalk webhook."""
    content = data.get("content")
    if content:
        content_dict = safe_get_dict(content)
        if content_dict.get("text"):
            return content_dict["text"]
        if isinstance(content, str) and content.strip():
            return content

    msg = data.get("message") or {}
    msg = safe_get_dict(msg)
    msg_content = msg.get("content")
    if msg_content:
        mc = safe_get_dict(msg_content)
        if mc.get("text"):
            return mc["text"]
        if isinstance(msg_content, str):
            return msg_content

    if data.get("text"):
        return data["text"]

    if msg.get("text"):
        return msg["text"]

    if data.get("body"):
        return str(data["body"])

    return ""


def parse_ticket_message(text: str):
    """
    Parse nội dung tin nhắn tìm:
      - header: dòng "Hi Anh/Chị..." hoặc "CẢNH BÁO TỔNG HỢP..."
      - am_items: list dict {am_name, raw_line} từ các dòng AM có phiếu
    """
    header = ""
    am_items = []

    normalized_text = re.sub(r'(?i)<br\s*/?>', '\n', text)

    for raw_line in normalized_text.split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        # Header
        if not header and any(kw in line for kw in [
            "Hi Anh/Chị", "tính tới thời điểm", "ƯU TIÊN XỬ LÝ", "CẢNH BÁO TỔNG HỢP"
        ]):
            header = re.sub(r'<[^>]+>', '', line).strip()
            continue

        # Dòng AM
        if "(AM):" in line or (("👤" in line or "📦" in line) and ":" in line):
            if any(kw in line for kw in ["Hành động", "hỗ trợ vui lòng", "Hi Anh/Chị"]):
                continue

            clean_line = re.sub(r'<[^>]+>', '', line).strip()
            m = re.match(r'^[^\w\u00C0-\u024F\u1E00-\u1EFF]*(.+?)\s*(?:\(AM\))?\s*:', clean_line)
            if m:
                am_name = m.group(1).strip()
                if am_name:
                    am_items.append({"am_name": am_name, "raw_line": clean_line})

    return header, am_items


# ─── GỬI GTALK ──────────────────────────────────────────────

def send_gtalk(group_id: str, text: str, oa_token: str):
    payload = {
        "channelId": str(group_id),
        "clientMsgId": str(int(get_vn_now().timestamp() * 1000)),
        "content": {"parseMode": "HTML", "text": text},
        "oaToken": oa_token,
    }
    try:
        r = requests.post(GTALK_API_URL, json=payload, timeout=15, verify=False)
        res = {}
        try:
            res = r.json()
        except Exception:
            pass
        if r.status_code == 200 and res.get("errorCode") == "success":
            return True, "OK"
        return False, f"HTTP {r.status_code} - {r.text[:200]}"
    except Exception as e:
        return False, str(e)


def build_message(header: str, raw_line: str) -> str:
    lines = [
        "🚨 <b>CẢNH BÁO TỒN PHIẾU HỐI GIAO/LẤY/TRẢ (GLT)</b> 🚨",
        "",
        f"<b>{raw_line}</b>"
    ]
    if header:
        lines.insert(1, f"⏰ <i>{header}</i>")

    lines.append("")
    lines.append("👉 <i>Hành động: Nhờ AM đôn đốc các bưu cục xử lý ngay nhé!</i>")
    lines.append(f"🔗 <b>Link kiểm tra & xử lý:</b> <a href=\"{DEFAULT_PORTAL_URL}\">{DEFAULT_PORTAL_URL}</a>")
    return "\n".join(lines)


# ─── FLASK ROUTES ────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def health():
    if request.method == "POST":
        return webhook()
    return jsonify({
        "status": "ok",
        "message": "Ticket Webhook Server đang chạy ✅",
        "trigger_groups": TRIGGER_GROUPS,
        "spreadsheet_id": SPREADSHEET_ID,
        "time": get_vn_now().strftime("%d/%m/%Y %H:%M:%S"),
    })

@app.route("/test_sheet", methods=["GET"])
def test_sheet():
    try:
        from sync_attendance_sheets import get_sheet_client
        gc = get_sheet_client()
        sh = gc.open_by_key('147nvGXc2D7UJNJGsWaFjaIZ6FkJDD7Zmevn3lBs-Bl0')
        ws = sh.worksheet('BC GTC dưới 50')
        rows = ws.get_all_values()
        return jsonify({"status": "ok", "rows_count": len(rows), "sample": rows[1] if len(rows) > 1 else []})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


recent_logs = []

@app.route("/debug/recent", methods=["GET"])
def debug_recent():
    try:
        from diem_danh_bot import get_db
        db_rows = []
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, message_id, channel_id, sender_name, received_at, detected_am_name, detected_milestone, is_valid, error_reason FROM raw_messages ORDER BY id DESC LIMIT 15")
            db_rows = [dict(r) for r in cur.fetchall()]
        return jsonify({
            "status": "ok",
            "recent_logs": recent_logs[-15:],
            "db_messages": db_rows
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e), "recent_logs": recent_logs[-15:]})

@app.route("/health", methods=["GET"])
def health_check():
    return jsonify({"status": "ok", "time": get_vn_now().strftime("%Y-%m-%d %H:%M:%S")})

@app.route("/admin/fix_duy", methods=["GET", "POST"])
def admin_fix_duy():
    try:
        from diem_danh_bot import get_db, get_vn_today
        from sync_attendance_sheets import sync_daily_to_sheet
        today_str = get_vn_today().strftime("%Y-%m-%d")
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
            INSERT OR REPLACE INTO excuses (date, am_id, am_name, milestone_id, reason, raw_text, created_at, excuse_type)
            VALUES (?, 'am_duy_pd', 'Phan Đình Duy', 5, 'Đi tuyến / Miễn báo cáo Mốc 5', 'Xin miễn mốc 5', CURRENT_TIMESTAMP, 'EXEMPTION')
            """, (today_str,))
            cur.execute("""
            INSERT INTO attendance_records (date, am_id, am_name, milestone_id, submitted_at, status, late_minutes, penalty_amount, updated_at)
            VALUES (?, 'am_duy_pd', 'Phan Đình Duy', 5, CURRENT_TIMESTAMP, 'EXEMPT', 0, 0, CURRENT_TIMESTAMP)
            ON CONFLICT(date, am_id, milestone_id) DO UPDATE SET
                status = 'EXEMPT',
                late_minutes = 0,
                penalty_amount = 0,
                updated_at = CURRENT_TIMESTAMP
            """, (today_str,))
            conn.commit()
        sync_daily_to_sheet(get_vn_today())
        return jsonify({"status": "ok", "message": "Updated AM Duy exemption in SQLite and synced to Google Sheet"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/cron/recap/<int:m_id>", methods=["GET", "POST"])
def trigger_cron_recap(m_id):
    try:
        from diem_danh_bot import generate_milestone_recap, send_gtalk_message, load_config
        cfg = load_config()
        channel = cfg.get("gtalk", {}).get("channel_id_group_b") if m_id == 5 else cfg.get("gtalk", {}).get("channel_id_group_a")
        msg = generate_milestone_recap(m_id)
        if msg:
            ok, err = send_gtalk_message(msg, channel_id=channel)
            return jsonify({"status": "ok", "sent": ok, "err": str(err)})
        return jsonify({"status": "skipped", "reason": "empty message"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/cron/daily", methods=["GET", "POST"])
def trigger_cron_daily():
    try:
        sent_image = False
        try:
            from render_daily_recap_card import send_daily_attendance_recap_image
            ok_img, err_img = send_daily_attendance_recap_image()
            if ok_img:
                sent_image = True
        except Exception as e_img:
            print(f"⚠️ Không render được ảnh cron daily: {e_img}")

        if not sent_image:
            from diem_danh_bot import generate_daily_recap, send_gtalk_message, load_config
            cfg = load_config()
            channel = cfg.get("gtalk", {}).get("channel_id_group_a")
            msg = generate_daily_recap()
            if msg:
                ok, err = send_gtalk_message(msg, channel_id=channel)
                return jsonify({"status": "ok", "sent_text": ok, "err": str(err)})
        return jsonify({"status": "ok", "sent_image": sent_image})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/webhook", methods=["POST"])
@app.route("/gtalk/webhook", methods=["POST"])
def webhook():
    ts = get_vn_now().strftime("%H:%M:%S")

    try:
        data = request.get_json(force=True, silent=True) or {}
    except Exception:
        data = {}

    print(f"\n[{ts}] === WEBHOOK RECEIVED ===")
    print(json.dumps(data, ensure_ascii=False, indent=2)[:500])

    # Lấy channel ID
    msg_obj = safe_get_dict(data.get("message"))
    channel_id = str(data.get("channelId") or data.get("channel_id") or msg_obj.get("channelId") or "")

    # Trích text
    msg_text = extract_text(data)

    recent_logs.append({
        "time": ts,
        "channel_id": channel_id,
        "text_preview": msg_text[:120] if msg_text else "(empty)",
        "keys": list(data.keys())
    })
    if len(recent_logs) > 50:
        recent_logs.pop(0)

    if not msg_text:
        print(f"[{ts}] SKIP — Không có nội dung")
        return jsonify({"status": "skipped", "reason": "empty message"})

    # Bỏ qua nếu là tin do Bot tự gửi (tránh loop khi gửi vào group test)
    if any(kw in msg_text for kw in [
        "CẢNH BÁO TỒN PHIẾU HỐI GIAO/LẤY/TRẢ (GLT)",
        "XÁC NHẬN ĐÃ GHI NHẬN BÁO CÁO",
        "XÁC NHẬN GHI NHẬN XIN PHÉP BÁO CÁO TRỄ",
        "ĐIỂM DANH MỐC",
        "BẢNG TỔNG HỢP ĐIỂM DANH",
        "TEST BOT ĐIỂM DANH"
    ]):
        print(f"[{ts}] SKIP — Tin nhắn từ Bot, bỏ qua để tránh lặp")
        return jsonify({"status": "skipped", "reason": "bot message"})

    # Chỉ xử lý khi tin nhắn được gửi vào trigger group
    if channel_id and channel_id not in TRIGGER_GROUPS:
        print(f"[{ts}] SKIP — Group {channel_id} không phải trigger group ({TRIGGER_GROUPS})")
        return jsonify({"status": "skipped", "reason": f"not trigger group: {channel_id}"})

    print(f"[{ts}] TEXT ĐẬP VÀO GROUP {channel_id}:\n{msg_text[:300]}")

    # ─── 1. KIỂM TRA & GHI NHẬN BÁO CÁO ĐIỂM DANH AM (ƯU TIÊN SỐ 1) ──────────
    try:
        from diem_danh_bot import record_submission
        sender_obj = data.get("sender") or msg_obj.get("sender") or {}
        sender_name = sender_obj.get("displayName") or sender_obj.get("name") or data.get("senderName") or ""
        sender_id = sender_obj.get("id") or str(data.get("senderId") or "")
        msg_id = str(data.get("id") or msg_obj.get("id") or time.time())

        res_dd, status_dd = record_submission(
            sender_name=sender_name,
            sender_id=sender_id,
            raw_text=msg_text,
            channel_id=channel_id,
            msg_id=msg_id
        )
        if res_dd:
            print(f"[{ts}] 🎯 [ĐIỂM DANH AM] Đã ghi nhận: {res_dd['am_name']} - Mốc {res_dd['milestone_id']} ({res_dd['status']} - Phạt: {res_dd['penalty']:,}đ)")
            
            # Phản hồi xác nhận tức thì vào Group
            status_emoji = "✅" if res_dd["status"] == "ON_TIME" else ("⚠️" if res_dd["status"] == "LATE" else "🚫")
            fine_text = f" (Phạt {res_dd['penalty']:,}đ)" if res_dd["penalty"] > 0 else " (Đúng hạn - 0đ)"
            hubs_line = f"🏢 <b>Bưu cục:</b> {', '.join(res_dd['hubs'])}\n" if res_dd.get("hubs") else ""
            m5_line = f"{res_dd['m5_progress_note']}\n" if res_dd.get("m5_progress_note") else ""
            confirm_msg = (
                f"{status_emoji} <b>XÁC NHẬN ĐÃ GHI NHẬN BÁO CÁO</b>\n"
                f"👤 <b>Khu vực AM:</b> {res_dd['am_name']}\n"
                f"{hubs_line}"
                f"📋 <b>Loại báo cáo:</b> Mốc {res_dd['milestone_id']} ({res_dd['milestone_name']})\n"
                f"⏰ <b>Thời gian nộp:</b> {res_dd['submit_time']} — <i>{res_dd['note']}{fine_text}</i>\n"
                f"{m5_line}"
            ).strip()
            from diem_danh_bot import send_gtalk_message, load_config
            cfg_am = load_config()
            group_b_id = str(cfg_am.get("gtalk", {}).get("channel_id_group_b") or "2097270568973508608")
            group_a_id = str(cfg_am.get("gtalk", {}).get("channel_id_group_a") or "2097277790030348288")

            # Mốc 5 (BC Điểm nóng) luôn xác nhận vào Group B; Mốc 1-4 xác nhận vào Group gửi đến / Group A
            if res_dd.get("milestone_id") == 5:
                reply_channel = group_b_id
            else:
                reply_channel = channel_id or group_a_id

            send_gtalk_message(confirm_msg, reply_channel)

            # Tự động đồng bộ thời gian thực lên Google Sheet
            try:
                from sync_attendance_sheets import sync_daily_to_sheet
                threading.Thread(target=sync_daily_to_sheet, args=(get_vn_now().date(),), daemon=True).start()
            except Exception:
                pass

            return jsonify({"status": "recorded_attendance", "data": res_dd})
        else:
            print(f"[{ts}] [ĐIỂM DANH] Bỏ qua: {status_dd}")
            if status_dd not in ("Không phải mẫu báo cáo 1-5", "Không xác định được AM"):
                return jsonify({"status": "skipped", "reason": status_dd})
    except Exception as e_dd:
        import traceback
        traceback.print_exc()
        print(f"[{ts}] [ĐIỂM DANH] Lỗi: {e_dd}")
        return jsonify({"status": "error", "error": str(e_dd), "trace": traceback.format_exc()})

    # ─── 2. KIỂM TRA TIN NHẮN XIN PHÉP TRỄ / OFF PHÉP / MIỄN BÁO CÁO ────────
    try:
        from diem_danh_bot import detect_excuse_request, send_gtalk_message, load_config
        sender_obj = data.get("sender") or msg_obj.get("sender") or {}
        sender_name = sender_obj.get("displayName") or sender_obj.get("name") or data.get("senderName") or ""
        excuse_info = detect_excuse_request(msg_text, sender_name, channel_id=channel_id)
        if excuse_info:
            print(f"[{ts}] 📝 [XIN PHÉP TRỄ / MIỄN] Đã ghi nhận: {excuse_info['am']['full_name']} - {excuse_info['scope_label']}")
            cfg_am = load_config()
            group_b_id = str(cfg_am.get("gtalk", {}).get("channel_id_group_b") or "2097270568973508608")
            group_a_id = str(cfg_am.get("gtalk", {}).get("channel_id_group_a") or "2097277790030348288")
            reply_ch = group_b_id if (5 in excuse_info.get("milestones", []) or str(channel_id) == group_b_id) else (channel_id or group_a_id)
            send_gtalk_message(excuse_info["reply_msg"], reply_ch)
            try:
                from sync_attendance_sheets import sync_daily_to_sheet
                threading.Thread(target=sync_daily_to_sheet, args=(get_vn_now().date(),), daemon=True).start()
            except Exception:
                pass
            return jsonify({"status": "recorded_excuse", "data": excuse_info})
    except Exception as e_ex:
        print(f"[{ts}] [XIN PHÉP] Lỗi: {e_ex}")

    # Parse danh sách AM (dành cho Cảnh báo Ticket GLT)
    header, am_items = parse_ticket_message(msg_text)
    if not am_items:
        print(f"[{ts}] SKIP — Không tìm thấy dòng AM nào")
        return jsonify({"status": "skipped", "reason": "no AM data"})

    print(f"[{ts}] PARSED — Header: {header[:60]} | AMs tìm thấy: {len(am_items)}")

    # Load mapping từ Google Sheet (để luôn lấy dữ liệu mới nhất)
    try:
        am_map, token = load_am_mapping()
    except Exception as e:
        print(f"[{ts}] ERROR loading sheet: {e}")
        return jsonify({"status": "error", "reason": str(e)}), 500

    # Bắn cảnh báo cho từng AM
    results = []
    success_count = 0
    for item in am_items:
        am_name = item["am_name"]
        raw_line = item["raw_line"]
        am_norm = norm(am_name)

        # Tìm exact match
        am_info = am_map.get(am_norm)

        # Fuzzy match nếu không tìm thấy
        if not am_info:
            for key, val in am_map.items():
                if am_norm == key or am_norm in key or key in am_norm:
                    am_info = val
                    break

        if not am_info:
            print(f"[{ts}] ⚠️ BỎ QUA AM '{am_name}' — Chưa có Group ID trong Google Sheet")
            results.append({"am": am_name, "success": False, "reason": "Không có Group ID"})
            continue

        group_id = am_info["group_id"]
        display_name = am_info["display_name"]
        msg = build_message(header, raw_line)

        ok, err = send_gtalk(group_id, msg, token)
        if ok:
            print(f"[{ts}] ✅ BẮN THÀNH CÔNG → {display_name} (Group: {group_id})")
            results.append({"am": display_name, "success": True, "group_id": group_id})
            success_count += 1
        else:
            print(f"[{ts}] ❌ Bắn lỗi → {display_name}: {err}")
            results.append({"am": display_name, "success": False, "reason": err})

        time.sleep(0.5)

    print(f"[{ts}] HOÀN TẤT — Bắn thành công {success_count}/{len(am_items)} AM")
    return jsonify({
        "status": "ok",
        "total": len(am_items),
        "success": success_count,
        "details": results,
    })


# ─── NGROK SETUP ────────────────────────────────────────────

def get_ngrok_token():
    """Đọc token từ file hoặc hỏi người dùng."""
    if os.path.exists(NGROK_TOKEN_FILE):
        with open(NGROK_TOKEN_FILE, "r") as f:
            token = f.read().strip()
            if token:
                return token

    print("\n" + "=" * 60)
    print("🔑 CẦN NGROK AUTHTOKEN")
    print("   1. Vào https://dashboard.ngrok.com/authtokens")
    print("   2. Đăng nhập (miễn phí) → copy token")
    print("=" * 60)
    token = input("Paste ngrok authtoken vào đây: ").strip()
    if token:
        with open(NGROK_TOKEN_FILE, "w") as f:
            f.write(token)
        print(f"✅ Đã lưu token vào {NGROK_TOKEN_FILE}")
    return token


def start_ngrok(port: int):
    """Khởi động ngrok tunnel và trả về public URL."""
    from pyngrok import ngrok, conf

    token = get_ngrok_token()
    if token:
        conf.get_default().auth_token = token

    try:
        ngrok.kill()
        time.sleep(1)
    except Exception:
        pass

    try:
        tunnel = ngrok.connect(port, "http", domain=STATIC_DOMAIN)
        print(f"✅ Đã kết nối với Static Domain: {STATIC_DOMAIN}")
    except Exception as ex:
        print(f"⚠️ Kết nối static domain {STATIC_DOMAIN} thất bại: {ex}. Thử kết nối mặc định...")
        tunnel = ngrok.connect(port, "http")

    public_url = tunnel.public_url
    if public_url.startswith("http://"):
        public_url = public_url.replace("http://", "https://", 1)
    return public_url


# ─── MAIN ────────────────────────────────────────────────────

def main():
    print("=" * 65)
    print("🚀 TICKET WEBHOOK SERVER — VÙNG NTB")
    print(f"📌 TRIGGER GROUPS (Group nhận tin để bắn đi): {TRIGGER_GROUPS}")
    print(f"🔑 BOT TOKEN: {OA_TOKEN[:20]}...{OA_TOKEN[-5:]}")
    print(f"📊 GOOGLE SHEET ID: {SPREADSHEET_ID} (Tab: {SHEET_TAB})")
    print("=" * 65)

    # Thử khởi động ngrok
    public_url = None
    try:
        from pyngrok import ngrok
        public_url = start_ngrok(WEBHOOK_PORT)
        webhook_url = f"{public_url}/webhook"
        print("\n" + "=" * 65)
        print("✅ NGROK TUNNEL ĐÃ SẴN SÀNG!")
        print(f"   👉 Webhook URL: {webhook_url}")
        print("=" * 65)
        print("\nCách hoạt động:")
        print(f"  1. Webhook GTalk đã trỏ đúng: {webhook_url}")
        print(f"  2. Khi bạn fw tin nhắn tồn phiếu vào Group {TRIGGER_GROUPS}:")
        print("     → Server tự động parse và bắn cảnh báo riêng cho từng AM theo Group ID ở Cột C/D Google Sheet!")
        print("\n⌛ Server đang lắng nghe webhook... (Ctrl+C để dừng)\n")
    except Exception as e:
        print(f"⚠️  Không khởi động được ngrok: {e}")
        print(f"   Chạy thủ công: ngrok http {WEBHOOK_PORT}")
        print(f"   Sau đó dùng URL https://xxxx.ngrok.app/webhook\n")

    # Chạy Attendance Scheduler & Khôi phục DB từ Google Sheet (phòng ngừa restart Render)
    try:
        from diem_danh_bot import init_db, start_scheduler
        init_db()
        try:
            from sync_attendance_sheets import restore_db_from_sheet
            restore_db_from_sheet()
        except Exception as e_res:
            print(f"⚠️ Restore from sheet error: {e_res}")
        start_scheduler()
    except Exception as e:
        print(f"⚠️ Attendance scheduler error: {e}")

    # Chạy Self-Ping thread giữ Render luôn thức (chống Free Tier ngủ sau 15 phút)
    def keep_alive():
        render_url = "https://auto-report-shol.onrender.com/health"
        while True:
            time.sleep(600)  # Ping mỗi 10 phút
            try:
                requests.get(render_url, timeout=15)
            except Exception:
                pass
    threading.Thread(target=keep_alive, daemon=True).start()

    # Chạy Flask
    app.run(host="0.0.0.0", port=WEBHOOK_PORT, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
