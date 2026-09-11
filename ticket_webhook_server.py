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
from datetime import datetime

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
TRIGGER_GROUPS   = ["2077278419534073856", "2095921878551764992"]   # Group nhận tin nhắn → trigger bắn sang các AM
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
        "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
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
        "time": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
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


@app.route("/webhook", methods=["POST"])
@app.route("/gtalk/webhook", methods=["POST"])
def webhook():
    ts = datetime.now().strftime("%H:%M:%S")

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
        return jsonify({"status": "skipped", "reason": "not trigger group"})

    print(f"[{ts}] TEXT ĐẬP VÀO GROUP {channel_id}:\n{msg_text[:300]}")

    # ─── KIỂM TRA TIN NHẮN XIN PHÉP TRỄ / OFF PHÉP ────────
    try:
        from diem_danh_bot import detect_excuse_request, send_gtalk_message
        sender_obj = data.get("sender") or msg_obj.get("sender") or {}
        sender_name = sender_obj.get("displayName") or sender_obj.get("name") or data.get("senderName") or ""
        excuse_info = detect_excuse_request(msg_text, sender_name)
        if excuse_info:
            print(f"[{ts}] 📝 [XIN PHÉP TRỄ] Đã ghi nhận: {excuse_info['am']['full_name']} - {excuse_info['scope_label']}")
            send_gtalk_message(excuse_info["reply_msg"], channel_id)
            try:
                from sync_attendance_sheets import sync_daily_to_sheet
                threading.Thread(target=sync_daily_to_sheet, args=(datetime.now().date(),), daemon=True).start()
            except Exception:
                pass
            return jsonify({"status": "recorded_excuse", "data": excuse_info})
    except Exception as e_ex:
        print(f"[{ts}] [XIN PHÉP] Lỗi: {e_ex}")

    # ─── KIỂM TRA & GHI NHẬN BÁO CÁO ĐIỂM DANH AM ──────────
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
            confirm_msg = (
                f"{status_emoji} <b>XÁC NHẬN ĐÃ GHI NHẬN BÁO CÁO</b>\n"
                f"👤 <b>Khu vực AM:</b> {res_dd['am_name']}\n"
                f"{hubs_line}"
                f"📋 <b>Loại báo cáo:</b> Mốc {res_dd['milestone_id']} ({res_dd['milestone_name']})\n"
                f"⏰ <b>Thời gian nộp:</b> {res_dd['submit_time']} — <i>{res_dd['note']}{fine_text}</i>"
            )
            from diem_danh_bot import send_gtalk_message
            send_gtalk_message(confirm_msg, channel_id)

            # Tự động đồng bộ thời gian thực lên Google Sheet
            try:
                from sync_attendance_sheets import sync_daily_to_sheet
                threading.Thread(target=sync_daily_to_sheet, args=(datetime.now().date(),), daemon=True).start()
            except Exception:
                pass

            return jsonify({"status": "recorded_attendance", "data": res_dd})
    except Exception as e_dd:
        print(f"[{ts}] [ĐIỂM DANH] Lỗi: {e_dd}")

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

    # Chạy Attendance Scheduler
    try:
        from diem_danh_bot import init_db, start_scheduler
        init_db()
        start_scheduler()
    except Exception as e:
        print(f"⚠️ Attendance scheduler error: {e}")

    # Chạy Flask
    app.run(host="0.0.0.0", port=WEBHOOK_PORT, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
