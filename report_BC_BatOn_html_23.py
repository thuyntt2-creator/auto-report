# -*- coding: utf-8 -*-
"""
Script: report_BC_BatOn_html.py
Đọc dữ liệu từ Google Sheet (NTB + Cocau), build báo cáo HTML/CSS đẹp,
chụp ảnh bằng Playwright, gửi ảnh lên Telegram.

CÀI ĐẶT (chạy 1 lần):
    pip install gspread oauth2client pillow requests pytz playwright --break-system-packages
    playwright install chromium

Nếu lệnh trên báo lỗi --break-system-packages không hợp lệ trên Windows, bỏ flag đó:
    pip install gspread oauth2client requests pytz playwright
    playwright install chromium
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
import os
from playwright.sync_api import sync_playwright
import requests
import urllib3
import sys

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

from dotenv import load_dotenv
env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv(override=True)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============ CONFIG ============
SERVICE_ACCOUNT_CANDIDATES = [
    r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
    r"C:\Users\lap4all\Downloads\credentials.json",
    r"C:\Users\lap4all\Downloads\service_account.json",
    r"C:\Users\lap4all\Desktop\credentials.json",
    "credentials.json",
    "service_account.json",
]


def find_service_account_file():
    for p in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(p):
            return p
    raise FileNotFoundError(
        "Không tìm thấy file credentials.json. Thêm đúng path vào SERVICE_ACCOUNT_CANDIDATES."
    )


SHEET_ID = "1lmQv8KwHJzDFs_RMz64ydu4SOmG3M1YAzILNFGtzFec"
SHEET_NTB_NAME = "NTB"
SHEET_COCAU_NAME = "Cocau"

TELEGRAM_TOKEN = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
TELEGRAM_CHAT_ID = "-5058464865"

SHEET_LINK = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?gid=250113221"

OUTPUT_IMAGE = "bc_baton_report.png"
OUTPUT_HTML = "bc_baton_report.html"


# ============ GOOGLE SHEET HELPERS ============
def get_sheet_data():
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(find_service_account_file(), scope)
    client = gspread.authorize(creds)

    sh = client.open_by_key(SHEET_ID)
    ntb_ws = sh.worksheet(SHEET_NTB_NAME)
    cocau_ws = sh.worksheet(SHEET_COCAU_NAME)

    ntb_data = ntb_ws.get_all_values()
    cocau_data = cocau_ws.get_all_values()
    return ntb_data, cocau_data


def build_maps(cocau_data):
    """Trả về map: ten_bc_rut_gon -> AM"""
    am_map = {}
    for row in cocau_data[1:]:
        if len(row) < 4:
            continue
        bc_name = (row[1] or "").strip()   # CỘT B = Bưu cục
        am = (row[3] or "").strip()        # CỘT D = AM
        if bc_name:
            am_map[bc_name] = am
    return am_map


def match_am(kho_giao_name, am_map):
    for key, am in am_map.items():
        if key and (key in kho_giao_name or kho_giao_name in key):
            return am
    return "Không rõ AM"


def process_data(ntb_data, am_map):
    top_bat_on = []
    chuan_bi_nhay = []
    canh_bao = []
    am_bat_on_count = {}
    am_bat_on_list = {}
    am_chuan_bi_nhay = {}
    clear_today = []  # list of (bc_full, am, target_clear)

    # data thật bắt đầu từ dòng 5 -> index 4
    for row in ntb_data[4:]:
        if len(row) < 21:
            continue
        bc_full = (row[4] or "").strip()    # cột E = kho_giao_name
        status = (row[20] or "").strip()    # cột U = Trạng thái

        if not bc_full or not status:
            continue

        am = match_am(bc_full, am_map)

        if "Bất ổn" in status:
            top_bat_on.append(bc_full)
            am_bat_on_count[am] = am_bat_on_count.get(am, 0) + 1
            am_bat_on_list.setdefault(am, []).append(bc_full)
        elif "Chuẩn bị nhảy nhóm" in status:
            chuan_bi_nhay.append(bc_full)
            am_chuan_bi_nhay.setdefault(am, []).append(bc_full)
        elif "Cấp cảnh báo" in status:
            canh_bao.append(bc_full)

        # ---- Tính số đơn cần clear hôm nay (chỉ áp dụng Bất ổn / Chuẩn bị nhảy nhóm) ----
        if "Bất ổn" in status or "Chuẩn bị nhảy nhóm" in status:
            def to_num(v):
                s = str(v or "0").strip().replace(",", "").replace("%", "")
                try:
                    return float(s)
                except ValueError:
                    return 0.0

            bl_lm = to_num(row[6])          # cột G = BL LM
            bl_lm_5ngay = to_num(row[7])    # cột H = BL LM >5 ngày
            bl_ktc_tinh = to_num(row[10])   # cột K = BL KTC cùng tỉnh
            tao_avg = to_num(row[13])       # cột N = tao_avg_7ngay (đơn tạo mới TB/ngày)
            gtc_avg = to_num(row[15])       # cột P = gtc_avg_7ngay

            ly_do = (row[19] or "").strip() if len(row) > 19 else ""

            target_lm = 0
            target_ktc = 0

            if "Tồn LM" in ly_do or "Tồn Aging" in ly_do:
                target_aging = bl_lm_5ngay - 0.10 * bl_lm
                target_ton_lm = bl_lm + tao_avg - 1.8 * gtc_avg
                target_lm = max(0, round(max(target_aging, target_ton_lm)))

            if "Tồn KTC" in ly_do:
                target_ktc = max(0, round(bl_ktc_tinh + tao_avg - 1.8 * gtc_avg))

            if target_lm > 0 or target_ktc > 0:
                clear_today.append((bc_full, am, target_lm, target_ktc, ly_do))

    return {
        "top_bat_on": top_bat_on,
        "chuan_bi_nhay": chuan_bi_nhay,
        "canh_bao": canh_bao,
        "am_bat_on_count": am_bat_on_count,
        "am_bat_on_list": am_bat_on_list,
        "am_chuan_bi_nhay": am_chuan_bi_nhay,
        "clear_today": clear_today,
    }


# ============ HTML BUILDER ============
def esc(s):
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_html(data, today_str):
    total_bc = len(data["top_bat_on"]) + len(data["chuan_bi_nhay"]) + len(data["canh_bao"])

    # ---- AM stats table (gộp tên BC theo từng AM) ----
    am_stats_html = ""
    if data["am_bat_on_count"]:
        # map BC -> tổng số đơn cần clear (LM + KTC)
        clear_map = {bc: (t_lm + t_ktc) for bc, am, t_lm, t_ktc, ly_do in data["clear_today"]}

        sorted_am = sorted(data["am_bat_on_count"].items(), key=lambda x: -x[1])
        rows = ""
        idx = 1
        for am_idx, (am, count) in enumerate(sorted_am):
            color_class = f"am-color-{am_idx % 7}"
            bcs = data["am_bat_on_list"].get(am, [])
            for j, bc in enumerate(bcs):
                clear_val = clear_map.get(bc, 0)
                if j == 0:
                    rows += f"""<tr class='{color_class}'>
                      <td class='idx'>{idx}</td>
                      <td rowspan='{len(bcs)}'>{esc(am)}</td>
                      <td rowspan='{len(bcs)}' class='num red'>{count}</td>
                      <td>{esc(bc)}</td>
                      <td class='num orange'>{clear_val}</td>
                    </tr>"""
                else:
                    rows += f"""<tr class='{color_class}'>
                      <td class='idx'>{idx}</td>
                      <td>{esc(bc)}</td>
                      <td class='num orange'>{clear_val}</td>
                    </tr>"""
                idx += 1
        am_stats_html = f"""
        <div class="section">
          <div class="section-title red"><span class="badge">{len(sorted_am)} AM</span>Nhóm bưu cục bất ổn theo AM</div>
          <table class="data-table">
            <tr><th class="center"></th><th>AM</th><th class="center">Số BC</th><th>Bưu cục</th><th class="center">Cần xử lý</th></tr>
            {rows}
          </table>
        </div>"""

    # ---- AM cần lưu ý table ----
    am_warn_html = ""
    if data["am_chuan_bi_nhay"]:
        rows = ""
        idx = 1
        for am, bcs in data["am_chuan_bi_nhay"].items():
            for bc in bcs:
                rows += f"<tr><td class='idx'>{idx}</td><td>{esc(am)}</td><td>{esc(bc)}</td></tr>"
                idx += 1
        am_warn_html = f"""
        <div class="section">
          <div class="section-title orange"><span class="badge">{idx-1} BC</span>Nhóm sắp nhảy nhóm bất ổn - AM cần lưu ý</div>
          <table class="data-table">
            <tr><th class="center"></th><th>AM</th><th>Bưu cục</th></tr>
            {rows}
          </table>
        </div>"""

    # ---- List table builder ----
    def build_list_section(title, items, color_class):
        if not items:
            return ""
        rows = ""
        for i, bc in enumerate(items):
            rows += f"<tr><td class='idx'>{i+1}</td><td>{esc(bc)}</td></tr>"
        return f"""
        <div class="section">
          <div class="section-title {color_class}"><span class="badge">{len(items)} BC</span>{title}</div>
          <table class="data-table">
            <tr><th class="center"></th><th>Tên bưu cục</th></tr>
            {rows}
          </table>
        </div>"""

    top_bat_on_html = ""
    chuan_bi_nhay_html = ""

    clear_today_html = ""
    canh_bao_html = build_list_section("Nhóm cảnh báo - có nguy cơ bất ổn", data["canh_bao"], "yellow")

    html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<style>
  * {{
    box-sizing: border-box;
    font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
  }}
  body {{
    margin: 0;
    padding: 0;
    background: #ffffff;
  }}
  .container {{
    width: 900px;
    padding: 36px 40px 44px;
    color: #1a1a1a;
  }}
  .header {{
    margin-bottom: 28px;
  }}
  .header h1 {{
    font-size: 26px;
    font-weight: 600;
    margin: 0 0 6px 0;
    color: #1a1a1a;
  }}
  .header .subtitle {{
    color: #8a8a8a;
    font-size: 15px;
  }}

  .summary-bar {{
    display: flex;
    gap: 14px;
    margin-bottom: 32px;
  }}
  .summary-pill {{
    flex: 1;
    border-radius: 10px;
    padding: 16px 18px;
  }}
  .summary-pill .label {{
    font-size: 14px;
    margin-bottom: 8px;
    font-weight: 500;
  }}
  .summary-pill .value {{
    font-size: 32px;
    font-weight: 700;
  }}
  .summary-pill.red {{ background: #FCEBEB; }}
  .summary-pill.red .label {{ color: #A32D2D; }}
  .summary-pill.red .value {{ color: #A32D2D; }}
  .summary-pill.orange {{ background: #FAEEDA; }}
  .summary-pill.orange .label {{ color: #854F0B; }}
  .summary-pill.orange .value {{ color: #854F0B; }}
  .summary-pill.yellow {{ background: #FAEEDA; }}
  .summary-pill.yellow .label {{ color: #854F0B; }}
  .summary-pill.yellow .value {{ color: #854F0B; }}

  .section {{
    margin-bottom: 30px;
  }}
  .section-title {{
    font-size: 17px;
    font-weight: 600;
    margin-bottom: 12px;
    display: flex;
    align-items: center;
    gap: 10px;
    color: #1a1a1a;
  }}
  .section-title .badge {{
    font-size: 13px;
    font-weight: 600;
    padding: 4px 10px;
    border-radius: 6px;
  }}
  .section-title.red .badge {{ background: #A32D2D; color: #ffffff; }}
  .section-title.orange .badge {{ background: #D9822B; color: #ffffff; }}
  .section-title.yellow .badge {{ background: #C79A1E; color: #ffffff; }}

  .data-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 17px;
    border: 1px solid #c8c8c8;
    margin-bottom: 4px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
  }}
  .data-table th {{
    background: #404040;
    color: #ffffff;
    text-align: left;
    font-weight: 600;
    padding: 11px 16px;
    font-size: 14px;
    border: 1px solid #595959;
  }}
  .data-table th.center {{ text-align: center; }}
  .data-table td {{
    padding: 9px 14px;
    border: 1px solid #e0e0e0;
    color: #2a2a2a;
  }}
  .data-table tr:nth-child(even) td {{
    background: #FAF6F0;
  }}

  /* Màu phân nhóm theo AM */
  .data-table tr.am-color-0 td {{ background: #FDECEC; }}
  .data-table tr.am-color-1 td {{ background: #FFF4E0; }}
  .data-table tr.am-color-2 td {{ background: #FFF9E0; }}
  .data-table tr.am-color-3 td {{ background: #E8F5E9; }}
  .data-table tr.am-color-4 td {{ background: #E3F2FD; }}
  .data-table tr.am-color-5 td {{ background: #F3E5F5; }}
  .data-table tr.am-color-6 td {{ background: #FCE4EC; }}
  .data-table td.center {{
    text-align: center;
  }}
  .data-table td.num {{
    font-weight: 700;
    text-align: center;
    font-size: 18px;
    color: #c4c4c4;
  }}
  .data-table td.red {{ color: #A32D2D; }}
  .data-table td.num.orange {{ color: #D9822B; }}
  .inline-clear {{
    color: #D9822B;
    font-weight: 600;
    font-size: 14px;
  }}
  .data-table td.idx {{
    color: #999999;
    width: 32px;
    text-align: center;
  }}
  .data-table td.left-text {{
    text-align: left;
    line-height: 1.7;
  }}
  .data-table td[rowspan] {{
    vertical-align: middle;
    border-bottom: 2px solid #c8c8c8;
  }}

  .footer {{
    margin-top: 24px;
    padding-top: 16px;
    border-top: 1px solid #ececec;
    font-size: 13px;
    color: #8a8a8a;
    text-align: center;
  }}
</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Danh sách bưu cục cảnh báo</h1>
      <div class="subtitle">{esc(today_str)} · Tổng {total_bc} bưu cục</div>
    </div>

    <div class="summary-bar">
      <div class="summary-pill red">
        <div class="label">Bất ổn</div>
        <div class="value">{len(data['top_bat_on'])}</div>
      </div>
      <div class="summary-pill orange">
        <div class="label">Sắp nhảy nhóm</div>
        <div class="value">{len(data['chuan_bi_nhay'])}</div>
      </div>
      <div class="summary-pill yellow">
        <div class="label">Cấp cảnh báo</div>
        <div class="value">{len(data['canh_bao'])}</div>
      </div>
    </div>

    {clear_today_html}
    {am_stats_html}
    {am_warn_html}
    {top_bat_on_html}
    {chuan_bi_nhay_html}
    {canh_bao_html}

    <div class="footer">NTB - Bưu cục bất ổn · Cập nhật {esc(today_str)}</div>
  </div>
</body>
</html>"""
    return html


# ============ RENDER TO IMAGE ============
def render_image(html_content, html_path, image_path):
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 900, "height": 800})
        page.goto(f"file:///{os.path.abspath(html_path)}")
        page.wait_for_timeout(300)
        page.screenshot(path=image_path, full_page=True)
        browser.close()


# ============ TELEGRAM SENDER ============
def send_photo_telegram(image_path, caption=""):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendPhoto"
    with open(image_path, "rb") as f:
        files = {"photo": f}
        payload = {
            "chat_id": TELEGRAM_CHAT_ID,
            "caption": caption,
            "parse_mode": "HTML",
        }
        resp = requests.post(url, data=payload, files=files)
    result = resp.json()
    if result.get("ok"):
        print("✅ Đã gửi ảnh lên Telegram thành công.")
    else:
        print(f"❌ Lỗi gửi Telegram: {result}")
    return result


# ============ GTALK SENDER ============
def send_photo_gtalk(image_path, caption=""):
    oa_token = os.environ.get("GTALK_OA_TOKEN")
    channel_id = os.environ.get("GTALK_CHANNEL_ID")
    if not oa_token or not channel_id:
        print("⚠️ Không tìm thấy GTALK_OA_TOKEN hoặc GTALK_CHANNEL_ID trong .env. Bỏ qua gửi GTalk.")
        return False

    print("📡 Đang gửi ảnh báo cáo sang GTalk...")
    try:
        from PIL import Image
        img = Image.open(image_path)
        width, height = img.size
        file_size = os.path.getsize(image_path)
    except Exception as e:
        print(f"❌ Lỗi đọc ảnh {image_path}: {e}")
        return False

    # Step 1: Initiate Upload
    initiate_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
    payload_init = {
        "ChannelId": channel_id,
        "FileName": os.path.basename(image_path),
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": f'{{"width": {width}, "height": {height}}}',
        "oaToken": oa_token
    }
    headers = {"Content-Type": "application/json"}
    try:
        res_init = requests.post(initiate_url, json=payload_init, headers=headers, timeout=20, verify=False)
        if res_init.status_code != 200:
            print(f"❌ Lỗi initiate upload HTTP {res_init.status_code}: {res_init.text}")
            return False
        res_data = res_init.json()
        if res_data.get("errorCode") != "success":
            print(f"❌ Lỗi initiate upload API: {res_data.get('error')}")
            return False
        
        presigned_url = res_data["data"]["PresignedURL"]
        upload_id = res_data["data"]["UploadId"]
    except Exception as e:
        print(f"❌ Lỗi kết nối khi initiate upload GTalk: {e}")
        return False

    # Step 2: Upload to S3
    try:
        with open(image_path, "rb") as f:
            headers_put = {"Content-Type": "image/png"}
            res_put = requests.put(presigned_url, data=f, headers=headers_put, timeout=60, verify=False)
            if res_put.status_code != 200:
                print(f"❌ Lỗi PUT lên S3 HTTP {res_put.status_code}: {res_put.text}")
                return False
    except Exception as e:
        print(f"❌ Lỗi upload file lên S3 GTalk: {e}")
        return False

    # Step 3: Complete Upload
    complete_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
    payload_complete = {
        "oaToken": oa_token,
        "UploadId": upload_id
    }
    try:
        res_comp = requests.post(complete_url, json=payload_complete, headers=headers, timeout=20, verify=False)
        if res_comp.status_code != 200:
            print(f"❌ Lỗi complete upload HTTP {res_comp.status_code}: {res_comp.text}")
            return False
        res_data_comp = res_comp.json()
        if res_data_comp.get("errorCode") != "success":
            print(f"❌ Lỗi complete upload API: {res_data_comp.get('error')}")
            return False
        file_id = res_data_comp["data"]["Id"]
    except Exception as e:
        print(f"❌ Lỗi kết nối khi complete upload GTalk: {e}")
        return False

    # Step 4: Send Message
    import time
    send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
    client_msg_id = str(int(time.time() * 1000))
    payload_send = {
        "channelId": channel_id,
        "clientMsgId": client_msg_id,
        "content": {
            "parseMode": "HTML",
            "attachment": {
                "caption": caption,
                "items": [
                    {
                        "image": {
                            "fileId": file_id,
                            "width": width,
                            "height": height
                        }
                    }
                ]
            }
        },
        "oaToken": oa_token
    }
    try:
        res_send = requests.post(send_url, json=payload_send, headers=headers, timeout=20, verify=False)
        if res_send.status_code == 200:
            res_data_send = res_send.json()
            if res_data_send.get("errorCode") == "success":
                print("✅ Đã gửi ảnh báo cáo sang GTalk thành công!")
                return True
            else:
                print(f"❌ Lỗi gửi tin nhắn GTalk API: {res_data_send.get('error')}")
        else:
            print(f"❌ Lỗi HTTP {res_send.status_code}: {res_send.text}")
    except Exception as e:
        print(f"❌ Lỗi kết nối khi gửi tin nhắn GTalk: {e}")
    return False


# ============ MAIN ============
def main():
    print("Đang đọc dữ liệu từ Google Sheet...")
    ntb_data, cocau_data = get_sheet_data()

    print("Đang xử lý dữ liệu...")
    am_map = build_maps(cocau_data)
    data = process_data(ntb_data, am_map)

    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    today_str = datetime.now(tz).strftime("%d/%m/%Y %H:%M")

    print("Đang build HTML...")
    html_content = build_html(data, today_str)

    print("Đang chụp ảnh báo cáo...")
    html_path = os.path.join(os.getcwd(), OUTPUT_HTML)
    image_path = os.path.join(os.getcwd(), OUTPUT_IMAGE)
    render_image(html_content, html_path, image_path)
    print(f"✅ Đã tạo ảnh: {image_path}")

    total_bc = len(data["top_bat_on"]) + len(data["chuan_bi_nhay"]) + len(data["canh_bao"])
    date_str = datetime.now(tz).strftime("%d/%m/%Y")

    # Dynamic AM GTalk & Tele Mentions
    import json
    import unicodedata
    def normalize_name(name):
        if not name:
            return ""
        name = unicodedata.normalize('NFC', name)
        return " ".join(name.strip().lower().split())

    fallback_ids = {
        "Võ Tấn Lợi": "1959937322596192256",
        "Trần Công Hậu": "1982691000129204224",
        "Nguyễn Ngọc Khánh": "1964615676064722944",
        "Trần Văn Phước": "1999372176819601408",
        "Trần Thị Nhung": "1982691054671896576",
        "Nguyễn Duy Long": "1982690554635362304",
        "Huỳnh Tấn Hiền": "1944763878326951936",
        "Phạm Bá Thành Công": "1982689297459884032",
        "Nguyễn Hoàng Phi": "1980106579622596608",
        "Trầm Hữu Tiến": "2000384276208652288",
        "Huỳnh Thị Kim Chi": "1945750622928121856",
        "Lê Minh Đại": "1945048971813609472",
        "Lê Văn Trường": "2039198176261496832",
        "Hồng Bích Nga": "1982689048301449216",
        "Hồng Bích Nga": "1982689048301449216",
        "Thái Thị Thanh Thư": "1982692716857831424",
        "Lê Thanh Nhựt": "2039198305731272704",
        "Phan Đình Duy": "1982690517352230912",
        "Nguyễn Thanh Long": "2061728677681455104",
        "Nguyễn Thanh Long": "2061728677681455104",
        "Nguyễn Lê Nguyên Vũ": "2041079293033897984",
        "Nguyễn Tống Hùng Phong": "2001976220106891264"
    }

    resolved_ids = {}
    json_path = r"C:\Users\lap4all\Downloads\resolved_gtalk_ids.json"
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                raw_ids = json.load(f)
                resolved_ids = {normalize_name(k): v for k, v in raw_ids.items()}
        except Exception:
            pass

    # Gather unstable & warning-to-change AMs
    ams_to_tag = sorted(list({am for am in (set(data["am_bat_on_count"].keys()) | set(data["am_chuan_bi_nhay"].keys())) if am and am != "Không rõ AM"}))
    
    # GTalk mentions
    am_tags = []
    for am in ams_to_tag:
        norm_am = normalize_name(am)
        user_id = resolved_ids.get(norm_am) or fallback_ids.get(am)
        if not user_id:
            for k, v in fallback_ids.items():
                if normalize_name(k) == norm_am:
                    user_id = v
                    break
        if user_id:
            am_tags.append(f'<a href="mention://{user_id}">@{am}</a>')
        else:
            am_tags.append(f'@{am}')

    am_tags_str = ", ".join(am_tags) if am_tags else "Không có"

    # Telegram text-based mentions
    am_plain_str = ", ".join([f"@{am}" for am in ams_to_tag]) if ams_to_tag else "Không có"

    caption_tele = (
        f"🌐 <b>DANH SÁCH BƯU CỤC CẢNH BÁO (LOGIC MỚI)</b>\n"
        f"📅 <b>NGÀY:</b> {date_str}\n"
        f"Tổng bưu cục: {total_bc} BC\n"
        f"🚫 Nằm TOP Bất ổn: {len(data['top_bat_on'])} BC\n"
        f"🆘 Chuẩn bị nhảy nhóm: {len(data['chuan_bi_nhay'])} BC\n"
        f"❌ Cấp cảnh báo: {len(data['canh_bao'])} BC\n\n"
        f"🔥Link: <a href=\"{SHEET_LINK}\">NTB - BƯU CỤC BẤT ỔN</a>\n"
        f"===============\n\n"
        f"👉 <b>tag AM:</b> {am_plain_str}"
    )

    caption_gtalk = (
        f"🌐 <b>DANH SÁCH BƯU CỤC CẢNH BÁO (LOGIC MỚI)</b>\n"
        f"📅 <b>NGÀY:</b> {date_str}\n"
        f"Tổng bưu cục: {total_bc} BC\n"
        f"🚫 Nằm TOP Bất ổn: {len(data['top_bat_on'])} BC\n"
        f"🆘 Chuẩn bị nhảy nhóm: {len(data['chuan_bi_nhay'])} BC\n"
        f"❌ Cấp cảnh báo: {len(data['canh_bao'])} BC\n\n"
        f"🔥Link: <a href=\"{SHEET_LINK}\">NTB - BƯU CỤC BẤT ỔN</a>\n"
        f"===============\n\n"
        f"👉 <b>tag AM:</b> {am_tags_str}"
    )

    print("Đang gửi lên Telegram...")
    send_photo_telegram(image_path, caption_tele)

    send_photo_gtalk(image_path, caption_gtalk)


if __name__ == "__main__":
    main()
